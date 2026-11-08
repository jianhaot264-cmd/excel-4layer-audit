# -*- coding: utf-8 -*-
"""
发布前全检：一条命令确认「这个技能可以发了」。

跑完输出 8 组检查 + 总结论。任一 FAIL 则退出码非 0。

用法
----
    python scripts/release_check.py                  # 静态 + 自检 + 打包 + dry-run
    python scripts/release_check.py --fast          # 跳过工具链自检（最快）
    python scripts/release_check.py --with-experiment   # 额外复现边界实验（需 Excel/WPS）
    python scripts/release_check.py --out D:/x.zip  # 指定发布包输出位置

检查项
------
  1 换行符一致性      全仓 CRLF 必须为 0（Windows 静默改写的高发点）
  2 无硬编码用户路径   scripts/ 里不得出现 C:/Users/<具体用户名> 这类机器特定值
  3 脚本语法          全部 .py 必须能编译
  4 frontmatter       本机工作版：双校验器都 valid
  5 发布包 frontmatter 必须含 SkillHub 5 必需字段
  6 工具链自检        selftest.py 全绿（--fast 跳过）
  7 边界实验复现      make_traps 生成的陷阱能被 L2 抓到（--with-experiment）
  8 发布包 dry-run    SkillHub 官方 CLI 本地预检通过

退出码：0 = 全部通过，可发布；1 = 有 FAIL；2 = 缺少前置（CLI / requests）
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL_DIR = HERE.parent
SKILLHUB_KEYS = ("slug", "displayName", "version", "summary", "license")

FAILS = []
WARNS = []


def ok(msg):
    print("  ✔ %s" % msg)


def fail(msg):
    FAILS.append(msg)
    print("  ✘ FAIL  %s" % msg)


def warn(msg):
    WARNS.append(msg)
    print("  ▲ WARN  %s" % msg)


def head(t):
    print("\n== %s ==" % t)


def find_python(need_requests=True):
    """找一个依赖齐备的 Python。

    ⚠️ 必须同时校验 openpyxl：selftest.py 缺 openpyxl 时会走 `except ImportError`
       静默跳过一批用例，只报"1 个失败"——看起来像坏了，其实是**覆盖不足**。
    同理缺 python-docx 会跳过 docx 用例。所以这里一次把硬依赖都验上。
    """
    cands = ["python3", "python", "py"]
    for base in (Path.home() / "AppData/Local/Programs/Python", Path("/c/Users")):
        if base.is_dir():
            try:
                for d in sorted(base.glob("Python3*")):
                    cands.append(str(d / "python.exe"))
            except OSError:
                pass
    probe = "import requests, openpyxl" if need_requests else "pass"
    for c in cands:
        exe = c if os.path.isabs(c) else shutil.which(c)
        if not exe or not Path(exe).exists():
            continue
        try:
            if subprocess.run([exe, "-c", probe], capture_output=True, timeout=20).returncode == 0:
                return exe
        except Exception:
            continue
    return ""


# ---------- 1 换行符 ----------
def check_eol():
    head("1 换行符一致性")
    crlf = lf = 0
    offenders = []
    for p in sorted(SKILL_DIR.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in (".md", ".py", ".json", ".sh"):
            continue
        d = p.read_bytes()
        if b"\x00" in d:
            continue
        c = d.count(b"\r\n")
        l = d.count(b"\n") - c
        crlf += c
        lf += l
        if c:
            offenders.append("%s (%d)" % (p.relative_to(SKILL_DIR), c))
    if crlf:
        fail("有 %d 处 CRLF（Windows 静默改写高发点）：%s" % (crlf, "; ".join(offenders[:5])))
    else:
        ok("全仓 LF 一致（%d 处 LF，0 处 CRLF）" % lf)


# ---------- 2 硬编码用户路径 ----------
# 匹配 C:/Users/<真实用户名>… ；排除 <占位符> 形式（文档示例不算）
USER_PATH_RE = re.compile(r"[A-Za-z]:[/\\]Users[/\\](?!<)([^/\\\s\"'<>]+)")


def check_no_hardcoded_paths():
    head("2 无硬编码用户路径（脚本要能在别人机器上跑）")
    bad = []
    for p in sorted((SKILL_DIR / "scripts").glob("*.py")):
        for i, ln in enumerate(p.read_text(encoding="utf-8", errors="ignore").split("\n"), 1):
            for m in USER_PATH_RE.finditer(ln):
                seg = m.group(0)
                bad.append("%s:%d  %s" % (p.name, i, seg))
    if bad:
        fail("脚本里硬编码了具体用户目录：\n      " + "\n      ".join(bad[:8]))
    else:
        ok("scripts/ 无硬编码用户目录")


# ---------- 3 语法 ----------
def check_syntax(py):
    head("3 脚本语法")
    if not py:
        fail("找不到可用的 Python")
        return
    bad = []
    # ⚠️ 不用 `python -m py_compile`：它会显式往 scripts/__pycache__ 写 .pyc
    #    （PYTHONDONTWRITEBYTECODE 也拦不住），污染技能目录并被打进发布包（体积翻倍）。
    #    用 compile() 只做语法+字节码编译，不落盘。
    for p in sorted((SKILL_DIR / "scripts").glob("*.py")):
        code = ("import sys;p=sys.argv[1];"
                "compile(open(p,encoding='utf-8').read(),p,'exec')")
        r = subprocess.run([py, "-c", code, str(p)], capture_output=True)
        if r.returncode:
            bad.append(p.name)
    # 兜底清理（早期版本可能已留下）
    for d in SKILL_DIR.rglob("__pycache__"):
        shutil.rmtree(d, ignore_errors=True)
    if bad:
        fail("编译失败: %s" % ", ".join(bad))
    else:
        ok("%d 个脚本全部编译通过" % len(list((SKILL_DIR / "scripts").glob("*.py"))))


# ---------- 4/5 frontmatter ----------
def read_fm(path):
    t = Path(path).read_text(encoding="utf-8")
    m = re.match(r"^---\n(.*?)\n---", t, re.S)
    return (m.group(1) if m else None), t


def check_frontmatter_local():
    head("4 本机工作版 frontmatter")
    fm, _ = read_fm(SKILL_DIR / "SKILL.md")
    if not fm:
        fail("SKILL.md 缺 frontmatter")
        return
    import yaml
    d = yaml.safe_load(fm)
    for k in ("name", "description"):
        if not d.get(k):
            fail("缺 %s" % k)
    if not re.fullmatch(r"[a-z0-9-]+", str(d.get("name", ""))):
        fail("name 不是 hyphen-case: %r" % d.get("name"))
    desc = str(d.get("description", ""))
    if "<" in desc or ">" in desc:
        fail("description 含尖括号（会被拒）")
    if len(desc) > 1024:
        fail("description 超过 1024 字符（%d）" % len(desc))
    if not FAILS:
        ok("name=%s，description %d 字符，斜杠检查通过" % (d.get("name"), len(desc)))


def check_frontmatter_publish(pkg):
    head("5 发布包 frontmatter（SkillHub 5 必需字段）")
    if not pkg or not Path(pkg).is_file():
        fail("发布包不存在")
        return
    with zipfile.ZipFile(pkg) as z:
        if "SKILL.md" not in z.namelist():
            fail("发布包根目录没有 SKILL.md")
            return
        t = z.read("SKILL.md").decode("utf-8")
    m = re.match(r"^---\n(.*?)\n---", t, re.S)
    if not m:
        fail("发布包 SKILL.md 的 frontmatter 不可解析（可能是 CRLF）")
        return
    import yaml
    d = yaml.safe_load(m.group(1))
    miss = [k for k in SKILLHUB_KEYS if not d.get(k)]
    if miss:
        fail("缺字段: %s" % ", ".join(miss))
    else:
        ok("5 个必需字段齐全：slug=%s version=%s license=%s"
           % (d["slug"], d["version"], d["license"]))
    if d.get("slug") and not re.fullmatch(r"[A-Za-z0-9-]+", str(d["slug"])):
        fail("slug 只能含字母/数字/连字符")
    if d.get("version") and not re.match(r"^\d+\.\d+\.\d+$", str(d["version"])):
        fail("version 不是语义化版本")


# ---------- 6 自检 ----------
def check_selftest(py, fast):
    head("6 工具链自检")
    cmd = [py, str(SKILL_DIR / "scripts/selftest.py")]
    if not fast:
        cmd.append("--with-engines")
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=str(SKILL_DIR))
    tail = [x for x in r.stdout.split("\n") if x.strip()][-3:]
    print("      " + " | ".join(t.strip() for t in tail))
    if r.returncode == 0:
        ok("自检通过" + ("（快速模式，跳过引擎）" if fast else ""))
    else:
        fail("自检未通过（退出码 %d）" % r.returncode)


# ---------- 7 边界实验复现 ----------
def check_experiment(py):
    head("7 边界实验复现（L2 能否抓到三类陷阱）")
    tmp = Path(tempfile.mkdtemp(prefix="relcheck-exp-"))
    try:
        r = subprocess.run([py, str(SKILL_DIR / "scripts/make_traps.py"), "--out", str(tmp)],
                           capture_output=True, text=True)
        if r.returncode:
            fail("make_traps 失败：%s" % (r.stderr or r.stdout)[-200:])
            return
        cli = SKILL_DIR / "scripts/recalc_verify.py"
        for name in ("A-缓存陷阱", "B-公式错误", "C-越界引用"):
            rr = subprocess.run([py, str(cli), str(tmp / (name + ".xlsx")), "--engine", "excel"],
                                capture_output=True, text=True)
            # ⚠️ recalc_verify 的退出码语义是「是否发现不一致」：
            #    0 = 全部一致（= 没检出）  1 = 发现不一致（= 成功检出）  2 = 无引擎
            #    判反了会把"检出"当成"漏报"——这正是本脚本第一版犯的错。
            if rr.returncode == 1:
                ok("%s 被 L2 检出 ✔" % name)
            elif rr.returncode == 2:
                fail("%s 无法判定：没有可用引擎（装 Excel 或 WPS）" % name)
            else:
                fail("%s 未被 L2 检出（退出码 %d = 全部一致，说明陷阱没生效）" % (name, rr.returncode))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------- 8 dry-run ----------
def check_dryrun(py, pkg):
    head("8 SkillHub 官方 CLI dry-run")
    cli = Path.home() / ".skillhub" / "skills_store_cli.py"
    if not cli.is_file():
        warn("未装 SkillHub CLI，跳过 dry-run。安装："
             "curl -fsSL https://skillhub-1388575217.cos.ap-guangzhou.myqcloud.com/install/install.sh | bash -s -- --cli-only")
        return
    env = dict(os.environ, MSYS_NO_PATHCONV="1")
    r = subprocess.run([py, str(cli), "publish", str(pkg), "--dry-run"],
                       capture_output=True, text=True, env=env)
    out = (r.stdout or "").strip().split("\n")[-1]
    print("      " + out)
    if r.returncode == 0:
        ok("dry-run 通过")
    else:
        fail("dry-run 未通过（退出码 %d）" % r.returncode)


def main():
    ap = argparse.ArgumentParser(description="发布前全检")
    ap.add_argument("--fast", action="store_true", help="跳过工具链自检")
    ap.add_argument("--with-experiment", action="store_true", help="额外复现边界实验")
    ap.add_argument("--out", default="", help="发布包输出路径（默认写到技能目录外）")
    a = ap.parse_args()

    print("=" * 62)
    print("发布前全检 · %s" % SKILL_DIR.name)
    print("=" * 62)

    check_eol()
    check_no_hardcoded_paths()
    # 全程用同一个解释器：自检会按解释器是否装齐依赖决定跳过哪些用例，
    # 混用解释器会导致"项数变少但看起来只是个别失败"。
    py = find_python(need_requests=True)
    if not py:
        warn("没有同时具备 requests + openpyxl 的 Python，自检可能跳过部分用例")
        py = find_python(need_requests=False)
    check_syntax(py)
    if py:
        head("3b 解释器依赖")
        for mod, need in (("requests", "SkillHub CLI"), ("openpyxl", "自检必需"),
                          ("docx", "docx 用例"), ("win32com", "引擎用例")):
            r = subprocess.run([py, "-c", "import %s" % mod], capture_output=True)
            tag = "✔" if r.returncode == 0 else ("▲ 缺" if mod in ("docx", "win32com") else "✘ 缺")
            print("  %s %-10s %s" % (tag, mod, need))
    check_frontmatter_local()

    out = Path(a.out) if a.out else (Path.cwd() / (SKILL_DIR.name + ".zip"))
    r = subprocess.run([py, str(SKILL_DIR / "scripts/make_publish_package.py"), "--out", str(out)],
                       capture_output=True, text=True)
    print("\n== 打包 ==")
    print("      " + (r.stdout or "").strip().split("\n")[-1])
    if r.returncode:
        fail("打包失败")
    else:
        ok("发布包已生成：%s" % out)

    check_frontmatter_publish(out)
    if not a.fast:
        check_selftest(py, a.fast)
    if a.with_experiment:
        check_experiment(py)
    check_dryrun(py, out)

    print("\n" + "=" * 62)
    if FAILS:
        print("结论：✘ 有 %d 项 FAIL，%d 项 WARN —— 暂不可发布" % (len(FAILS), len(WARNS)))
        for f in FAILS:
            print("   ✘ " + f)
        return 1
    print("结论：✔ 全部检查通过（%d 项 WARN）—— 可以发布" % len(WARNS))
    for w in WARNS:
        print("   ▲ " + w)
    if not a.fast:
        print("\n下一步：")
        print('  MSYS_NO_PATHCONV=1 <python> "%s" publish "%s" --token skh_xxx --changelog "说明"'
              % (Path.home() / ".skillhub/skills_store_cli.py", out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
