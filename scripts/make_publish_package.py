# -*- coding: utf-8 -*-
"""
生成「SkillHub 发布包」：把本机 WorkBuddy 技能变成可直接 publish 的 zip。

为什么需要它
------------
WorkBuddy 技能与 SkillHub 技能的 frontmatter 要求不同：

- WorkBuddy：name / description / metadata（本地校验器只允许这 5 个顶层键）
- SkillHub ：**必须**再补 slug / displayName / version / summary / license

两边字段可以共存（SkillHub 官方实测通过）。本脚本在**不改本机技能**的前提下，
复制一份 → 增补 SkillHub 字段 → 附上公开 README → 打成 zip → 跑本地 dry-run 预检。

用法
----
    python scripts/make_publish_package.py                    # 生成 + 预检
    python scripts/make_publish_package.py --version 1.1.0    # 指定版本
    python scripts/make_publish_package.py --out D:/xxx.zip   # 指定输出

发布（需要 skh_ Token，Token 只在你自己终端里传，不要落到聊天记录里）：
    MSYS_NO_PATHCONV=1 <带 requests 的 python> \
        C:/Users/<你>/.skillhub/skills_store_cli.py publish <zip> \
        --token skh_xxx --changelog "说明"

坑位提醒（实测）
--------------
1. Git Bash 调 `skillhub` 包装脚本会把路径拼成 `c:\\c\\...` → 直接调
   `skills_store_cli.py` 并加 `MSYS_NO_PATHCONV=1`，路径写 `C:/...` 形式。
2. CLI 依赖 `requests`；WorkBuddy 托管的 3.13.12 **没有** requests，
   用系统 Python（如 `C:/Users/<你>/AppData/Local/Programs/Python/Python312/python.exe`）。
3. 报错 409 = slug 被占 → 加 handle 前缀（如 `jt-xxx`）。
4. 报错 566 / "fail to fetch" = 服务端内容扫描器崩了，**不是断网**。
   多半是发布包里某文件内容触发扫描 bug（内嵌代码尤其危险）。
   排查法：先用极简 SKILL.md + 全新 slug 做对照实验，再逐文件二分定位。
5. `--dry-run` 是纯本地校验，**不发 HTTP 请求**，所以发布前必须先跑。
"""
import argparse
import atexit
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL_DIR = HERE.parent                      # 技能根目录（含 SKILL.md 的那一层）
# ⚠️ 不得硬编码任何用户目录 —— 本脚本要随技能公开发布，别人装上后也必须能跑。
#    默认输出到「当前工作目录」，可用 --out 覆盖。
DEFAULT_OUT = Path.cwd() / "jt-excel-4layer-audit.zip"

# ---- SkillHub 必需字段（发布时可覆盖）----
DEFAULTS = {
    "slug": "jt-excel-4layer-audit",
    "displayName": "JT Excel 数据统计分析·四层审计",
    "version": "1.0.0",
    "summary": ("Excel/WPS 表格数据统计分析的端到端交付方法论。四层验证阶梯："
                "L1 独立重实现 / L2 Excel+WPS 多引擎强制重算 / L3 结构审计 / L4 不变量与边界；"
                "用受控注入实验证明只有 L2 能穿透「读缓存不等于真值」的盲区。"),
    "license": "MIT",
}

SKILLHUB_KEYS = ("slug", "displayName", "version", "summary", "license")


def find_python() -> str:
    """找一个装了 requests 的 Python（不硬编码任何用户路径）。"""
    cands = ["python3", "python", "py"]
    # 常见安装位置（存在才试；不写死用户名）
    for base in (Path.home() / "AppData/Local/Programs/Python",
                 Path("/c/Users")):
        if base.is_dir():
            try:
                for d in sorted(base.glob("Python3*")):
                    cands.append(str(d / "python.exe"))
            except OSError:
                pass
    for c in cands:
        exe = c if os.path.isabs(c) else shutil.which(c)
        if not exe or not Path(exe).exists():
            continue
        try:
            r = subprocess.run([exe, "-c", "import requests"], capture_output=True, timeout=20)
            if r.returncode == 0:
                return exe
        except Exception:
            continue
    return ""


def build_frontmatter(fields: dict) -> str:
    """在原 frontmatter 里增补 SkillHub 字段，保留原有字段与排版。"""
    p = SKILL_DIR / "SKILL.md"
    # ⚠️ 必须用 open(newline="")：Path.read_text/write_text 在 3.12 不支持 newline 参数；
    #    且默认换行转换会把 \n 翻成 \r\n（Windows），使发布包变成 CRLF ——
    #    读文本的工具（read_text）会做通用换行转换不受影响，但按字节解析的会失配。
    with open(p, "r", encoding="utf-8", newline="") as f:
        text = f.read()
    m = re.match(r"^---\r?\n(.*?)\r?\n---\r?\n", text, re.S)
    if not m:
        sys.exit("✘ SKILL.md 没有 YAML frontmatter")
    fm = m.group(1)
    eol = "\r\n" if "\r\n" in text else "\n"

    # 去掉可能已存在的 SkillHub 字段（保证幂等：重复生成不会堆叠）
    lines = fm.split("\n")
    kept = []
    for ln in lines:
        km = re.match(r"^([A-Za-z_]+):", ln)
        if km and km.group(1) in SKILLHUB_KEYS:
            continue
        kept.append(ln)
    fm = "\n".join(kept)

    # 插到 name: 行之后，保持字段分组可读
    inject = [
        "slug: %s" % fields["slug"],
        "displayName: \"%s\"" % fields["displayName"],
        "version: %s" % fields["version"],
        "summary: \"%s\"" % fields["summary"],
        "license: %s" % fields["license"],
    ]
    out, done = [], False
    for ln in fm.split("\n"):
        out.append(ln)
        if not done and re.match(r"^name:", ln):
            out.extend(inject)
            done = True
    if not done:
        out = inject + out
    fm_new = eol.join(out)
    return "---" + eol + fm_new + eol + "---" + eol + text[m.end():]


def main() -> int:
    ap = argparse.ArgumentParser(description="生成 SkillHub 发布包并本地预检")
    ap.add_argument("--version", default=DEFAULTS["version"])
    ap.add_argument("--slug", default=DEFAULTS["slug"])
    ap.add_argument("--display-name", default=DEFAULTS["displayName"])
    ap.add_argument("--summary", default=DEFAULTS["summary"])
    ap.add_argument("--license", default=DEFAULTS["license"])
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    ap.add_argument("--no-dry-run", action="store_true", help="跳过 CLI dry-run 预检")
    a = ap.parse_args()

    fields = {
        "slug": a.slug, "displayName": a.display_name, "version": a.version,
        "summary": a.summary, "license": a.license,
    }
    for k, v in fields.items():
        if not v or not str(v).strip():
            sys.exit("✘ %s 不能为空" % k)
        if "<" in str(v) or ">" in str(v):
            sys.exit("✘ %s 不能含尖括号 < >（SkillHub 校验会拒）" % k)
    if not re.fullmatch(r"[A-Za-z0-9-]+", fields["slug"]):
        sys.exit("✘ slug 只能含字母/数字/连字符")
    if not re.match(r"^\d+\.\d+\.\d+$", fields["version"]):
        sys.exit("✘ version 需为语义化版本，如 1.0.0")

    stage = Path(tempfile.mkdtemp(prefix="skillhub-publish-"))
    atexit.register(lambda: shutil.rmtree(stage, ignore_errors=True))
    shutil.copytree(SKILL_DIR, stage, dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))

    with open(stage / "SKILL.md", "w", encoding="utf-8", newline="") as f:
        f.write(build_frontmatter(fields))

    # 公开 README（市场详情页用）
    pub = SKILL_DIR / "references" / "public-readme.md"
    if pub.is_file():
        shutil.copy(pub, stage / "README.md")
    else:
        print("ℹ 未找到 references/public-readme.md，发布包将不含 README.md")

    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    files = []
    for root, dirs, names in os.walk(stage):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", ".git")]
        for n in names:
            fp = Path(root) / n
            if fp.is_file() and not n.endswith((".pyc", ".bak")):
                files.append(fp)
    files.sort()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for fp in files:
            z.write(fp, fp.relative_to(stage).as_posix())

    with zipfile.ZipFile(out) as z:
        assert "SKILL.md" in z.namelist(), "SKILL.md 必须在 zip 根目录"
        assert z.testzip() is None
        inner = z.read("SKILL.md").decode("utf-8")
        for k in SKILLHUB_KEYS:
            if not re.search(r"^%s:" % k, inner, re.M):
                sys.exit("✘ zip 内 SKILL.md 缺字段: %s" % k)
    print("✔ 已生成 %s（%.1f KB / %d 文件）" % (out, out.stat().st_size / 1024, len(files)))
    print("  slug=%s version=%s license=%s" % (fields["slug"], fields["version"], fields["license"]))

    if a.no_dry_run:
        return 0
    py = find_python()
    if not py:
        print("⚠ 未找到带 requests 的 Python，跳过 dry-run 预检。")
        print("   装：<python> -m pip install requests")
        return 0
    cli = Path.home() / ".skillhub" / "skills_store_cli.py"
    if not cli.is_file():
        print("⚠ 未找到 SkillHub CLI，跳过 dry-run。安装：curl -fsSL %s | bash -s -- --cli-only"
              % "https://skillhub-1388575217.cos.ap-guangzhou.myqcloud.com/install/install.sh")
        return 0
    print()
    print("== CLI dry-run 本地预检（不发 HTTP 请求）==")
    env = dict(os.environ, MSYS_NO_PATHCONV="1")
    r = subprocess.run([py, str(cli), "publish", str(out), "--dry-run"],
                       capture_output=True, text=True, env=env)
    sys.stdout.write(r.stdout or "")
    sys.stderr.write(r.stderr or "")
    if r.returncode != 0:
        print("✗ dry-run 未通过，按 references/public-readme.md 里的报错对照表排查")
        return r.returncode
    print("✔ dry-run 通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
