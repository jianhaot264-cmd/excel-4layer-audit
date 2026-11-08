#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键跑完四层验证，产出可复现的证据包。

四层里 L2/L3/L4 是通用脚本，本程序直接调用；L1（独立重实现）是**任务专属**的，
所以用 `--l1` 把你的重算脚本挂进来一起跑（约定：脚本打印含 "不一致项: 无" 或退出码 0 视为通过）。

用法:
    python run_all.py 成品.xlsx --orig 原始.xlsx --expect expected.json --invariants inv.json \
                       --source 数据原表 --l1 我的独立重算.py --out ./验证报告

    # 只跑通用三层（不挂 L1，会在报告里标注"未执行"）
    python run_all.py 成品.xlsx --out ./验证报告

产出:
    <out>/report.md          人读的汇总报告（含 SHA256、各层结论、复现命令）
    <out>/report.json        机读汇总
    <out>/L2_重算.json / L3_结构审计.json / L4_不变量.json / L1_输出.txt

退出码:
    0 = 四层全部执行且全部通过
    1 = 有层 FAIL
    3 = 无 FAIL 但有层未执行（"部分验证"，不可宣称已验证）
    2 = 前置条件不满足
"""
from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def run(cmd: list[str], outfile: str | None = None):
    p = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    text = (p.stdout or "") + (p.stderr or "")
    if outfile:
        with open(outfile, "w", encoding="utf-8") as f:
            f.write(text)
    return p.returncode, text


def main() -> int:
    ap = argparse.ArgumentParser(description="一键四层验证")
    ap.add_argument("xlsx")
    ap.add_argument("--orig")
    ap.add_argument("--expect")
    ap.add_argument("--invariants")
    ap.add_argument("--source", default="数据原表")
    ap.add_argument("--engines", default="auto", help="auto | excel | wps | excel,wps")
    ap.add_argument("--l1", help="任务专属的独立重实现脚本")
    ap.add_argument("--out", default="验证报告")
    a = ap.parse_args()

    if not os.path.isfile(a.xlsx):
        sys.exit("文件不存在: %s" % a.xlsx)
    os.makedirs(a.out, exist_ok=True)

    hash_before = sha256(a.xlsx)
    stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    layers: list[dict] = []

    print("=" * 78)
    print("四层验证 · %s" % a.xlsx)
    print("SHA256(开始) = %s" % hash_before)
    print("=" * 78)

    # ---- L1 独立重实现（任务专属脚本） ----
    if a.l1:
        if not os.path.isfile(a.l1):
            sys.exit("L1 脚本不存在: %s" % a.l1)
        print("\n[L1] 独立重实现 …")
        rc, txt = run([PY, a.l1], os.path.join(a.out, "L1_输出.txt"))
        ok = rc == 0
        print(txt.strip()[-500:] if not ok else "  ✔ 通过（退出码 0）")
        layers.append({"layer": "L1 独立重实现", "ok": ok, "detail": "退出码 %d" % rc})
    else:
        print("\n[L1] 未执行（未提供 --l1 脚本）")
        layers.append({"layer": "L1 独立重实现", "ok": None,
                       "detail": "未执行 —— 请提供任务专属的重算脚本"})

    # ---- L2 多引擎重算 ----
    print("\n[L2] 多引擎重算 …")
    cmd = [PY, os.path.join(HERE, "recalc_verify.py"), a.xlsx, "--engine", a.engines,
           "--json", os.path.join(a.out, "L2_重算.json")]
    if a.expect:
        cmd += ["--expect", a.expect]
    rc, txt = run(cmd, os.path.join(a.out, "L2_输出.txt"))
    ok = rc == 0
    print("  " + ("✔ 全部引擎一致" if ok else "✘ 存在不一致（退出码 %d，若为 2 表示无可用引擎）" % rc))
    layers.append({"layer": "L2 多引擎重算", "ok": ok, "detail": "退出码 %d" % rc})

    # ---- L3 结构审计 ----
    print("\n[L3] 结构审计 …")
    cmd = [PY, os.path.join(HERE, "audit_xlsx.py"), a.xlsx,
           "--json", os.path.join(a.out, "L3_结构审计.json")]
    if a.orig:
        cmd += ["--orig", a.orig]
    if a.expect:
        cmd += ["--expect", a.expect]
    rc, txt = run(cmd, os.path.join(a.out, "L3_输出.txt"))
    n_fail = txt.count("✘ FAIL")
    n_warn = txt.count("▲ WARN")
    ok = rc == 0
    print("  " + ("✔ FAIL 0" if ok else "✘ 有 FAIL") + " / WARN %d" % n_warn)
    layers.append({"layer": "L3 结构审计", "ok": ok, "detail": "FAIL %d / WARN %d" % (n_fail, n_warn)})

    # ---- L4 不变量与边界 ----
    print("\n[L4] 不变量与边界 …")
    cmd = [PY, os.path.join(HERE, "invariant_check.py"), a.xlsx]
    if a.invariants:
        cmd += ["--spec", a.invariants]
    cmd += ["--auto-edge", "--source", a.source, "--json", os.path.join(a.out, "L4_不变量.json")]
    rc, txt = run(cmd, os.path.join(a.out, "L4_输出.txt"))
    n_fail = txt.count("✘ FAIL")
    n_warn = txt.count("▲ WARN")
    ok = rc == 0
    print("  " + ("✔ FAIL 0" if ok else "✘ 有 FAIL") + " / WARN %d" % n_warn)
    layers.append({"layer": "L4 不变量与边界", "ok": ok, "detail": "FAIL %d / WARN %d" % (n_fail, n_warn)})

    # ---- 收尾：确认文件未被验证过程改动 ----
    hash_after = sha256(a.xlsx)
    intact = hash_before == hash_after
    print("\nSHA256(结束) = %s  %s" % (hash_after, "✔ 未改动" if intact else "✘ 文件被改动了！"))
    if not intact:
        layers.append({"layer": "文件完整性", "ok": False,
                       "detail": "验证过程中文件被修改，此前结论全部失效"})

    hard_fail = [l for l in layers if l["ok"] is False]
    skipped = [l for l in layers if l["ok"] is None]

    # 报告 ----
    md = ["# 四层验证报告", "",
          "- **被验证文件**：`%s`" % os.path.abspath(a.xlsx),
          "- **SHA256**：`%s`" % hash_after,
          "- **验证时间**：%s" % stamp,
          "- **原始文件**：`%s`" % (os.path.abspath(a.orig) if a.orig else "（未提供，跳过逐格 diff）"),
          "",
          "## 结论", ""]
    if hard_fail:
        md += ["**存在 FAIL —— 不得宣称已验证：**", ""]
        for l in hard_fail:
            md.append("- %s：%s" % (l["layer"], l["detail"]))
        md.append("")
    elif skipped:
        md += ["**部分验证 —— 无 FAIL，但有层未执行，不得宣称『四层全绿』：**", ""]
        for l in skipped:
            md.append("- %s：%s" % (l["layer"], l["detail"]))
        md.append("")
    else:
        md += ["**四层全部执行且无 FAIL。**", ""]
    md += ["| 层 | 结果 | 说明 |", "|---|---|---|"]
    for l in layers:
        mark = "✔ PASS" if l["ok"] else ("✘ FAIL" if l["ok"] is False else "— 未执行")
        md.append("| %s | %s | %s |" % (l["layer"], mark, l["detail"]))
    md += ["", "## 口径与残留风险", "",
           "> 本报告只证明『算得对』。『解得对』受题目表述与业务含义约束，",
           "> 请把主口径 / 备选口径 / 数据质量问题 / 无法消除的不确定性一并写入交付说明。", "",
           "## 复现方式", "", "```bash"]
    md.append("python run_all.py %s%s%s%s --out %s" % (
        os.path.basename(a.xlsx),
        " --orig %s" % os.path.basename(a.orig) if a.orig else "",
        " --expect %s" % os.path.basename(a.expect) if a.expect else "",
        " --invariants %s" % os.path.basename(a.invariants) if a.invariants else "",
        a.out))
    md += ["```", "",
           "**修改纪律**：任何一次修改后必须重跑本脚本——改一个单元格，之前结论即失效。", ""]
    with open(os.path.join(a.out, "report.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(md))
    json.dump({"file": os.path.abspath(a.xlsx), "sha256": hash_after, "time": stamp,
               "complete": not hard_fail and not skipped, "has_fail": bool(hard_fail),
               "layers": layers, "skipped": [s["layer"] for s in skipped]},
              open(os.path.join(a.out, "report.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)

    if hard_fail:
        rc, verdict = 1, "✘ 存在 FAIL"
    elif skipped:
        rc, verdict = 3, "▲ 部分验证（有层未执行，不可宣称已验证）"
    else:
        rc, verdict = 0, "✔ 四层全部执行且通过"

    print("\n" + "=" * 78)
    print("结论: %s" % verdict)
    if skipped:
        print("      未执行: %s" % ", ".join(s["layer"] for s in skipped))
    print("报告: %s" % os.path.join(a.out, "report.md"))
    print("=" * 78)
    return rc


if __name__ == "__main__":
    sys.exit(main())
