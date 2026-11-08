#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
技能自检：在临时目录里造合成样本，把本技能所有脚本的正常路径与异常路径跑一遍。

目的：换一台机器 / 换一个项目时，先证明**工具链本身是好的**，再去验真实交付物。
否则可能拿着一个悄悄失效的检查脚本去"验证"，得出假阳性结论。

用法:
    python selftest.py            # 只跑不依赖 Office 的检查
    python selftest.py --with-engines    # 额外跑 Excel/WPS 引擎相关用例（较慢）

依赖: openpyxl（必需）、python-docx（docx 用例，缺失则跳过）、pywin32（引擎用例，缺失则跳过）
完全在系统临时目录内进行，不触碰任何真实文件。

退出码: 0 = 全部通过；1 = 有失败。
"""
from __future__ import annotations

import gc
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

RESULTS: list[tuple[str, str, str]] = []


def run(name, args, expect_ok=(0,)):
    p = subprocess.run([PY, os.path.join(HERE, args[0])] + args[1:],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (p.stdout or "") + (p.stderr or "")
    tb = "Traceback (most recent call last)" in out
    ok = p.returncode in expect_ok and not tb
    detail = ""
    if p.returncode not in expect_ok:
        detail = "退出码=%s(期望%s)" % (p.returncode, list(expect_ok))
    if tb:
        detail += (" | " if detail else "") + "未捕获异常: " + out.strip().splitlines()[-1][:120]
    RESULTS.append(("PASS" if ok else "FAIL", name, detail))
    if not ok:
        print("  [失败详情] %s" % name)
        for l in out.splitlines():
            if l.startswith(("✘", "▲")) or "Traceback" in l or l.strip().startswith("File "):
                print("      " + l[:160])
        print("      …… 退出码 %s" % p.returncode)
    return out


def build_sample(workdir: str, name: str = "sample.xlsx") -> str:
    """造一个结构像真实交付物的合成工作簿。"""
    import openpyxl
    from openpyxl.chart import BarChart, Reference

    path = os.path.join(workdir, name)
    wb = openpyxl.Workbook()
    q = wb.active
    q.title = "命题"
    q["A1"] = "合成测试题"
    q["A2"] = "1、各品类的合计金额前三及其占比？"

    src = wb.create_sheet("数据原表")
    src.append(["日期", "品类", "数量", "单价", "门店"])
    rows = [("2026-01-01", "甲", 10, 2.5, "A店"), ("2026-01-01", "乙", 5, 4.0, "B店"),
            ("2026-01-02", "甲", 4, 2.5, "A店"), ("2026-01-02", "丙", 7, 1.5, "A店"),
            ("2026-01-03", "乙", 6, 4.0, "A店"), ("2026-01-03", "丁", 2, 9.0, "B店"),
            ("2026-01-02", "甲", 4, 2.5, "A店")]        # ← 与第 3 行完全重复，用于触发 auto-edge
    for r in rows:
        src.append(list(r))
    src.append(["2026-01-04", 711, 1, 1, "711"])         # ← 分类列混入数值，用于触发类型混杂告警

    ans = wb.create_sheet("作答区")
    ans["A1"] = "作答区"
    ans["A2"] = "统计天数"
    ans["B2"] = "=SUMPRODUCT(1/COUNTIF(数据原表!$A$2:$A$9,数据原表!$A$2:$A$9))"
    ans.append([])
    ans.append(["品类", "合计金额", "排名"])
    for i, cat in enumerate(("甲", "乙", "丙", "丁")):
        r = 5 + i
        ans.cell(r, 1, cat)
        ans.cell(r, 2, '=SUMPRODUCT((数据原表!$B$2:$B$9=$A%d)*数据原表!$C$2:$C$9*数据原表!$D$2:$D$9)' % r)
        ans.cell(r, 3, '=RANK(B%d,$B$5:$B$8)' % r)
    ans["A10"] = "答案"
    for i in range(3):
        r = 11 + i
        ans.cell(r, 1, i + 1)
        ans.cell(r, 2, '=INDEX($A$5:$A$8,MATCH(LARGE($B$5:$B$8,%d),$B$5:$B$8,0))' % (i + 1))
        ans.cell(r, 3, '=LARGE($B$5:$B$8,%d)' % (i + 1))
    ans["A15"] = "图表数据源"
    ans.append(["品类", "合计金额"])
    for i in range(4):
        r = 16 + i
        ans.cell(r, 1, '=INDEX($A$5:$A$8,MATCH(LARGE($B$5:$B$8,%d),$B$5:$B$8,0))' % (i + 1))
        ans.cell(r, 2, '=LARGE($B$5:$B$8,%d)' % (i + 1))
    ch = BarChart()
    ch.title = "品类合计"
    ch.add_data(Reference(ans, min_col=2, min_row=15, max_row=20), titles_from_data=True)
    ch.set_categories(Reference(ans, min_col=1, min_row=16, max_row=20))
    ans.add_chart(ch, "E2")

    wb.save(path)
    return path


def make_broken(src: str, dst: str, style: str) -> int:
    """把公式标签改坏，返回改坏处数。style: 'plain'（<f>=…）| 'array'（<f t="array">=…）"""
    n = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            d = zin.read(it.filename)
            if it.filename.startswith("xl/worksheets/") and d:
                t = d.decode("utf-8", "ignore")
                if style == "plain":
                    n += t.count("<f>")
                    t = t.replace("<f>", "<f>=")
                else:
                    def _to_array(m: re.Match) -> str:
                        return '<f t="array" ref="A1">=' + m.group(1) + "</f>"
                    t, k = re.subn(r"<f>(?!=)([^<]*)</f>", _to_array, t)
                    n += k
                d = t.encode("utf-8")
            zout.writestr(it, d)
    return n


def materialize(path: str) -> tuple[bool, str]:
    """用真实引擎打开并另存，使公式产生缓存值。

    openpyxl 直接写出的 xlsx **没有公式缓存**（data_only 读回全是 None），
    所有"读缓存比对"的检查都会失灵。必须先让真实引擎保存一次。
    依次尝试 Excel 与 WPS；都不可用则返回 False（调用方跳过依赖缓存的用例）。
    """
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        return False, "未安装 pywin32"
    last = ""
    for pid in ("Excel.Application", "KET.Application"):
        app = wb = None
        try:
            pythoncom.CoInitialize()
            app = win32.DispatchEx(pid)
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=False)
            try:
                app.CalculateFullRebuild()
            except Exception:                               # noqa: BLE001
                pass
            wb.Save()
            wb.Close(SaveChanges=False)
            app.Quit()
            wb = app = None
            gc.collect()
            pythoncom.CoUninitialize()
            return True, pid
        except Exception as e:                              # noqa: BLE001
            last = "%s: %s" % (type(e).__name__, e)
            for closer in (lambda: wb and wb.Close(SaveChanges=False),
                           lambda: app and app.Quit()):
                try:
                    closer()
                except Exception:                           # noqa: BLE001
                    pass
            wb = app = None
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:                               # noqa: BLE001
                pass
    return False, last


def main() -> int:
    with_engines = "--with-engines" in sys.argv
    T = tempfile.mkdtemp(prefix="excelwps_selftest_")
    print("=" * 78)
    print("技能自检 · 工作目录 %s" % T)
    print("=" * 78)

    try:
        import openpyxl                                      # noqa: F401
    except ImportError:
        sys.exit("缺少依赖：pip install openpyxl")

    good = build_sample(T, "sample.xlsx")
    orig = build_sample(T, "orig.xlsx")          # 独立生成，避免与 good 指向同一文件
    bp = os.path.join(T, "broken_plain.xlsx")
    ba = os.path.join(T, "broken_array.xlsx")
    n1 = make_broken(good, bp, "plain")
    n2 = make_broken(good, ba, "array")
    print("\n样本: %s\n坏样本: plain=%d 处, array=%d 处" % (good, n1, n2))

    notzip = os.path.join(T, "fake.xlsx")
    open(notzip, "w").write("not a zip")
    missing = os.path.join(T, "nope.xlsx")

    # openpyxl 写出的 xlsx 没有公式缓存 —— 先让真实引擎保存一次，才有"缓存值"可比
    ok_eng, via = materialize(good)
    if ok_eng:
        print("\n缓存物化: ✔ 已用 %s 重算并保存" % via)
    else:
        print("\n缓存物化: ✘ 跳过依赖缓存的用例（%s）" % via)

    # audit 的 --orig 必须是"打补丁之前"的状态，否则 patch 用例的改动会被判成"覆盖"。
    # 这本身就是一个提醒：**审计基准要和被审计状态对齐**。
    audit_base = os.path.join(T, "audit_base.xlsx")
    shutil.copy2(good, audit_base)
    # patch 用例单独用一份副本，避免改坏后续 audit 用的样本
    patch_target = os.path.join(T, "patch_target.xlsx")
    shutil.copy2(good, patch_target)

    exp_ok = os.path.join(T, "expect_ok.json")
    json.dump({"作答区!B2": 4}, open(exp_ok, "w", encoding="utf-8"), ensure_ascii=False)
    exp_bad = os.path.join(T, "expect_bad.json")
    json.dump({"作答区!B2": 999}, open(exp_bad, "w", encoding="utf-8"), ensure_ascii=False)

    inv = os.path.join(T, "inv.json")
    json.dump({"checks": [
        {"type": "rank_contiguous", "ref": "作答区!C5:C8"},
        {"type": "sorted_desc", "ref": "作答区!B16:B19"},
        {"type": "sum_equals", "range": "作答区!B16:B19", "target": "作答区!B5:B8"},
        {"type": "unique", "ref": "作答区!A5:A8"},
        {"type": "non_negative", "ref": "作答区!B5:B8"},
        {"type": "top_k_match", "answer": "作答区!B11:B13", "pool_labels": "作答区!A5:A8",
         "pool_values": "作答区!B5:B8", "k": 3},
        {"type": "no_hardcoded", "sheets": ["命题"], "allow": ["命题!A11:A13"]},
    ]}, open(inv, "w", encoding="utf-8"), ensure_ascii=False)

    patch = os.path.join(T, "patch.json")
    json.dump({"作答区!A30": "自检写入的说明", "作答区!B30": 12.5, "作答区!C30": True},
              open(patch, "w", encoding="utf-8"), ensure_ascii=False)
    patch_dup = os.path.join(T, "patch_dup.json")
    json.dump({"作答区!A1": "试图覆盖"}, open(patch_dup, "w", encoding="utf-8"), ensure_ascii=False)

    # ---------- fix_formula_tags ----------
    print("\n--- fix_formula_tags.py ---")
    out = run("fix 修复普通形式", ["fix_formula_tags.py", bp])
    print("   ", [l for l in out.splitlines() if "复验" in l])
    run("fix 修复数组形式", ["fix_formula_tags.py", ba])
    run("fix dry-run", ["fix_formula_tags.py", good, "--dry-run"])
    run("fix 已合规", ["fix_formula_tags.py", good])
    run("fix 文件不存在", ["fix_formula_tags.py", missing], expect_ok=(1, 2))
    run("fix keep-array", ["fix_formula_tags.py", ba, "--keep-array"])
    with zipfile.ZipFile(bp) as z:
        left = sum(len(re.findall(r"<f(\s[^>]*)?>(=)", z.read(n).decode("utf-8", "ignore"), re.S))
                   for n in z.namelist() if n.startswith("xl/worksheets/"))
    RESULTS.append(("PASS" if left == 0 else "FAIL", "fix 修复后残留 0 处（实际 %d）" % left, ""))

    # ---------- patch_cells ----------
    print("\n--- patch_cells.py ---")
    PT = patch_target
    run("patch 新建行多类型", ["patch_cells.py", PT, "--spec", patch])
    run("patch 拒绝覆盖", ["patch_cells.py", PT, "--spec", patch_dup], expect_ok=(1,))
    run("patch dry-run", ["patch_cells.py", PT, "--spec", patch, "--dry-run"])
    run("patch 表不存在", ["patch_cells.py", PT, "--set", '{"没有表!A1":"x"}'], expect_ok=(1,))
    run("patch 非法坐标", ["patch_cells.py", PT, "--set", '{"作答区!1A":"x"}'], expect_ok=(1,))
    run("patch 非法 JSON", ["patch_cells.py", PT, "--set", '{"作答区!A1":'], expect_ok=(1,))
    run("patch 文件不存在", ["patch_cells.py", missing, "--set", '{"A!A1":"x"}'], expect_ok=(1,))
    run("patch 非 zip", ["patch_cells.py", notzip, "--set", '{"A!A1":"x"}'], expect_ok=(1,))
    with zipfile.ZipFile(PT) as z:
        d = z.read("xl/worksheets/sheet3.xml").decode("utf-8")
    RESULTS.append(("PASS" if not re.search(r"</row>\s*<c\s", d) else "FAIL",
                    "patch 未产生脱离 <row> 的单元格", ""))
    import openpyxl
    v = openpyxl.load_workbook(PT, data_only=True)["作答区"]
    RESULTS.append(("PASS" if (v["A30"].value == "自检写入的说明" and v["B30"].value == 12.5
                               and v["C30"].value is True) else "FAIL",
                    "patch 三种类型值读回正确", ""))

    # ---------- audit_xlsx ----------
    print("\n--- audit_xlsx.py ---")
    run("audit 基础", ["audit_xlsx.py", good, "--json", os.path.join(T, "a.json")])
    if ok_eng:
        run("audit 全参数（未打补丁，应 0 覆盖）",
            ["audit_xlsx.py", good, "--orig", audit_base, "--expect", exp_ok,
             "--json", os.path.join(T, "a2.json")])
        out = run("audit 补丁=新增而非覆盖",
                  ["audit_xlsx.py", PT, "--orig", audit_base, "--json", os.path.join(T, "a3.json")])
        line = next((l for l in out.splitlines() if "新增单元格" in l), "")
        mm = re.search(r"(\d+)", line)
        n_added = int(mm.group(1)) if mm else 0
        RESULTS.append(("PASS" if n_added >= 3 else "FAIL",
                        "audit 把 patch 判为新增 %d 个（而非覆盖）" % n_added, ""))
    # 「允许覆盖」会真的改掉 A1，必须放在 audit 之后跑，否则 audit 会正确地报"被覆盖"
    run("patch 允许覆盖", ["patch_cells.py", PT, "--spec", patch_dup, "--overwrite"])

    run("audit 期望值不符", ["audit_xlsx.py", good, "--expect", exp_bad], expect_ok=(1,))
    run("audit 文件不存在", ["audit_xlsx.py", missing], expect_ok=(1, 2))
    run("audit 非 zip", ["audit_xlsx.py", notzip], expect_ok=(1, 2))
    run("audit expect 不存在", ["audit_xlsx.py", good, "--expect", missing], expect_ok=(1,))
    run("audit orig 不存在", ["audit_xlsx.py", good, "--orig", missing], expect_ok=(1, 2))

    # ---------- invariant_check ----------
    print("\n--- invariant_check.py ---")
    run("inv 仅 auto", ["invariant_check.py", good, "--auto-edge", "--source", "数据原表"])
    if ok_eng:
        run("inv spec（含缓存）", ["invariant_check.py", good, "--spec", inv,
                                  "--auto-edge", "--source", "数据原表",
                                  "--json", os.path.join(T, "i.json")])
    run("inv 无参数", ["invariant_check.py", good])
    badinv = os.path.join(T, "bad_inv.json")
    json.dump({"checks": [{"type": "没这个检查", "ref": "作答区!A1"}]},
              open(badinv, "w", encoding="utf-8"), ensure_ascii=False)
    run("inv 未知检查类型", ["invariant_check.py", good, "--spec", badinv], expect_ok=(1,))
    run("inv 文件不存在", ["invariant_check.py", missing], expect_ok=(1,))
    run("inv spec 不存在", ["invariant_check.py", good, "--spec", missing], expect_ok=(1,))
    run("inv 数据源不存在", ["invariant_check.py", good, "--auto-edge", "--source", "没有表"], expect_ok=(1,))

    # ---------- run_all ----------
    print("\n--- run_all.py ---")
    l1 = os.path.join(T, "l1.py")
    open(l1, "w", encoding="utf-8").write(
        "print('L1 合成自检脚本'); raise SystemExit(0)\n")
    if ok_eng:
        run("run_all 完整（含 L1）", ["run_all.py", good, "--orig", orig, "--expect", exp_ok,
                                     "--invariants", inv, "--source", "数据原表", "--l1", l1,
                                     "--out", os.path.join(T, "rep")])
        try:
            j = json.load(open(os.path.join(T, "rep", "report.json"), encoding="utf-8"))
            RESULTS.append(("PASS" if j.get("complete") is True else "FAIL",
                            "run_all report.json complete=true",
                            str(j.get("layers"))[:90]))
        except Exception as e:                              # noqa: BLE001
            RESULTS.append(("FAIL", "run_all report.json 可读", str(e)[:90]))
    run("run_all 缺 L1 → 退出码 3", ["run_all.py", good, "--out", os.path.join(T, "rep2")],
        expect_ok=(3,))
    RESULTS.append(("PASS" if os.path.exists(os.path.join(T, "rep2", "report.md")) else "FAIL",
                    "run_all 产出 report.md", ""))

    # ---------- 引擎相关（可选） ----------
    print("\n--- 引擎相关 ---")
    run("recalc list-engines", ["recalc_verify.py", good, "--list-engines"], expect_ok=(0, 2))
    run("recalc 不存在的引擎", ["recalc_verify.py", good, "--engine", "notepad"], expect_ok=(2,))
    run("recalc 文件不存在", ["recalc_verify.py", missing, "--list-engines"], expect_ok=(0, 2))
    try:
        from docx import Document                          # noqa: F401
        docx = os.path.join(T, "s.docx")
        run("docx 生成自检样本", ["make_docx_field_sample.py", docx])
        run("docx list-engines", ["verify_docx.py", docx, "--list-engines"], expect_ok=(0, 2))
        run("docx 不存在的引擎", ["verify_docx.py", docx, "--engine", "notepad"], expect_ok=(2,))
        run("docx expect 不存在", ["verify_docx.py", docx, "--expect", missing], expect_ok=(1,))
        if with_engines:
            run("docx 双引擎（真重算）", ["verify_docx.py", docx, "--engine", "word,wps"], expect_ok=(0, 2))
        if with_engines:
            run("recalc excel 引擎", ["recalc_verify.py", good, "--engine", "excel"], expect_ok=(0, 2))
    except ImportError:
        print("  · python-docx 未安装，跳过 docx 用例（pip install python-docx）")

    # ---------- 汇总 ----------
    print("\n" + "=" * 78)
    fails = [r for r in RESULTS if r[0] == "FAIL"]
    for st, name, d in RESULTS:
        print("%s  %s%s" % ("✔" if st == "PASS" else "✘", name, ("  ← " + d) if d else ""))
    print("=" * 78)
    print("通过 %d / %d，失败 %d" % (len(RESULTS) - len(fails), len(RESULTS), len(fails)))
    print("=" * 78)
    shutil.rmtree(T, ignore_errors=True)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
