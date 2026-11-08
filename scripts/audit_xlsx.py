#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx 交付前全维度审计。

用法:
    python audit_xlsx.py 成品.xlsx
    python audit_xlsx.py 成品.xlsx --orig 原始.xlsx
    python audit_xlsx.py 成品.xlsx --expect expected.json
    python audit_xlsx.py 成品.xlsx --orig 原始.xlsx --json report.json

expected.json 形如:
    {"命题!B4": "趣多多", "命题!C4": 6256, "作答区!C39": 86}
    （值为数字时按 0.005 容差比较，字符串按去空白后精确比较）

退出码: 0 = 全部通过；1 = 存在 FAIL。
仅依赖 openpyxl（+ 标准库）。缺 openpyxl 时脚本会给出明确提示。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:
    sys.exit("缺少依赖：请先安装 openpyxl（pip install openpyxl）")

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

ERROR_VALUES = {"#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "#ERROR!"}

# 365 / 新版 WPS 专属函数，旧版会 #NAME?。
# ⚠️ 必须用「函数名前不能是字母数字」的边界断言：黑名单里的 IFS( 会命中 COUNTIFS( / SUMIFS(，
#    SORT( 会命中 SORTBY( 的前缀 —— 这正是本类审计最容易出现的假警报。
_NEW_FUNCS = (
    "UNIQUE", "FILTER", "SORTBY", "SORT", "SEQUENCE", "XLOOKUP", "LET",
    "TEXTJOIN", "IFS", "MAXIFS", "MINIFS", "RANDARRAY", "TAKE", "DROP",
)
NEW_FUNC_RE = re.compile(r"(?<![A-Z0-9_])(%s)\s*\(" % "|".join(_NEW_FUNCS))


REPORTS: list[tuple[str, str, str]] = []   # (level, title, detail)  level: PASS/FAIL/WARN/INFO


def rec(level: str, title: str, detail: str = "") -> None:
    REPORTS.append((level, title, detail))


def is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def sheet_xml_map(z: zipfile.ZipFile) -> dict[str, str]:
    """工作表名 -> zip 内 XML 路径（按 r:id 正确解析，不靠 sheetN 顺序猜）。"""
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        return {}
    rid2target = {}
    for r in rels:
        rid2target[r.get("Id")] = r.get("Target", "")
    out = {}
    for sh in wb.iter(NS_MAIN + "sheet"):
        tgt = rid2target.get(sh.get(NS_REL + "id"), "")
        tgt = tgt.lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        out[sh.get("name")] = tgt
    return out


# ---------------------------------------------------------------- 1. 包 / XML
def check_package(z: zipfile.ZipFile) -> None:
    bad_zip = z.testzip()
    rec("PASS" if bad_zip is None else "FAIL", "zip 完整性",
        "损坏成员: %s" % bad_zip if bad_zip else "全部成员可读")

    failed = []
    for n in z.namelist():
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:                      # noqa: BLE001
                failed.append("%s (%s)" % (n, str(e)[:60]))
    rec("PASS" if not failed else "FAIL", "XML 可解析性",
        "失败 %d 个: %s" % (len(failed), failed[:5]) if failed else "%d 个部件全部合法" % len(z.namelist()))

    if "xl/workbook.xml" not in z.namelist():
        rec("FAIL", "工作簿主部件", "缺少 xl/workbook.xml")


# ---------------------------------------------------------------- 2. 公式标签
def check_formula_tags(z: zipfile.ZipFile) -> None:
    lead, arr, total = [], [], 0
    nocache = []
    for n in z.namelist():
        if not n.startswith("xl/worksheets/") or not n.endswith(".xml"):
            continue
        txt = z.read(n).decode("utf-8", "ignore")
        for m in re.finditer(r"<f(\s[^>]*)?>(.*?)</f>", txt, re.S):
            total += 1
            if m.group(2).startswith("="):
                lead.append((n, m.group(2)[:45]))
            if m.group(1) and 't="array"' in m.group(1):
                arr.append((n, m.group(1).strip()))
        # 公式单元格缺缓存值 -> 部分阅读器会显示空白
        for cm in re.finditer(r'<c\s+r="([A-Z]+\d+)"[^>]*>(.*?)</c>', txt, re.S):
            if "<f" in cm.group(2) and "<v>" not in cm.group(2):
                nocache.append("%s!%s" % (n.rsplit("/", 1)[-1], cm.group(1)))
    rec("PASS" if not lead else "FAIL", "公式标签无前导等号",
        "公式总数 %d，残留 %d 个%s" % (total, len(lead), ("：%s" % lead[:3]) if lead else ""))
    rec("PASS" if not arr else "WARN", "无遗留数组公式标记",
        "残留 %d 个%s" % (len(arr), ("：%s" % arr[:3]) if arr else
                        "（数组公式会让阅读者编辑时必须按 Ctrl+Shift+Enter）"))
    rec("PASS" if not nocache else "WARN", "公式单元格均有缓存值",
        "全部 %d 个公式单元格都带 <v> 缓存" % total if not nocache else
        "%d 个公式缺缓存值，部分阅读器会显示空白: %s" % (len(nocache), nocache[:8]))


def check_new_functions(z: zipfile.ZipFile) -> None:
    hits = []
    for n in z.namelist():
        if not n.startswith("xl/worksheets/") or not n.endswith(".xml"):
            continue
        txt = z.read(n).decode("utf-8", "ignore")
        for m in re.finditer(r"<f(\s[^>]*)?>(.*?)</f>", txt, re.S):
            for fn in NEW_FUNC_RE.finditer(m.group(2).upper()):
                hits.append((n, fn.group(1)))
    rec("PASS" if not hits else "FAIL", "函数版本兼容性",
        "无 365/新版 WPS 独占函数" if not hits else
        "发现 %d 处: %s（旧版 Excel/WPS 会 #NAME?，并连带毁掉引用它的单元格）"
        % (len(hits), sorted(set(h[1] for h in hits))))


# ---------------------------------------------------------------- 3. 内容层
def check_content(path: str, orig: str | None) -> None:
    wbv = openpyxl.load_workbook(path, data_only=True)
    wbf = openpyxl.load_workbook(path)

    # 3.1 错误值
    errs = []
    for w in wbv.worksheets:
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip() in ERROR_VALUES:
                    errs.append("%s!%s=%s" % (w.title, c.coordinate, c.value))
    rec("PASS" if not errs else "FAIL", "错误值扫描",
        "全表无公式错误值" if not errs else "发现 %d 处: %s" % (len(errs), errs[:8]))

    # 3.2 公式结构一致性
    lines = []
    for w in wbf.worksheets:
        groups = Counter()
        where = defaultdict(list)
        for row in w.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    norm = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "#", v)
                    groups[norm] += 1
                    where[norm].append(c.coordinate)
        lines.append("  [%s] %d 种结构 / %d 个公式" % (w.title, len(groups), sum(groups.values())))
        for norm, cnt in groups.most_common(6):
            lines.append("      x%-4d %s" % (cnt, norm[:96]))
        singletons = [n for n, c in groups.items() if c == 1]
        if singletons and len(groups) > 1:
            lines.append("      ⚠ 仅出现 1 次的结构: %s"
                         % ", ".join("%s @ %s" % (n[:48], where[n][0]) for n in singletons[:3]))
    rec("INFO", "公式结构分布", "\n".join(lines))

    # 3.3 引用范围 vs 真实数据行数
    lines = []
    for w in wbf.worksheets:
        used = {r[0].row for r in w.iter_rows() if any(c.value is not None for c in r)}
        if used:
            lines.append("  [%s] 有值行 %d~%d（共 %d 行）"
                         % (w.title, min(used), max(used), len(used)))
    # 抓取跨表引用里出现的最大行号
    refmax = defaultdict(set)
    for w in wbf.worksheets:
        for row in w.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    for m in re.finditer(r"([\u4e00-\u9fa5A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?(\d+):\$?[A-Z]{1,3}\$?(\d+)", v):
                        refmax[(w.title, m.group(1))].add((int(m.group(2)), int(m.group(3))))
    for (st, ref), rngs in sorted(refmax.items()):
        lines.append("  [%s] 引用「%s」的区间: %s" % (st, ref, sorted(rngs)))
    rec("INFO", "真实数据边界与引用区间", "\n".join(lines))

    # 3.4 分类列取值体检：只在"纯表格工作表"上做，避免把 表头+数据、问答式工作表 误报
    #     典型真问题：分类列（如"零售商"）里混进了代码/数字（如把 7-Eleven 写成 711），
    #     任何按该列筛选/分组的统计都会静默漏掉这批行。
    lines, awk = [], []
    for w in wbv.worksheets:
        wf = wbf[w.title]
        usedcols = {c.column_letter for r in w.iter_rows() for c in r if c.value is not None}
        datarows = [r for r in w.iter_rows() if any(c.value is not None for c in r)]
        if not usedcols or len(datarows) < 8:
            continue
        # 纯表格判定：已用到的每一列，首行都有值（无断列）
        if any(w[col][0].value is None for col in usedcols):
            lines.append("  [%s] 跳过（非纯表格：首行存在空列，可能是问答/分区式排版）" % w.title)
            continue
        coltypes = defaultdict(Counter)
        for r in w.iter_rows(min_row=2):                # 排除表头行
            for c in r:
                v = c.value
                if v is None:
                    continue
                if isinstance(v, bool):
                    t = "text"
                elif isinstance(v, (int, float)):
                    t = "num"
                elif isinstance(v, str):
                    t = "text"
                else:
                    t = "other"
                coltypes[c.column_letter][t] += 1
        for col, tc in coltypes.items():
            if tc["num"] >= 2 and tc["text"] >= 2:
                awk.append("  ⚠ [%s] %s 列混入异质取值：数值 %d 个 + 文本 %d 个"
                           % (w.title, col, tc["num"], tc["text"]))
        lines.append("  [%s] 纯表格，已检查 %d 列" % (w.title, len(coltypes)))
    rec("PASS" if not awk else "WARN", "分类列类型一致性",
        "无分类列混入异质取值" if not awk else
        "发现 %d 列混入异质取值（按该列筛选/分组会静默漏行，务必确认是否影响结论）:\n%s"
        % (len(awk), "\n".join(awk)))
    rec("INFO", "分类列体检范围", "\n".join(lines) or "  无可检查的纯表格")

    # 3.5 数字格式（只盯"有公式的输出列"，源数据列无需处理）
    lines = []
    for w in wbv.worksheets:
        wf = wbf[w.title]
        has_formula = defaultdict(bool)
        for row in wf.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    has_formula[c.column_letter] = True
        colinfo = defaultdict(list)
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    colinfo[c.column_letter].append(wf[c.coordinate].number_format)
        for col, fmts in colinfo.items():
            if len(fmts) < 3 or not has_formula[col]:
                continue
            top, cnt = Counter(fmts).most_common(1)[0]
            lines.append("  [%s] %s 列 %d 个数值（含公式），主格式 %r%s"
                         % (w.title, col, len(fmts), top,
                            "  ⚠ 输出列未设数字格式" if top == "General" else ""))
    rec("INFO", "数字格式抽查（仅输出列）", "\n".join(lines) or "  无公式输出列")

    # 3.5 与原文件 diff
    if orig:
        # ⚠️ 必须"同视图"比较：原文件若未经真实引擎保存，公式单元格没有缓存值，
        #    用 data_only 读是 None、用普通模式读是公式文本——与成品（可能有缓存）混比，
        #    会把每个公式单元格都误判成"内容被覆盖"。
        o_v = openpyxl.load_workbook(orig, data_only=True)     # 静态值视图
        o_f = openpyxl.load_workbook(orig)                     # 公式视图
        added, overwritten, formula_kept, formula_lost = [], [], 0, 0
        for name in o_v.sheetnames:
            ov, of = o_v[name], o_f[name]
            if name not in wbv.sheetnames:
                rec("FAIL", "工作表缺失", "原始工作表 %r 在成品中不存在" % name)
                continue
            dv, df = wbv[name], wbf[name]
            # ⚠️ 必须遍历"两份文件范围的并集"：只遍历原文件的 max_row/max_column，
            #    会漏掉成品里超出原表范围的新增单元格（比如新写的作答行），统计就永远是 0
            max_r = max(ov.max_row, dv.max_row)
            max_c = max(ov.max_column, dv.max_column)
            for r in range(1, max_r + 1):
                for c in range(1, max_c + 1):
                    cell = ov.cell(r, c)
                    vo_v, vd_v = cell.value, dv.cell(r, c).value
                    vo_f, vd_f = of.cell(r, c).value, df.cell(r, c).value
                    coord = "%s!%s" % (name, cell.coordinate)
                    if isinstance(vo_f, str) and vo_f.startswith("="):
                        # 原来是公式：只要成品里它仍是公式就算保留；
                        # 若被换成了硬编码常数，才是真正的"被覆盖"
                        if isinstance(vd_f, str) and vd_f.startswith("="):
                            formula_kept += 1
                        else:
                            formula_lost += 1
                            overwritten.append("%s 原=公式%s 现=常数%r" % (coord, vo_f[:30], vd_v))
                        continue
                    if vo_v is None and vd_v is None:
                        continue
                    if vo_v == vd_v:
                        continue
                    # 数值/日期/文本统一走容差比较（Excel 的 1e-16 级浮点噪声不该算覆盖）
                    if is_num(vo_v) and is_num(vd_v):
                        if abs(float(vo_v) - float(vd_v)) <= max(0.005, abs(float(vo_v)) * 1e-9):
                            continue
                    (overwritten if vo_v is not None else added).append(
                        "%s 原=%r 现=%r" % (coord, vo_v, vd_v))
        rec("PASS" if not overwritten else "FAIL", "原始内容未被覆盖",
            ("静态值全部一致；原公式 %d 个全部保留" % formula_kept) if not overwritten
            else "被覆盖 %d 处: %s" % (len(overwritten), overwritten[:8]))
        if formula_lost:
            rec("FAIL", "公式被硬编码常数替换", "%d 个公式变成了常数，数据一变就不会跟着变" % formula_lost)
        rec("INFO", "新增单元格 %d 个" % len(added), "\n".join("  " + a for a in added[:40])
            + ("\n  …" if len(added) > 40 else ""))
        new_sheets = [s for s in wbv.sheetnames if s not in o_v.sheetnames]
        rec("PASS" if new_sheets else "WARN", "新增工作表", new_sheets or "无")


def check_expect(path: str, expect_path: str) -> None:
    wbv = openpyxl.load_workbook(path, data_only=True)
    try:
        spec = json.load(open(expect_path, encoding="utf-8"))
    except Exception as e:                              # noqa: BLE001
        rec("FAIL", "期望值清单读取失败", "%s: %s（请用绝对路径）" % (type(e).__name__, e))
        return
    bad = []
    for ref, exp in spec.items():
        sheet, _, coord = ref.partition("!")
        if sheet not in wbv.sheetnames:
            bad.append("%s -> 工作表不存在" % ref)
            continue
        got = wbv[sheet][coord].value
        if isinstance(exp, (int, float)) and isinstance(got, (int, float)):
            ok = abs(float(got) - float(exp)) < 0.005
        else:
            ok = str(got).strip() == str(exp).strip()
        if not ok:
            bad.append("%s 期望=%r 实际=%r" % (ref, exp, got))
    rec("PASS" if not bad else "FAIL", "期望值比对 (%d 项)" % len(spec),
        "全部命中" if not bad else "不符 %d 项: %s" % (len(bad), bad[:8]))


def check_charts(z: zipfile.ZipFile) -> None:
    charts = [n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n)]
    if not charts:
        rec("INFO", "图表", "无图表")
        return
    lines = []
    ok = True
    for n in sorted(charts):
        d = z.read(n).decode("utf-8", "ignore")
        refs = [r.replace("&apos;", "") for r in re.findall(r"<c:f>(.*?)</c:f>", d)]
        pts = re.findall(r"<c:v>([-\d.]+)</c:v>", d)
        title = (re.findall(r"<a:t>(.*?)</a:t>", d) or ["(无标题)"])[0]
        bd = re.findall(r'<c:barDir val="(.*?)"/>', d)
        lines.append("  %s 标题=%r 类型=%s" % (n, title, bd[0] if bd else "-"))
        lines.append("    引用: %s" % refs)
        lines.append("    缓存点数: %d  首值: %s" % (len(pts), pts[:8]))
        if not pts:
            ok = False
            lines.append("    ⚠ 无缓存数值，图表可能空白")
        if any(r.strip() == "" for r in refs):
            ok = False
    rec("PASS" if ok else "WARN", "图表完整性", "\n".join(lines))


def main() -> int:
    ap = argparse.ArgumentParser(description="xlsx 交付前全维度审计")
    ap.add_argument("xlsx")
    ap.add_argument("--orig", help="原始文件，用于逐格 diff")
    ap.add_argument("--expect", help="期望值 JSON，用于定点校验")
    ap.add_argument("--json", help="把报告同时写成 JSON")
    a = ap.parse_args()

    # 前置校验：给出清晰错误，而不是抛 FileNotFoundError / BadZipFile 的堆栈
    for label, path in (("目标文件", a.xlsx), ("原始文件(--orig)", a.orig)):
        if path and not os.path.isfile(path):
            sys.exit("文件不存在（%s）: %s" % (label, path))
    try:
        z = zipfile.ZipFile(a.xlsx)
    except zipfile.BadZipFile as e:
        sys.exit("不是合法的 xlsx（zip 打不开）: %s —— %s" % (a.xlsx, e))

    print("=" * 78)
    print("审计对象: %s" % a.xlsx)
    print("=" * 78)

    check_package(z)
    check_formula_tags(z)
    check_new_functions(z)
    try:
        check_content(a.xlsx, a.orig)
    except Exception as e:                              # noqa: BLE001
        rec("FAIL", "内容层检查异常", "%s: %s" % (type(e).__name__, e))
    if a.expect:
        check_expect(a.xlsx, a.expect)
    check_charts(z)
    z.close()

    print()
    order = {"FAIL": 0, "WARN": 1, "PASS": 2, "INFO": 3}
    marker = {"FAIL": "✘ FAIL", "WARN": "▲ WARN", "PASS": "✔ PASS", "INFO": "· INFO"}
    for lvl in ("FAIL", "WARN", "PASS", "INFO"):
        for l, t, d in REPORTS:
            if l != lvl:
                continue
            print("%s  %s" % (marker[l], t))
            if d:
                for line in str(d).splitlines():
                    print("        " + line)

    nf = sum(1 for l, _, _ in REPORTS if l == "FAIL")
    nw = sum(1 for l, _, _ in REPORTS if l == "WARN")
    print()
    print("=" * 78)
    print("结论: FAIL %d 项 / WARN %d 项 / PASS %d 项"
          % (nf, nw, sum(1 for l, _, _ in REPORTS if l == "PASS")))
    print("-" * 78)
    print("本脚本只做「结构与静态」审计，属验证阶梯的 L3。要宣称「已验证」，还需：")
    print("  L1 独立重实现   ：脱离表格引擎，从数据源用纯 Python 重算全部答案并逐格比对")
    print("  L2 第二引擎重算 ：python scripts/recalc_verify.py 成品.xlsx --expect expected.json")
    print("                    （读缓存值 ≠ 真值，必须用 Excel 引擎强制全量重算）")
    print("  L4 不变量与边界 ：python scripts/invariant_check.py 成品.xlsx --spec inv.json --auto-edge")
    print("=" * 78)

    if a.json:
        json.dump([{"level": l, "title": t, "detail": d} for l, t, d in REPORTS],
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return 1 if nf else 0


if __name__ == "__main__":
    sys.exit(main())
