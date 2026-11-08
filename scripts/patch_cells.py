#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接往 xlsx 指定单元格写值——不经过表格引擎，因此：
  - 不会重新触发 editor_sdk 的写入行为（不会把公式改回 <f>=... 的前导等号形态）
  - 图表、样式、透视表、共享字符串等其余部件原样保留
  - 不依赖任何引擎在线，也不会与宿主预览抢文件锁

适用场景：往已有表格里**补说明行 / 补单元格 / 追加结论**这类确定性写入。
不适合：批量重算、格式调整、透视与图表——那些仍应走表格专用通道。

用法:
    python patch_cells.py 文件.xlsx --set '{"作答区!A80": "说明文字", "命题!F5": 123}'
    python patch_cells.py 文件.xlsx --spec patch.json
    python patch_cells.py 文件.xlsx --spec patch.json --dry-run
    python patch_cells.py 文件.xlsx --spec patch.json --overwrite   # 允许覆盖已有单元格

spec.json 形如:
    {"作答区!A80": "文本", "作答区!B80": 86, "命题!E3": true}

行为约定:
  - 默认**拒绝覆盖已有非空单元格**（防止误伤），需要覆盖时显式加 --overwrite
  - 行/单元格不存在时按 Excel 的排序规则自动创建
  - 自动把该工作表的 <dimension> 扩展到新范围
  - 默认先备份为 文件.xlsx.bak
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

CELL_RE = re.compile(r'<c\s+r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)
ROW_RE = re.compile(r'<row\s+([^>]*?)(/>|>(.*?)</row>)', re.S)


def col_to_idx(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def idx_to_col(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def sheet_xml_map(z: zipfile.ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {r.get("Id"): r.get("Target", "") for r in rels}
    out = {}
    for sh in wb.iter(NS_MAIN + "sheet"):
        t = rid2t.get(sh.get(NS_REL + "id"), "").lstrip("/")
        out[sh.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def add_shared_string(parts: dict[str, bytes], text: str) -> int:
    """返回该文本在 sharedStrings 中的索引（已存在则复用）。"""
    key = "xl/sharedStrings.xml"
    ss = parts[key].decode("utf-8")
    body = re.search(r"<sst[^>]*>(.*)</sst>", ss, re.S).group(1)
    items = re.findall(r"<si>.*?</si>|<si/>", body, re.S)

    want = "<si><t>%s</t></si>" % esc(text)
    if text.strip() != text:
        want = '<si><t xml:space="preserve">%s</t></si>' % esc(text)
    # 已存在则复用（含 <si><t/></si> 的空串特例）
    for i, it in enumerate(items):
        m = re.fullmatch(r"<si><t(?:\s[^>]*)?>(.*?)</t></si>", it, re.S)
        if m and m.group(1) == esc(text):
            return i
    idx = len(items)
    parts[key] = ss.replace("</sst>", want + "</sst>").encode("utf-8")
    parts[key] = re.sub(rb'count="(\d+)"',
                        lambda m: b'count="%d"' % (int(m.group(1)) + 1), parts[key], count=1)
    parts[key] = re.sub(rb'uniqueCount="(\d+)"',
                        lambda m: b'uniqueCount="%d"' % (int(m.group(1)) + 1), parts[key], count=1)
    return idx


def make_cell(ref_col: str, row: int, value, ss_idx, inline: bool = False) -> str:
    ref = "%s%d" % (ref_col, row)
    if isinstance(value, bool):
        return '<c r="%s" t="b"><v>%d</v></c>' % (ref, 1 if value else 0)
    if isinstance(value, (int, float)):
        return '<c r="%s"><v>%s</v></c>' % (ref, repr(value) if isinstance(value, float) else value)
    if inline or ss_idx is None:
        # 内联字符串：不依赖共享字符串表。注意 openpyxl 生成的 xlsx **没有**
        # xl/sharedStrings.xml（它本来就写内联字符串），此时必须走这条路。
        sp = ' xml:space="preserve"' if (value != str(value).strip()) else ""
        return '<c r="%s" t="inlineStr"><is><t%s>%s</t></is></c>' % (ref, sp, esc(str(value)))
    return '<c r="%s" t="s"><v>%d</v></c>' % (ref, ss_idx)


def patch_sheet(xml: str, edits: dict[tuple[int, str], object], overwrite: bool,
                build_cell) -> tuple[str, int, list[str]]:
    """edits: {(row, col_letters): 原始值}；build_cell(row, col, value) 惰性构造单元格 XML。

    build_cell 必须**惰性调用**——因为写字符串值会往共享字符串表里追加条目，
    只有在确定这个单元格真的会被写入时才该追加，否则被拒绝的写入会留下无人引用的
    孤儿字符串（共享字符串计数虚高）。
    """
    problems: list[str] = []
    applied = 0

    for (row, col), value in edits.items():
        row_m = None
        for m in ROW_RE.finditer(xml):
            if re.search(r'\br="%d"' % row, m.group(1) or ""):
                row_m = m
                break
        if row_m is None:
            # 新建整行：必须把单元格包进 <row>，否则会成为 <sheetData> 的非法直接子节点
            cell_xml = build_cell(row, col, value)
            ins = None
            for m in ROW_RE.finditer(xml):
                mm = re.search(r'\br="(\d+)"', m.group(1) or "")
                if mm and int(mm.group(1)) < row:
                    ins = m.end()
            if ins is None:
                ins = xml.index(">", xml.index("<sheetData")) + 1
            new_row = '<row r="%d" spans="1:%d">%s</row>' % (row, col_to_idx(col) + 1, cell_xml)
            xml = xml[:ins] + new_row + xml[ins:]
            applied += 1
            continue

        body = row_m.group(3) or ""       # 自闭合行为 None
        cells = list(CELL_RE.finditer(body))
        target_idx = col_to_idx(col)
        existing = next((c for c in cells if c.group(1) == col), None)
        if existing is not None:
            if not overwrite:
                problems.append("%s%d 已存在内容，未覆盖（需要覆盖请加 --overwrite）" % (col, row))
                continue
            new_body = body[:existing.start()] + build_cell(row, col, value) + body[existing.end():]
        else:
            pos = None
            for c in cells:
                if col_to_idx(c.group(1)) > target_idx:
                    pos = c.start()
                    break
            cxml = build_cell(row, col, value)
            new_body = body[:pos] + cxml + body[pos:] if pos is not None else body + cxml
        xml = xml[:row_m.start(3)] + new_body + xml[row_m.end(3):]
        applied += 1

    return xml, applied, problems


def bump_dimension(xml: str, rows: list[int], cols: list[str]) -> str:
    m = re.search(r'<dimension ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"/>', xml)
    if not m:
        return xml
    old_end_r = int(m.group(4) or m.group(2))
    old_end_c = m.group(3) or m.group(1)
    new_r = max(old_end_r, *rows)
    new_c = max([old_end_c] + cols, key=col_to_idx)
    m2 = re.search(r"<dimension ref=\"[^\"]*\"/>", xml)
    return xml[:m2.start()] + '<dimension ref="A1:%s%d"/>' % (new_c, new_r) + xml[m2.end():]


def main() -> int:
    ap = argparse.ArgumentParser(description="直接往 xlsx 单元格写值（不经过表格引擎）")
    ap.add_argument("xlsx")
    ap.add_argument("--set", help="内联 JSON，形如 '{\"表名!A80\": \"文本\"}'")
    ap.add_argument("--spec", help="JSON 文件路径")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    a = ap.parse_args()

    spec = {}
    if a.spec:
        if not os.path.isfile(a.spec):
            sys.exit("spec 文件不存在: %s" % a.spec)
        try:
            spec.update(json.load(open(a.spec, encoding="utf-8")))
        except Exception as e:                             # noqa: BLE001
            sys.exit("spec 文件不是合法 JSON（%s）: %s —— %s" % (a.spec, type(e).__name__, e))
    if a.set:
        try:
            spec.update(json.loads(a.set))
        except Exception as e:                             # noqa: BLE001
            sys.exit("--set 不是合法 JSON（注意 Windows 下请用双引号包裹、内部用 \\\" 转义，"
                     "或改用 --spec 传文件）: %s" % e)
    if not spec:
        sys.exit("没有要写入的内容：请用 --set 或 --spec 指定")

    src = a.xlsx
    if not os.path.isfile(src):
        sys.exit("文件不存在: %s" % src)
    if not zipfile.is_zipfile(src):
        sys.exit("不是合法的 xlsx（zip 打不开）: %s" % src)

    z = zipfile.ZipFile(src)
    parts = {i.filename: z.read(i.filename) for i in z.infolist()}
    infos = z.infolist()
    smap = sheet_xml_map(z)
    z.close()
    by_sheet: dict[str, dict] = {}
    for ref, val in spec.items():
        sheet, _, coord = ref.partition("!")
        if not sheet or not coord:
            sys.exit("单元格引用格式应为 '工作表名!A1': %s" % ref)
        m = re.fullmatch(r"([A-Z]+)(\d+)", coord.upper())
        if not m:
            sys.exit("无法解析坐标: %s" % ref)
        by_sheet.setdefault(sheet, {})[(int(m.group(2)), m.group(1))] = val

    total, all_problems = 0, []
    for sheet, cells in by_sheet.items():
        if sheet not in smap:
            sys.exit("工作表不存在: %s（现有: %s）" % (sheet, list(smap)))
        xml = parts[smap[sheet]].decode("utf-8")

        def build_cell(row, col, value, _parts=parts):
            """惰性构造：只有确定要写入时才会往共享字符串表追加条目。"""
            if isinstance(value, (int, float, bool)):
                return make_cell(col, row, value, None)
            if "xl/sharedStrings.xml" not in _parts:
                # openpyxl 等工具生成的 xlsx 没有共享字符串表 —— 改用内联字符串
                return make_cell(col, row, value, None, inline=True)
            return make_cell(col, row, value, add_shared_string(_parts, str(value)))

        xml, n, probs = patch_sheet(xml, cells, a.overwrite, build_cell)
        if n:
            xml = bump_dimension(xml, [r for r, _ in cells], [c for _, c in cells])
        parts[smap[sheet]] = xml.encode("utf-8")
        total += n
        all_problems += probs
        print("  [%s] 写入 %d 个单元格" % (sheet, n))

    for p in all_problems:
        print("  ⚠ " + p)
    print("合计写入 %d 个单元格" % total)
    if a.dry_run:
        print("--dry-run：未落盘（改动只存在于内存，磁盘文件完全未动）")
        return 0
    if total == 0:
        print("没有任何单元格被写入，文件未改动。")
        return 1

    if not a.no_backup:
        shutil.copy2(src, src + ".bak")
        print("已备份 -> %s.bak" % src)

    tmp = src + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for i in infos:
            out.writestr(i, parts[i.filename])

    for attempt in range(5):
        try:
            os.replace(tmp, src)
            print("已写回 -> %s" % src)
            break
        except PermissionError:
            print("第 %d 次落盘被占用，等待 3 秒…（先 close_file / 关掉打开该文件的程序）" % (attempt + 1))
            time.sleep(3)
    else:
        print("落盘失败，修补结果暂存于: %s" % tmp)
        return 2

    # 复验：结构合法 + zip 完整 + 目标单元格实际可读；任一项失败自动回滚备份
    import openpyxl
    issues = []
    zz = zipfile.ZipFile(src)
    if zz.testzip() is not None:
        issues.append("zip 损坏")
    for m in re.finditer(r"</row>(\s*)(<c\s)", "".join(
            zz.read(n).decode("utf-8", "ignore") for n in zz.namelist()
            if n.startswith("xl/worksheets/") and n.endswith(".xml"))):
        issues.append("存在脱离 <row> 的单元格节点（结构非法）")
        break
    wb = openpyxl.load_workbook(src, data_only=True)
    for sheet, cells in by_sheet.items():
        for (row, col), want in cells.items():
            got = wb[sheet]["%s%d" % (col, row)].value
            ok = (abs(float(got) - float(want)) < 1e-9) if isinstance(want, (int, float)) \
                and not isinstance(want, bool) and isinstance(got, (int, float)) \
                else (str(got) == str(want))
            print("  复验 %s!%s%d = %r  %s" % (sheet, col, row, str(got)[:60], "OK" if ok else "**不符**"))
            if not ok:
                issues.append("%s!%s%d 写入后读回不符（期望 %r，实际 %r）"
                              % (sheet, col, row, want, got))

    if issues:
        print("复验未通过：%s" % issues)
        if not a.no_backup:
            shutil.copy2(src + ".bak", src)
            print("已自动回滚到备份 %s.bak" % src)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
