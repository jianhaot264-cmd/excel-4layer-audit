#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
不变量（property-based）与边界条件检查。

独立重算解决"算得对不对"，本脚本解决"逻辑讲不讲得通"和"边界会不会翻车"：
  · 排名是不是完整的 1..N（有没有断号/重号）
  · 降序区是不是真的非递增
  · 各分组之和是不是等于总计（配平检查）
  · **结论区的 Top-K 是不是真的是全量池里的前 K 名** ← 端到端逻辑不变量，最值钱
  · 有没有并列卡在第 K 名边界（此时"前三"没有唯一解，必须声明 tie-break 规则）
  · 有没有硬编码数字混在结论区（绕过了公式，改数据不跟着变）

用法:
    # A. 按 spec 检查（推荐）
    python invariant_check.py 文件.xlsx --spec invariants.json

    # B. 纯自动边界体检（不需要 spec）
    python invariant_check.py 文件.xlsx --auto-edge --source 数据原表

    # 两者可一起用
    python invariant_check.py 文件.xlsx --spec inv.json --auto-edge --source 数据原表

spec 支持的 check 类型:
    {"type": "rank_contiguous", "ref": "作答区!E7:E16"}
    {"type": "sorted_desc",     "ref": "作答区!B52:B61"}
    {"type": "sum_equals",      "range": "作答区!B52:B61", "target": "作答区!B7:B16"}
                                # range 与 target 都可为区间；无合计单元格时用"区间配平"
    {"type": "unique",          "ref": "作答区!A7:A16"}
    {"type": "non_negative",    "ref": "作答区!B7:B16"}
    {"type": "top_k_match",     "answer": "作答区!B20:B22",
                                "pool_labels": "作答区!A7:A16",
                                "pool_values": "作答区!C7:C16", "k": 3}
    {"type": "no_hardcoded",    "sheets": ["命题", "作答区"], "allow": ["作答区!A20:A22"]}

退出码: 0 = 全部通过；1 = 有 FAIL。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter

FAILS: list[str] = []
WARNS: list[str] = []
PASSES: list[str] = []


def ref_range(sheet, ref: str):
    """'A7:E16' -> (r1,c1,r2,c2)，1-based 闭区间；单点引用也支持。"""
    coord = ref.split("!")[-1].replace("$", "")
    if ":" in coord:
        a, b = coord.split(":")
    else:
        a = b = coord

    def parse(s):
        i = 0
        while i < len(s) and s[i].isalpha():
            i += 1
        col = 0
        for ch in s[:i]:
            col = col * 26 + (ord(ch.upper()) - 64)
        return int(s[i:]), col
    r1, c1 = parse(a)
    r2, c2 = parse(b)
    return r1, c1, r2, c2


def read_col(wb, ref: str) -> list:
    sheet, _, _ = ref.partition("!")
    r1, c1, r2, c2 = ref_range(sheet, ref)
    ws = wb[sheet]
    return [ws.cell(r, c1).value for r in range(r1, r2 + 1)]


def read_pairs(wb_lab, wb_val, labels_ref: str, values_ref: str):
    return list(zip(read_col(wb_lab, labels_ref), read_col(wb_val, values_ref)))


def is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ------------------------------------------------------------------ 各类检查
def c_rank_contiguous(wb, spec):
    vals = [v for v in read_col(wb, spec["ref"]) if v is not None]
    nums = [int(v) for v in vals if is_num(v)]
    if not nums:
        return FAILS.append("rank_contiguous %s: 区域内无数值" % spec["ref"])
    uniq = sorted(set(nums))
    expect = list(range(1, max(nums) + 1))
    if uniq != expect:
        FAILS.append("rank_contiguous %s: 排名不连续，实际 %s，应为 1..%d"
                     % (spec["ref"], uniq, max(nums)))
    elif len(nums) != len(uniq):
        WARNS.append("rank_contiguous %s: 存在并列名次（%d 个值 / %d 个不同名次）"
                     % (spec["ref"], len(nums), len(uniq)))
    else:
        PASSES.append("rank_contiguous %s: 1..%d 完整且无并列" % (spec["ref"], max(nums)))


def c_sorted_desc(wb, spec):
    vals = [v for v in read_col(wb, spec["ref"]) if is_num(v)]
    bad = [(i, vals[i], vals[i + 1]) for i in range(len(vals) - 1) if vals[i] < vals[i + 1] - 1e-9]
    if bad:
        FAILS.append("sorted_desc %s: 第 %s 处逆序（%s < %s）"
                     % (spec["ref"], [b[0] + 1 for b in bad][:5], bad[0][1], bad[0][2]))
    else:
        PASSES.append("sorted_desc %s: %d 项非递增" % (spec["ref"], len(vals)))


def c_sum_equals(wb, spec):
    """range 与 target 都可为区间：两边各自的数值之和必须相等。

    这样即使文件里没有"合计"单元格，也能做配平检查
    （例如"图表数据源之和"必须等于"明细表之和"）。
    """
    vals = [v for v in read_col(wb, spec["range"]) if is_num(v)]
    tgts = [v for v in read_col(wb, spec["target"]) if is_num(v)]
    if not tgts:
        return FAILS.append("sum_equals: 目标 %s 内无数值" % spec["target"])
    s, t = sum(vals), sum(tgts)
    if abs(s - t) > max(0.005, abs(t) * 1e-9):
        FAILS.append("sum_equals: %s 合计=%.6f ≠ %s 合计=%.6f（差 %.6f）"
                     % (spec["range"], s, spec["target"], t, s - t))
    else:
        PASSES.append("sum_equals: %s 合计=%.6f 与 %s 配平"
                      % (spec["range"], s, spec["target"]))


def c_unique(wb, spec):
    vals = [v for v in read_col(wb, spec["ref"]) if v is not None]
    dup = [k for k, n in Counter(vals).items() if n > 1]
    if dup:
        FAILS.append("unique %s: 存在重复值 %s" % (spec["ref"], dup[:6]))
    else:
        PASSES.append("unique %s: %d 项无重复" % (spec["ref"], len(vals)))


def c_non_negative(wb, spec):
    vals = [v for v in read_col(wb, spec["ref"]) if is_num(v)]
    neg = [v for v in vals if v < 0]
    if neg:
        FAILS.append("non_negative %s: 存在负值 %s" % (spec["ref"], neg[:6]))
    else:
        PASSES.append("non_negative %s: %d 项均非负" % (spec["ref"], len(vals)))


def c_top_k_match(wb, spec):
    """最值钱的端到端不变量：结论区的 Top-K 必须等于全量池按值降序的前 K。"""
    k = int(spec.get("k", 3))
    pairs = [(l, v) for l, v in read_pairs(wb, wb, spec["pool_labels"], spec["pool_values"]) if is_num(v)]
    pairs.sort(key=lambda x: -x[1])
    truth = pairs[:k]
    pool_labels = {l for l, _ in pairs}

    ans_raw = read_col(wb, spec["answer"])
    ans = [v for v in ans_raw if v is not None and v in pool_labels]
    skipped = [v for v in ans_raw if v is not None and v not in pool_labels]

    if len(ans) != k:
        FAILS.append("top_k_match: 结论区解析出 %d 项（期望 %d）；被忽略的非池内值: %s"
                     % (len(ans), k, skipped[:5]))
    if set(ans) != {l for l, _ in truth}:
        FAILS.append("top_k_match: 结论 %s ≠ 池内前 %d 名 %s"
                     % (ans, k, [l for l, _ in truth]))
    elif ans != [l for l, _ in truth]:
        WARNS.append("top_k_match: 集合正确但顺序不同，结论序 %s / 池内降序 %s"
                     % (ans, [l for l, _ in truth]))
    else:
        PASSES.append("top_k_match: 结论 %s 与池内前 %d 名完全一致" % (ans, k))

    # 关键边界：第 K 名与第 K+1 名并列 → "前三"没有唯一解
    if len(pairs) > k and abs(pairs[k - 1][1] - pairs[k][1]) < 1e-9:
        WARNS.append("top_k_match: **第 %d 名与第 %d 名并列**（值=%s）——"
                     "「前 %d」没有唯一解，交付物必须声明 tie-break 规则"
                     % (k, k + 1, pairs[k][1], k))


def c_no_hardcoded(wb_formula, spec):
    allow = set()
    for a in spec.get("allow", []):
        sh, _, co = a.partition("!")
        r1, c1, r2, c2 = ref_range(sh, a)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                allow.add("%s!%s" % (sh, wb_formula[sh].cell(r, c).coordinate))
    hits = []
    for name in spec.get("sheets", []):
        ws = wb_formula[name]
        for row in ws.iter_rows():
            for c in row:
                if is_num(c.value) and "%s!%s" % (name, c.coordinate) not in allow:
                    hits.append("%s!%s=%s" % (name, c.coordinate, c.value))
    if hits:
        WARNS.append("no_hardcoded: %d 处硬编码数字（非公式）——数据一变就不会跟着变，确认是有意为之: %s"
                     % (len(hits), hits[:10]))
    else:
        PASSES.append("no_hardcoded: 指定工作表内无硬编码数字")


CHECKERS = {
    "rank_contiguous": c_rank_contiguous, "sorted_desc": c_sorted_desc,
    "sum_equals": c_sum_equals, "unique": c_unique, "non_negative": c_non_negative,
    "top_k_match": c_top_k_match, "no_hardcoded": c_no_hardcoded,
}


# ------------------------------------------------------------------ 自动边界体检
def auto_edge(wb_val, wb_f, source: str) -> None:
    ws = wb_val[source]
    wsf = wb_f[source]
    rows = [r for r in ws.iter_rows(values_only=True)]
    if not rows:
        return FAILS.append("auto-edge: 工作表 %s 为空" % source)
    header, body = rows[0], [r for r in rows[1:] if any(v is not None for v in r)]

    print("\n【自动边界体检】数据源 %s：表头 %s，数据 %d 行" % (source, list(header), len(body)))

    # 1) 整行重复
    c = Counter(tuple(r) for r in body)
    dup = {k: v for k, v in c.items() if v > 1}
    if dup:
        rows_of = {}
        for i, r in enumerate(body, start=2):
            if tuple(r) in dup:
                rows_of.setdefault(tuple(r), []).append(i)
        WARNS.append("auto-edge: 存在 %d 组完全重复的行 —— 必须明确「按全量」还是「去重」，"
                     "并把两种结果都写进交付物。行号: %s"
                     % (len(dup), list(rows_of.values())[:4]))
    else:
        PASSES.append("auto-edge: 无完全重复行")

    # 2) 每列的数值/文本类型构成
    ncol = len(header)
    for ci in range(ncol):
        nums = [r[ci] for r in body if is_num(r[ci])]
        strs = [r[ci] for r in body if isinstance(r[ci], str) and str(r[ci]).strip()]
        blanks = [i for i, r in enumerate(body, start=2) if r[ci] is None]
        name = header[ci] if ci < len(header) else "第%d列" % (ci + 1)
        if nums and strs and len(nums) >= 2 and len(strs) >= 2:
            ex = sorted({v for v in nums})[:4]
            WARNS.append("auto-edge: 列「%s」同时存在数值与文本（数值样例 %s）——"
                         "按该列筛选/分组会静默漏行，须确认是否影响结论" % (name, ex))
        if blanks:
            WARNS.append("auto-edge: 列「%s」有 %d 个空白单元格（行 %s）"
                         % (name, len(blanks), blanks[:6]))
        if nums and min(nums) < 0:
            WARNS.append("auto-edge: 列「%s」存在负值（最小 %s）——确认是否符合业务含义"
                         % (name, min(nums)))
        if nums and max(nums) == 0:
            WARNS.append("auto-edge: 列「%s」全部为 0" % name)

    # 3) 分组数 / 单成员分组 / 全零分组
    for ci, name in enumerate(header):
        vals = [r[ci] for r in body if isinstance(r[ci], str) and str(r[ci]).strip()]
        cnt = Counter(vals)
        if 2 <= len(cnt) <= 30 and len(vals) >= len(body) * 0.8:
            singles = [k for k, n in cnt.items() if n == 1]
            if singles:
                WARNS.append("auto-edge: 分组列「%s」有 %d 个只出现 1 次的取值（如 %s）——"
                             "这类组的「日均」类指标极度不稳，分母口径必须写明"
                             % (name, len(singles), singles[:5]))

    # 4) 日期列跨度
    import datetime as _dt
    for ci, name in enumerate(header):
        ds = [r[ci] for r in body if isinstance(r[ci], (_dt.datetime, _dt.date))]
        if ds and len(ds) >= len(body) * 0.8:
            uniq = sorted(set(ds))
            PASSES.append("auto-edge: 日期列「%s」跨度 %s ~ %s，共 %d 个不同日期"
                          % (name, uniq[0], uniq[-1], len(uniq)))
            # 日期是否有断档（缺日）——影响"日均"类分母口径
            gaps = [(uniq[i + 1] - uniq[i]).days for i in range(len(uniq) - 1)]
            missing = [(uniq[i], uniq[i + 1]) for i, g in enumerate(gaps) if g > 1]
            if missing:
                WARNS.append("auto-edge: 日期列「%s」存在断档 %s —— 计算「日均」时分母用"
                             "跨度天数还是实际有数据的天数，必须写明口径"
                             % (name, missing[:4]))


def main() -> int:
    ap = argparse.ArgumentParser(description="不变量与边界条件检查")
    ap.add_argument("xlsx")
    ap.add_argument("--spec", help="不变量 spec JSON")
    ap.add_argument("--auto-edge", action="store_true", help="对数据源做自动边界体检")
    ap.add_argument("--source", default="数据原表", help="数据源工作表名（配合 --auto-edge）")
    ap.add_argument("--json", help="把结果写成 JSON")
    a = ap.parse_args()

    if not os.path.isfile(a.xlsx):
        sys.exit("文件不存在: %s" % a.xlsx)
    if a.spec and not os.path.isfile(a.spec):
        sys.exit("spec 文件不存在: %s" % a.spec)

    import openpyxl
    wb = openpyxl.load_workbook(a.xlsx, data_only=True)
    wb_f = openpyxl.load_workbook(a.xlsx)

    print("=" * 78)
    print("不变量与边界检查: %s" % a.xlsx)
    print("=" * 78)

    if a.spec:
        spec = json.load(open(a.spec, encoding="utf-8"))
        if "checks" in spec:
            items = spec["checks"]
        else:
            items = []
            if spec.get("source"):
                a.auto_edge = True
                a.source = spec["source"].get("sheet", a.source)
        for item in items:
            fn = CHECKERS.get(item.get("type"))
            if not fn:
                FAILS.append("未知检查类型: %r" % item.get("type"))
                continue
            try:
                if item["type"] == "no_hardcoded":
                    fn(wb_f, item)
                else:
                    fn(wb, item)
            except Exception as e:                      # noqa: BLE001
                FAILS.append("检查 %s 执行异常: %s: %s" % (item, type(e).__name__, e))

    if a.auto_edge:
        if a.source not in wb.sheetnames:
            FAILS.append("数据源工作表不存在: %s" % a.source)
        else:
            auto_edge(wb, wb_f, a.source)

    for f in FAILS:
        print("\n✘ FAIL  " + f)
    for w in WARNS:
        print("\n▲ WARN  " + w)
    for p in PASSES:
        print("\n✔ PASS  " + p)

    print()
    print("=" * 78)
    print("结论: FAIL %d 项 / WARN %d 项 / PASS %d 项" % (len(FAILS), len(WARNS), len(PASSES)))
    print("=" * 78)
    if a.json:
        json.dump({"fails": FAILS, "warns": WARNS, "passes": PASSES,
                   "ok": not FAILS},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
