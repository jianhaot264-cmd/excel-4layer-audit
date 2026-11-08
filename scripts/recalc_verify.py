#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第二重引擎重算验证（多引擎：Excel / WPS）。

为什么需要它——这是纯 openpyxl 审计的**根本盲区**：
    openpyxl 的 data_only=True 读的是 xlsx 里**缓存的**公式结果（<v> 标签）。
    缓存由"最后保存这个文件的那个程序"写入。如果那个程序的引擎算错了、
    或者缓存是陈旧值，那么"读缓存 → 比对期望值"会全部通过，而用户用 Excel
    打开看到的却是另一组数字。这类错误**任何基于缓存值的检查都发现不了**。

本脚本用**真实办公软件引擎**强制全量重算，然后：
    ① 缓存值   vs  各引擎重算值   → 找出陈旧/错误的缓存（关键）
    ② 引擎重算 vs  期望值         → 权威结论校验
    ③ 引擎之间互相比对            → N-version 交叉验证（引擎差异会在这里暴露）
    ④ 重算后的错误值扫描

引擎可用性探测：
    Excel : ProgID "Excel.Application"
    WPS   : ProgID "KET.Application"（WPS 表格；注意 WPS 会把 Name 伪装成
            "Microsoft Excel"，所以要靠 ProgID + 安装路径 + 版本号共同判定）

用法:
    python recalc_verify.py 文件.xlsx                                   # 自动探测全部可用引擎
    python recalc_verify.py 文件.xlsx --engine wps
    python recalc_verify.py 文件.xlsx --engine excel,wps --expect expected.json
    python recalc_verify.py 文件.xlsx --json report.json
    python recalc_verify.py 文件.xlsx --sheet 命题 --sheet 作答区        # 只看指定表（更快）
    python recalc_verify.py 文件.xlsx --list-engines                    # 只探测引擎

退出码: 0 = 全部一致；1 = 发现不一致；2 = 无可用引擎。
"""
from __future__ import annotations

import argparse
import datetime as _dt
import gc
import json
import os
import sys

TOL = 1e-9
REL = 1e-7

# 引擎注册表：key -> (ProgID, 该引擎安装路径可能包含的关键字元组)
# 注意：WPS 会把 Name 自报名成 "Microsoft Excel"（兼容层行为），所以判定引擎不能
#       只看 Name，必须靠 ProgID + 安装路径共同识别。
ENGINES: dict[str, tuple[str, tuple[str, ...]]] = {
    "excel": ("Excel.Application", ("microsoft office",)),
    "wps": ("KET.Application", ("kingsoft", "wps office")),
}

XL_ERR = {
    -2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?",
    -2146826288: "#NULL!", -2146826252: "#NUM!", -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def norm(v):
    """把引擎 / openpyxl 的值归一化，便于跨引擎比较。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v != v:               # NaN
            return "NaN"
        if isinstance(v, int) and v in XL_ERR:
            return XL_ERR[v]
        if isinstance(v, float) and abs(v) >= 2_000_000_000:
            return XL_ERR.get(int(v), "XLERR(%d)" % int(v))
        return round(float(v), 10)
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.isoformat()
    if isinstance(v, str):
        s = v.strip()
        return s
    return str(v)


def same(a, b) -> bool:
    a, b = norm(a), norm(b)
    if a is None and b is None:
        return True
    if isinstance(a, str) and a in XL_ERR.values():
        return a == b
    if isinstance(a, float) and isinstance(b, float):
        if a == b:
            return True
        return abs(a - b) <= max(TOL, REL * max(abs(a), abs(b)))
    return a == b


def probe_engine(key: str) -> dict | None:
    """探测某引擎是否可用，返回 {'progid','version','name','path'} 或 None。"""
    progid, _ = ENGINES[key]
    try:
        import win32com.client as win32
        app = win32.DispatchEx(progid)
        info = {"key": key, "progid": progid,
                "version": str(getattr(app, "Version", "?")),
                "name": str(getattr(app, "Name", "?")),
                "path": str(getattr(app, "Path", "?"))}
        try:
            app.Quit()
        except Exception:                                  # noqa: BLE001
            pass
        del app
        gc.collect()
        return info
    except Exception:                                      # noqa: BLE001
        return None


def notify_wrong_engine(key: str, info: dict) -> str | None:
    """WPS 会把 Name 伪装成 Microsoft Excel；用安装路径反查是否挂错引擎。"""
    _, keywords = ENGINES[key]
    path = info["path"].lower().replace("\\", "/")
    if not any(k in path for k in keywords):
        return ("警告：%s 解析到 ProgID %s，但安装路径为 %s —— 该路径不含预期关键字 %s，"
                "可能被另一引擎接管，结论请谨慎采信"
                % (key, info["progid"], info["path"], list(keywords)))
    return None


def normalize_cell_value(v):
    """把引擎返回的负 HRESULT 错误码归一化为 Excel 错误名字符串。

    COM 读取公式错误单元格时，不同引擎/不同调用路径可能返回：
      - 错误字符串 "#REF!" / "#VALUE!" / ...
      - 或负的 HRESULT 整数（实测 =#NUM! 为 -2146826252 = 0x800A07F4，低16位 2036）
    不归一化的话，缓存比对列会显示成一串无意义的大负数，可读性差。
    """
    if not isinstance(v, int) or v >= 0:
        return v
    u = v & 0xFFFFFFFF
    # 只处理 0x800Axxxx 这一族（Excel 错误值编码）
    if (u >> 16) != 0x800A:
        return v
    code = u & 0xFFFF
    return _ERR_CODE.get(code, v)


# 低 16 位 → Excel 错误名（对应 xlErrXXX 常量）
_ERR_CODE = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
}

# 供文档/排查参考的 HRESULT 形态
_ERR_HRESULT_EXAMPLE = {k: "0x800A%04X" % (0x0700 | (k - 2000) + 0xD0) for k in _ERR_CODE}


def read_with_engine(key: str, path: str, sheets: list[str] | None) -> dict[str, dict]:
    """用指定引擎强制全量重算并读出全部单元格值。"""
    import win32com.client as win32
    import pythoncom

    progid, _ = ENGINES[key]
    pythoncom.CoInitialize()
    app = wb = None
    out: dict[str, dict] = {}
    try:
        app = win32.DispatchEx(progid)
        try:
            app.Visible = False
        except Exception:                                  # noqa: BLE001
            pass
        for attr, val in (("DisplayAlerts", False), ("AskToUpdateLinks", False),
                          ("EnableEvents", False), ("AutomationSecurity", 3),
                          ("Calculation", -4105)):         # xlCalculationAutomatic
            try:
                setattr(app, attr, val)
            except Exception:                              # noqa: BLE001
                pass
        wb = app.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=True,
                                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
        # 优先全量重建；旧引擎退化为 CalculateFull / Calculate
        for meth in ("CalculateFullRebuild", "CalculateFull", "Calculate"):
            try:
                getattr(app, meth)()
                break
            except Exception:                              # noqa: BLE001
                continue
        else:
            raise RuntimeError("该引擎不提供任何全量重算接口")

        names = sheets or [ws.Name for ws in wb.Worksheets]
        for name in names:
            ws = wb.Worksheets(name)
            ur = ws.UsedRange
            if ur is None or ur.Rows.Count == 0:
                out[name] = {}
                continue
            r0, c0 = int(ur.Row), int(ur.Column)
            nr, nc = int(ur.Rows.Count), int(ur.Columns.Count)
            vals = ur.Value
            if nr == 1 and nc == 1:
                vals = ((vals,),)
            cells = {}
            for i in range(nr):
                row = vals[i]
                if not isinstance(row, (tuple, list)):
                    row = (row,)
                for j in range(min(nc, len(row))):
                    v = row[j]
                    if v is None:
                        continue
                    v = normalize_cell_value(v)
                    n, s = c0 + j, ""
                    while n:
                        n, r = divmod(n - 1, 26)
                        s = chr(65 + r) + s
                    cells["%s%d" % (s, r0 + i)] = v
            out[name] = cells
    finally:
        for closer in (lambda: wb.Close(SaveChanges=False), lambda: app.Quit()):
            try:
                closer()
            except Exception:                              # noqa: BLE001
                pass
        del wb, app
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:                                  # noqa: BLE001
            pass
    return out


def scan_errors(book: dict[str, dict]):
    hits = []
    for name, cells in book.items():
        for coord, v in cells.items():
            if isinstance(v, str) and v.strip() in XL_ERR.values():
                hits.append((name, coord, v))
            elif isinstance(v, int) and v in XL_ERR:
                hits.append((name, coord, XL_ERR[v]))
    return hits


def main() -> int:
    ap = argparse.ArgumentParser(description="用真实办公软件引擎重算并校验（支持 Excel / WPS）")
    ap.add_argument("xlsx")
    ap.add_argument("--engine", default="auto",
                    help="auto | excel | wps | 逗号分隔的多个，如 excel,wps")
    ap.add_argument("--list-engines", action="store_true", help="只探测引擎可用性")
    ap.add_argument("--expect", help="期望值 JSON（{\"表!A1\": 值}）")
    ap.add_argument("--sheet", action="append", help="只检查该工作表，可多次传入")
    ap.add_argument("--json", help="把结果写成 JSON")
    ap.add_argument("--max-print", type=int, default=25)
    a = ap.parse_args()

    print("=" * 78)
    print("多引擎重算验证: %s" % a.xlsx)
    print("=" * 78)

    # ---------- 引擎探测 ----------
    avail = {}
    print("\n【引擎探测】")
    for key in ENGINES:
        info = probe_engine(key)
        if info:
            avail[key] = info
            warn = notify_wrong_engine(key, info)
            print("  ✔ %-6s ProgID=%-20s 版本=%-5s 路径=%s"
                  % (key, info["progid"], info["version"], info["path"]))
            print("     （引擎自报名: %s）" % info["name"])
            if warn:
                print("     ⚠ " + warn)
        else:
            print("  ✘ %-6s 不可用（ProgID %s 未注册）" % (key, ENGINES[key][0]))

    if a.list_engines:
        return 0 if avail else 2
    if not avail:
        print("\n✘ 无任何可用引擎。可用替代：pip install formulas（纯 Python 公式求值库）")
        return 2

    if a.engine == "auto":
        use = list(avail)
    else:
        use = [e.strip() for e in a.engine.split(",") if e.strip()]
        missing = [e for e in use if e not in avail]
        if missing:
            print("\n✘ 指定的引擎不可用: %s" % missing)
            return 2

    if not os.path.isfile(a.xlsx):
        sys.exit("文件不存在: %s" % a.xlsx)

    import openpyxl
    wb_cache = openpyxl.load_workbook(a.xlsx, data_only=True)
    wb_formula = openpyxl.load_workbook(a.xlsx)

    # ---------- 逐引擎重算 ----------
    books, engine_issues = {}, []
    for key in use:
        print("\n【%s 引擎重算】" % key)
        try:
            books[key] = read_with_engine(key, a.xlsx, a.sheet)
            print("  ✔ 完成全量重算")
        except Exception as e:                             # noqa: BLE001
            engine_issues.append("%s: %s: %s" % (key, type(e).__name__, e))
            print("  ✘ 失败：%s: %s" % (type(e).__name__, e))
    if not books:
        print("\n✘ 无引擎成功完成重算")
        return 2

    sheets = a.sheet or list(next(iter(books.values())).keys())
    all_ok = True

    # ---------- ① 缓存 vs 各引擎 ----------
    print("\n" + "=" * 78)
    print("① 缓存值 vs 引擎重算值")
    stale_by_engine, missing_by_engine = {}, {}
    n_formula = 0
    for key, book in books.items():
        stale, missing = [], []
        n_formula = 0
        for name in sheets:
            if name not in wb_formula.sheetnames:
                continue
            wf, wc, ex = wb_formula[name], wb_cache[name], book.get(name, {})
            for row in wf.iter_rows():
                for c in row:
                    if not (isinstance(c.value, str) and c.value.startswith("=")):
                        continue
                    n_formula += 1
                    cached = wc[c.coordinate].value
                    if cached is None:
                        # 区分「缓存算错」与「根本没有缓存」：后者是文件从未被真实引擎
                        # 保存过（如 openpyxl 直接写出的 xlsx），不是计算错误
                        missing.append((name, c.coordinate, ex.get(c.coordinate)))
                    elif not same(cached, ex.get(c.coordinate)):
                        stale.append((name, c.coordinate, cached, ex.get(c.coordinate)))
        stale_by_engine[key] = stale
        missing_by_engine[key] = missing
        tag = "✔ 完全一致" if not stale else "✘ %d 处不一致" % len(stale)
        print("  %-6s 检查 %d 个公式单元格 → %s" % (key, n_formula, tag))
        for s in stale[:a.max_print]:
            print("      %s!%s  缓存=%r  引擎=%r" % s)
        if stale:
            all_ok = False
        if missing:
            print("      ▲ 另有 %d 个公式**没有缓存值**（文件可能由脚本生成、未经真实引擎保存），"
                  "已把引擎重算值当作真值处理" % len(missing))
    no_cache = any(missing_by_engine.values())

    # ---------- ② 引擎之间交叉比对（N-version） ----------
    cross = []
    keys = list(books)
    print("\n" + "=" * 78)
    print("② 引擎之间交叉比对（N-version 交叉验证）")
    if len(keys) < 2:
        print("  · 仅一个可用引擎，跳过（铁律：至少要有一个与『最终读者』不同的引擎）")
    else:
        for i in range(len(keys)):
            for j in range(i + 1, len(keys)):
                k1, k2 = keys[i], keys[j]
                diff = []
                for name in sheets:
                    b1, b2 = books[k1].get(name, {}), books[k2].get(name, {})
                    for coord in set(b1) | set(b2):
                        if not same(b1.get(coord), b2.get(coord)):
                            diff.append((name, coord, b1.get(coord), b2.get(coord)))
                cross.append((k1, k2, diff))
                tag = "✔ 完全一致" if not diff else "✘ %d 处不同" % len(diff)
                print("  %s ↔ %s ：%s" % (k1, k2, tag))
                for d in diff[:a.max_print]:
                    print("      %s!%s  %s=%r  %s=%r" % (d[0], d[1], k1, d[2], k2, d[3]))
                if diff:
                    all_ok = False

    # ---------- ③ 期望值 ----------
    expect_bad = {}
    if a.expect:
        spec = json.load(open(a.expect, encoding="utf-8"))
        print("\n" + "=" * 78)
        print("③ 引擎重算值 vs 期望值（%d 项）" % len(spec))
        for key, book in books.items():
            bad = []
            for ref, want in spec.items():
                sh, _, co = ref.partition("!")
                if not same(want, book.get(sh, {}).get(co)):
                    bad.append((ref, want, book.get(sh, {}).get(co)))
            expect_bad[key] = bad
            print("  %-6s → %s" % (key, "✔ 全部命中" if not bad else "✘ %d 项不符" % len(bad)))
            for b in bad[:a.max_print]:
                print("      %s 期望=%r 引擎=%r" % b)
            if bad:
                all_ok = False

    # ---------- ④ 错误值 ----------
    print("\n" + "=" * 78)
    print("④ 重算后的错误值扫描")
    errs = {}
    for key, book in books.items():
        e = scan_errors(book)
        errs[key] = e
        print("  %-6s → %s" % (key, "✔ 无错误值" if not e else "✘ %d 处: %s" % (len(e), e[:8])))
        if e:
            all_ok = False

    print("\n" + "=" * 78)
    print("结论: %s" % ("✔ 多引擎验证通过（%s）" % "+".join(books) if all_ok
                       else "✘ 存在不一致，必须回溯修正"))
    if no_cache and all_ok:
        print("     ⚠ 注意：本文件的部分公式没有缓存值 —— 它多半是由脚本直接生成的、")
        print("       尚未经过真实引擎保存。此时『读缓存』类检查是空的，必须以本脚本的重算结果为准。")
    if not all_ok:
        print("     提醒：任何修正之后必须重跑本脚本 —— 改一个单元格，之前结论即失效。")
    print("=" * 78)

    if a.json:
        json.dump({
            "engines_available": avail,
            "engines_used": list(books),
            "engine_issues": engine_issues,
            "formula_cells": n_formula,
            "stale_by_engine": stale_by_engine,
            "missing_cache_by_engine": missing_by_engine,
            "no_cache": no_cache,
            "cross_engine": [{"a": c[0], "b": c[1], "diff": c[2]} for c in cross],
            "expect_mismatch": expect_bad,
            "errors": errs,
            "ok": all_ok,
        }, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
