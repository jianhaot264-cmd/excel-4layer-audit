#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx 多引擎验证（Microsoft Word / WPS 文字）。

定位：本脚本是 recalc_verify.py 的**文档类孪生**。表格靠 <v> 缓存，文档靠"域（Field）结果"
和"表格公式"——两者有同一个盲区：**读到的是上一次保存方写进文件的结果，不代表 Word/WPS
重新计算后会得到什么**。所以同样需要"另一个引擎强制更新域 + 跨引擎比对"。

它能做什么：
    ① 用两个引擎分别打开 docx，强制更新全部域（含表格公式 =SUM(ABOVE) 这类）
    ② 读出：非空段落文本 / 各表格单元格文本 / 各域的计算结果
    ③ 跨引擎逐项比对（N-version），差异即暴露引擎实现差异
    ④ 与期望值清单比对
    ⑤ 只读打开、不保存，事后可比对文件哈希

引擎注册表：
    word : ProgID "Word.Application"   → Microsoft Word
    wps  : ProgID "KWPS.Application"   → WPS 文字
    ⚠️ 与表格引擎同理：WPS 会把 Name 自报成 "Microsoft Word"，判定引擎必须靠
       ProgID + 安装路径，不能只看 Name（见 references/pitfalls.md P13）。

引用写法（--expect 用）：
    "para:3"        -> 第 3 个非空段落
    "table:1:R2C3"  -> 第 1 个表格的第 2 行第 3 列
    "field:2"       -> 第 2 个域的计算结果文本

用法:
    python verify_docx.py 文件.docx --list-engines
    python verify_docx.py 文件.docx --engine word,wps
    python verify_docx.py 文件.docx --engine word,wps --expect expected.json --json r.json

退出码: 0 = 全一致；1 = 有差异；2 = 无可用引擎。
"""
from __future__ import annotations

import argparse
import gc
import json
import os
import re
import sys

ENGINES: dict[str, tuple[str, tuple[str, ...]]] = {
    "word": ("Word.Application", ("microsoft office",)),
    "wps": ("KWPS.Application", ("kingsoft", "wps office")),
}

WD_DO_NOT_SAVE = 0


def clean(s) -> str:
    """清掉 Word 单元格/段落文本里的控制符（单元格标记 \\x07、段落符 、软换行 \\x0b）。"""
    if s is None:
        return ""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(s)).replace("\r", "").strip()


def probe_engine(key: str) -> dict | None:
    progid, _ = ENGINES[key]
    try:
        import win32com.client as win32
        app = win32.DispatchEx(progid)
        info = {"key": key, "progid": progid,
                "version": str(getattr(app, "Version", "?")),
                "name": str(getattr(app, "Name", "?")),
                "path": str(getattr(app, "Path", "?"))}
        try:
            app.Quit(WD_DO_NOT_SAVE)
        except Exception:                                  # noqa: BLE001
            try:
                app.Quit()
            except Exception:                              # noqa: BLE001
                pass
        del app
        gc.collect()
        return info
    except Exception:                                      # noqa: BLE001
        return None


def engine_path_warning(key: str, info: dict) -> str | None:
    _, keywords = ENGINES[key]
    path = info["path"].lower().replace("\\", "/")
    if not any(k in path for k in keywords):
        return ("警告：%s 解析到 %s，但安装路径 %s 不含预期关键字 %s —— 可能被另一引擎接管，"
                "结论请谨慎采信" % (key, info["progid"], info["path"], list(keywords)))
    return None


def read_with_engine(key: str, path: str) -> dict:
    """用指定引擎打开并强制更新全部域，读出段落 / 表格 / 域结果。"""
    import win32com.client as win32
    import pythoncom

    progid, _ = ENGINES[key]
    pythoncom.CoInitialize()
    app = doc = None
    out = {"paragraphs": [], "tables": [], "fields": []}
    try:
        app = win32.DispatchEx(progid)
        for attr, val in (("Visible", False), ("DisplayAlerts", 0)):
            try:
                setattr(app, attr, val)
            except Exception:                              # noqa: BLE001
                pass
        try:
            app.Options.ConfirmConversions = False
        except Exception:                                  # noqa: BLE001
            pass

        doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                 AddToRecentFiles=False, ConfirmConversions=False, Visible=False)

        # ---- 强制更新全部域（这是文档版的"重算"） ----
        try:
            doc.Fields.Update()
        except Exception:                                  # noqa: BLE001
            pass
        try:
            for i in range(1, doc.Tables.Count + 1):
                doc.Tables(i).Range.Fields.Update()         # 表格公式 =SUM(ABOVE) 等
        except Exception:                                  # noqa: BLE001
            pass

        # ---- 段落 ----
        for i in range(1, doc.Paragraphs.Count + 1):
            t = clean(doc.Paragraphs(i).Range.Text)
            if t:
                out["paragraphs"].append(t)

        # ---- 表格 ----
        for ti in range(1, doc.Tables.Count + 1):
            tb = doc.Tables(ti)
            try:
                nrow, ncol = int(tb.Rows.Count), int(tb.Columns.Count)
            except Exception:                              # noqa: BLE001
                nrow = ncol = 0
            grid = []
            for r in range(1, nrow + 1):
                line = []
                for c in range(1, ncol + 1):
                    try:
                        line.append(clean(tb.Cell(r, c).Range.Text))     # 合并单元格会抛异常
                    except Exception:                                # noqa: BLE001
                        line.append("<merged>")
                grid.append(line)
            out["tables"].append({"rows": nrow, "cols": ncol, "cells": grid})

        # ---- 域（含表格公式） ----
        # 注意：Word 对象模型里 Document.Fields 已包含表格内的域，
        #       再遍历 Tables(i).Range.Fields 会重复收集，这里只取 Document.Fields。
        for i in range(1, doc.Fields.Count + 1):
            f = doc.Fields(i)
            try:
                code = clean(f.Code.Text)
            except Exception:                              # noqa: BLE001
                code = "?"
            try:
                res = clean(f.Result.Text)
            except Exception:                              # noqa: BLE001
                res = "?"
            out["fields"].append({"code": code, "result": res})
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE)
        except Exception:                                  # noqa: BLE001
            pass
        try:
            if app is not None:
                app.Quit(WD_DO_NOT_SAVE)
        except Exception:                                  # noqa: BLE001
            pass
        del doc, app
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:                                  # noqa: BLE001
            pass
    return out


def resolve(book: dict, ref: str):
    """把 'table:1:R2C3' / 'para:3' / 'field:2' 解析成实际值。"""
    kind, _, rest = ref.partition(":")
    if kind == "para":
        i = int(rest) - 1
        return book["paragraphs"][i] if 0 <= i < len(book["paragraphs"]) else None
    if kind == "field":
        i = int(rest) - 1
        return book["fields"][i]["result"] if 0 <= i < len(book["fields"]) else None
    if kind == "table":
        m = re.fullmatch(r"(\d+):R(\d+)C(\d+)", rest)
        if not m:
            return None
        t, r, c = (int(x) for x in m.groups())
        try:
            return book["tables"][t - 1]["cells"][r - 1][c - 1]
        except IndexError:
            return None
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="docx 多引擎验证（Word + WPS 文字）")
    ap.add_argument("docx")
    ap.add_argument("--engine", default="auto", help="auto | word | wps | word,wps")
    ap.add_argument("--list-engines", action="store_true")
    ap.add_argument("--expect", help="期望值 JSON，引用形如 \"table:1:R2C3\"")
    ap.add_argument("--json", help="结果写成 JSON")
    ap.add_argument("--max-print", type=int, default=20)
    a = ap.parse_args()

    print("=" * 78)
    print("docx 多引擎验证: %s" % a.docx)
    print("=" * 78)

    print("\n【引擎探测】")
    avail = {}
    for key in ENGINES:
        info = probe_engine(key)
        if info:
            avail[key] = info
            print("  ✔ %-5s ProgID=%-22s 版本=%-5s 路径=%s"
                  % (key, info["progid"], info["version"], info["path"]))
            print("     （引擎自报名: %s）" % info["name"])
            w = engine_path_warning(key, info)
            if w:
                print("     ⚠ " + w)
        else:
            print("  ✘ %-5s 不可用（ProgID %s 未注册）" % (key, ENGINES[key][0]))
    if a.list_engines:
        return 0 if avail else 2
    if not avail:
        print("\n✘ 无可用引擎")
        return 2

    use = list(avail) if a.engine == "auto" else [e.strip() for e in a.engine.split(",") if e.strip()]
    bad = [e for e in use if e not in avail]
    if bad:
        print("\n✘ 指定引擎不可用: %s" % bad)
        return 2
    if not os.path.isfile(a.docx):
        sys.exit("文件不存在: %s" % a.docx)

    books, issues = {}, []
    for key in use:
        print("\n【%s 引擎读取】" % key)
        try:
            books[key] = read_with_engine(key, a.docx)
            b = books[key]
            print("  ✔ 段落 %d / 表格 %d / 域 %d"
                  % (len(b["paragraphs"]), len(b["tables"]), len(b["fields"])))
        except Exception as e:                             # noqa: BLE001
            issues.append("%s: %s: %s" % (key, type(e).__name__, e))
            print("  ✘ 失败：%s: %s" % (type(e).__name__, e))
    if not books:
        return 2

    all_ok = True

    # ---------- ① 跨引擎逐项比对 ----------
    print("\n" + "=" * 78)
    print("① 引擎之间逐项比对（N-version）")
    cross = []
    keys = list(books)
    if len(keys) < 2:
        print("  · 仅一个可用引擎，跳过（建议至少再用另一个引擎验一次）")
    for i in range(len(keys)):
        for j in range(i + 1, len(keys)):
            k1, k2 = keys[i], keys[j]
            diffs = []
            for sec in ("paragraphs", "tables", "fields"):
                v1, v2 = books[k1][sec], books[k2][sec]
                if len(v1) != len(v2):
                    diffs.append(("%s 数量" % sec, len(v1), len(v2)))
                    continue
                for idx, (x, y) in enumerate(zip(v1, v2)):
                    if x != y:
                        diffs.append(("%s[%d]" % (sec, idx + 1), x, y))
            cross.append({"a": k1, "b": k2, "diffs": diffs})
            print("  %s ↔ %s ：%s" % (k1, k2, "✔ 完全一致" if not diffs else "✘ %d 处不同" % len(diffs)))
            for d in diffs[:a.max_print]:
                print("      %s  %s=%s  %s=%s" % (d[0], k1, str(d[1])[:60], k2, str(d[2])[:60]))
            if diffs:
                all_ok = False

    # ---------- ② 域结果 ----------
    print("\n" + "=" * 78)
    print("② 域（含表格公式）计算结果")
    for key, b in books.items():
        if not b["fields"]:
            print("  %-5s → 无域" % key)
        for f in b["fields"][:a.max_print]:
            print("  %-5s %-28s = %s" % (key, f["code"][:28], f["result"]))

    # ---------- ③ 期望值 ----------
    expect_bad = {}
    if a.expect:
        try:
            spec = json.load(open(a.expect, encoding="utf-8"))
        except Exception as e:                         # noqa: BLE001
            print("\n✘ 期望值清单读取失败：%s: %s（请用绝对路径）" % (type(e).__name__, e))
            return 1
        print("\n" + "=" * 78)
        print("③ 引擎读值 vs 期望值（%d 项）" % len(spec))
        for key, b in books.items():
            bads = []
            for ref, want in spec.items():
                got = resolve(b, ref)
                if clean(got) != clean(want):
                    bads.append((ref, want, got))
            expect_bad[key] = bads
            print("  %-5s → %s" % (key, "✔ 全部命中" if not bads else "✘ %d 项不符" % len(bads)))
            for bd in bads[:a.max_print]:
                print("      %s 期望=%r 实际=%r" % bd)
            if bads:
                all_ok = False

    print("\n" + "=" * 78)
    print("结论: %s" % ("✔ 多引擎验证通过（%s）" % "+".join(books) if all_ok else "✘ 存在差异"))
    print("=" * 78)

    if a.json:
        json.dump({"engines": avail, "used": keys, "issues": issues,
                   "cross": cross, "expect_mismatch": expect_bad,
                   "fields": {k: books[k]["fields"] for k in books},
                   "tables": {k: [{"rows": t["rows"], "cols": t["cols"]} for t in books[k]["tables"]]
                              for k in books},
                   "ok": all_ok},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
