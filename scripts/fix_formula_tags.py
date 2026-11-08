#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
归一化 xlsx 内的公式标签。

背景：本地 editor_sdk 写出的公式，等号被塞在 <f> 标签**内部**，形如
    <f>=作答区!A20</f>                      （普通公式）
    <f t="array" ref="B20">=INDEX(...)</f>  （数组公式，标签带属性）
OOXML 规范要求 <f> 内容不含前导 "="。Excel/WPS 多数版本能容错，但 openpyxl
会读成 "==作答区!A20"，一旦再被脚本处理/重存就会真正写坏。

本脚本直接改 zip 内的 sheet XML，不动其他部件，因此：
  - 图表（xl/charts/*.xml）、共享字符串、样式全部原样保留
  - 不会重新触发 editor_sdk 的写入行为

默认同时去掉不必要的 t="array" 标记（INDEX/MATCH/LARGE 这类公式不需要数组输入，
留着会让阅读者编辑时必须按 Ctrl+Shift+Enter）。

用法:
    python fix_formula_tags.py 文件.xlsx
    python fix_formula_tags.py 文件.xlsx --dry-run
    python fix_formula_tags.py 文件.xlsx --keep-array     # 只去等号，保留数组标记

注意：写盘前必须先释放 editor_sdk 的文件锁（见 SKILL.md 阶段 5.1），
否则会报 WinError 32 / PermissionError 13。本脚本已内置重试与明确报错。
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import time
import zipfile

PLAIN_RE = re.compile(r"<f(\s[^>]*)?>(=)", re.S)                       # <f>=... / <f attr>=... -> 去等号
ARRAY_RE = re.compile(r'<f\s+t="array"\s+ref="[^"]*"\s*>(=)?', re.S)    # 数组公式整体归一化


def patch_bytes(data: bytes, keep_array: bool) -> tuple[bytes, int]:
    txt = data.decode("utf-8")
    n_plain = 0
    n_array = 0

    if not keep_array:
        txt, n_array = ARRAY_RE.subn("<f>", txt)
    # 去前导等号：兼容 <f> 与 <f attr...> 两种开标签
    # ⚠️ 这里的 (\s[^>]*)? 必须是「可选」：写成 (\s+[^>]*?) 就只能匹配带属性的标签，
    #    最常见的 <f>=... 会被静默漏掉 —— 那正是本脚本最容易骗过自己的地方。
    def _strip(m: re.Match) -> str:
        nonlocal n_plain
        n_plain += 1
        return "<f%s>" % (m.group(1) or "")
    txt = PLAIN_RE.sub(_strip, txt)

    return txt.encode("utf-8"), n_plain + n_array


def main() -> int:
    ap = argparse.ArgumentParser(description="归一化 xlsx 公式标签")
    ap.add_argument("xlsx")
    ap.add_argument("--dry-run", action="store_true", help="只报告，不写盘")
    ap.add_argument("--keep-array", action="store_true", help="保留 t=\"array\" 标记")
    ap.add_argument("--no-backup", action="store_true", help="不生成 .bak 备份")
    a = ap.parse_args()

    src = a.xlsx
    if not os.path.isfile(src):
        sys.exit("文件不存在: %s" % src)

    z = zipfile.ZipFile(src)
    parts = {i.filename: z.read(i.filename) for i in z.infolist()}
    infos = z.infolist()
    z.close()

    total = 0
    for name in list(parts):
        if name.startswith("xl/worksheets/") and name.endswith(".xml"):
            new, n = patch_bytes(parts[name], a.keep_array)
            if n:
                parts[name] = new
                total += n
                print("  %s: 修正 %d 处" % (name, n))

    print("待修正公式单元格合计: %d 处" % total)
    if total == 0:
        print("无需修改，文件已合规。")
        return 0
    if a.dry_run:
        print("--dry-run：未写盘。")
        return 0

    if not a.no_backup:
        bak = src + ".bak"
        shutil.copy2(src, bak)
        print("已备份 -> %s" % bak)

    tmp = src + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for i in infos:
            out.writestr(i, parts[i.filename])

    # 落盘：被 editor_sdk / Office / WPS 占用时给出可操作的提示
    for attempt in range(5):
        try:
            os.replace(tmp, src)
            print("已写回 -> %s" % src)
            break
        except PermissionError:
            print("第 %d 次落盘被占用，等待 3 秒…（请先 close_file / 关掉 WPS 里的该文件）"
                  % (attempt + 1))
            time.sleep(3)
    else:
        print("落盘失败：文件仍被占用。请先执行：\n"
              "  1) get_pool_status 找到该文件的 file_id\n"
              "  2) close_file file_id=<id> force=true\n"
              "  3) 删除同目录 ~$*.xlsx 归属文件\n"
              "修补结果暂存于: %s" % tmp)
        return 2

    # 复验
    zz = zipfile.ZipFile(src)
    left = sum(len(re.findall(r"<f(\s[^>]*)?>(=)", zz.read(n).decode("utf-8", "ignore"), re.S))
               for n in zz.namelist()
               if n.startswith("xl/worksheets/") and n.endswith(".xml"))
    arr = sum(len(re.findall(r'<f\s+t="array"', zz.read(n).decode("utf-8", "ignore")))
              for n in zz.namelist() if n.startswith("xl/worksheets/"))
    print("复验 -> 前导等号残留 %d 处，数组标记残留 %d 处，zip 完整性 %s"
          % (left, arr, "OK" if zz.testzip() is None else "异常"))
    zz.close()
    return 0 if left == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
