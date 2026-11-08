#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成 docx 域更新自检样本。

用途：验证 `verify_docx.py` 的"重算"是真的在重算、而不是在读文件里缓存的域结果。

做法：表格里放 10 / 20 / 30，下面一行放公式域 `{ =SUM(ABOVE) }`，
      **并把这个域的缓存结果故意写成 999999**。
      正确实现（会调 Fields.Update()）应算回 60；
      如果两个引擎都报 999999，说明它们只读了缓存 —— 那验证就是无效的。

用法:
    python make_docx_field_sample.py [输出路径.docx]
    # 然后：
    python verify_docx.py 输出路径.docx --engine word,wps

依赖：python-docx（pip install python-docx）
"""
from __future__ import annotations

import os
import sys

OUT_DEFAULT = "docx_域更新测试样本.docx"
WRONG_CACHE = "999999"      # 故意写错的域缓存
EXPECTED = 60               # 10 + 20 + 30


def main() -> int:
    out = sys.argv[1] if len(sys.argv) > 1 else OUT_DEFAULT
    try:
        from docx import Document
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
    except ImportError:
        sys.exit("缺少依赖：pip install python-docx")

    doc = Document()
    doc.add_paragraph("表格公式域更新测试")

    t = doc.add_table(rows=5, cols=2)
    try:
        t.style = "Table Grid"
    except Exception:                                      # noqa: BLE001
        pass
    t.cell(0, 0).text = "项目"
    t.cell(0, 1).text = "数量"
    for i, v in enumerate((10, 20, 30), start=1):
        t.cell(i, 0).text = "明细%d" % i
        t.cell(i, 1).text = str(v)

    # 公式行放在最后一行（即三个数字的**下方**），SUM(ABOVE) 才能取到 10+20+30
    row = t.rows[4]
    row.cells[0].text = "合计"
    p = row.cells[1].paragraphs[0]

    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "=SUM(ABOVE)")
    r = OxmlElement("w:r")
    tel = OxmlElement("w:t")
    tel.text = WRONG_CACHE          # ← 故意写错的缓存结果
    r.append(tel)
    fld.append(r)
    p._p.append(fld)

    doc.add_paragraph("结束")
    doc.save(out)

    print("已生成: %s" % os.path.abspath(out))
    print("自检口径：=SUM(ABOVE) 上方为 10 + 20 + 30，正确实现应报 %d" % EXPECTED)
    print("         文件里缓存的却是 %s —— 若引擎报 %s，说明它没重算，验证无效" % (WRONG_CACHE, WRONG_CACHE))
    return 0


if __name__ == "__main__":
    sys.exit(main())
