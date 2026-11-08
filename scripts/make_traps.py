# -*- coding: utf-8 -*-
"""
边界实验：给一份 xlsx 注入三类"本该被抓到"的陷阱，用于**实测**四层验证阶梯的能力边界。

它回答的不是"工具能不能用"，而是"**工具在哪一层会漏报**"。
实测结论（见 references/boundary-experiment.md）：L1/L3/L4 大面积漏报，只有 L2 全部命中
—— 因为 L1（openpyxl data_only=True）与 L3（静态 XML）读的都是**缓存**，共享同一盲区。

三类陷阱
--------
  A 缓存篡改   公式正确，把 <v> 缓存改成 12345            → 期望只有 L2 能抓
  B 公式改错   首个公式里的区域引用起始行 +1（缓存未同步） → 期望只有 L2 能抓
  C 越界引用   首个公式里的区域引用改到第 7000 行外        → 期望只有 L2 能抓

用法
----
    # 用自带 demo 样本（无需外部文件，仅需 openpyxl）
    python make_traps.py --out ./边界实验

    # 用你自己的成品（需含公式且带 <v> 缓存）
    python make_traps.py 成品.xlsx --out ./边界实验

生成后逐个跑四层，观察哪层报 FAIL：

    for t in A-缓存陷阱 B-公式错误 C-越界引用; do
        python run_all.py "$t.xlsx" --engine excel,wps --out "./out_$t"
    done

两个实现要点（都踩过坑，勿改回去）
--------------------------------
1. **只能在 <f>…</f> 内部找区域引用**。直接对整份 XML 做正则，会先匹配到
   `<dimension ref="A1:C3"/>`，把维度声明改了而公式纹丝不动 —— 陷阱静默失效。
2. **openpyxl 直出的公式单元格带的是空缓存 `<v></v>`**（不是没有 `<v>`）。
   所以 demo 样本必须显式写入真实的缓存值，否则"缓存 vs 重算"这一路根本无从比对。
"""
import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

SHEET_RE = re.compile(r"xl/worksheets/sheet\d+\.xml$")
# 区域引用：$B$7:$B$16 / B7:B16
RANGE_RE = re.compile(r"(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)")
# 单个引用：$B$7
REF_RE = re.compile(r"(\$?[A-Z]{1,3})(\$?)(\d+)")
# 公式标签（捕获内部文本）
FORMULA_RE = re.compile(r"(<f\b[^>]*>)([^<]*)(</f>)")
# 单元格 + 公式（可选缓存），用于写入/替换缓存
CELL_RE = re.compile(r"(<c\b[^>]*>)(<f\b[^>]*>[^<]*</f>)(?:<v>[^<]*</v>)?</c>")
# 已有缓存的公式单元格
CACHED_RE = re.compile(r"(<c\b[^>]*>)(<f\b[^>]*>.*?</f>)<v>[^<]*</v>", re.S)

TRAPS = {
    "A": ("A-缓存陷阱.xlsx", "缓存篡改（公式对、缓存错）"),
    "B": ("B-公式错误.xlsx", "公式引用改错（起始行 +1）"),
    "C": ("C-越界引用.xlsx", "越界引用（重算值必然不同）"),
}


def rewrite(src, dst, fn):
    """逐部件读出 xlsx，对每个 sheet xml 应用 fn，再写回（src 与 dst 可相同）。"""
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    for name in list(items):
        if SHEET_RE.match(name):
            items[name] = fn(name, items[name].decode("utf-8")).encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in items.items():
            z.writestr(n, b)


def set_caches(xml, values):
    """按单元格地址写入缓存值。openpyxl 直出的公式单元格缓存是空的，必须补。"""
    def rep(m):
        head, f = m.group(1), m.group(2)
        ref = re.search(r'r="([^"]+)"', head).group(1)
        v = values.get(ref)
        return m.group(0) if v is None else head + f + "<v>%s</v>" % v + "</c>"
    return CELL_RE.sub(rep, xml)


def build_demo(path: Path):
    """生成最小样本：A1:A10 = 1..10，三个公式并写入**正确的**缓存值。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "数据"
    for i in range(1, 11):
        ws.cell(i, 1, i)
    ws["C1"] = "=SUM(A1:A10)"            # 55
    ws["C2"] = "=LARGE(A1:A10,1)"        # 10
    ws["C3"] = '=SUMIF(A1:A10,">5")'     # 40
    wb.save(path)
    rewrite(path, path, lambda n, x: set_caches(x, {"C1": 55, "C2": 10, "C3": 40}))


def _patch_first_formula(xml: str, fn) -> str:
    """只对**第一个含区域引用的公式**应用 fn —— 绝不在整份 XML 上做替换。"""
    hit = []

    def rep(m):
        if hit:
            return m.group(0)
        new = fn(m.group(2))
        if new != m.group(2):
            hit.append(1)
            return m.group(1) + new + m.group(3)
        return m.group(0)

    return FORMULA_RE.sub(rep, xml)


def _shift_start_row(formula: str) -> str:
    """把公式里首个区域引用的起始行 +1（制造"少一行"的公式错误）。"""
    m = RANGE_RE.search(formula)
    if not m:
        return formula
    a = REF_RE.match(m.group(1))
    new_a = "%s%s%d" % (a.group(1), a.group(2), int(a.group(3)) + 1)
    return formula[:m.start()] + new_a + ":" + m.group(2) + formula[m.end():]


def _to_out_of_range(formula: str) -> str:
    """把公式里首个区域引用改到第 7000 行外（重算结果必然不同）。"""
    m = RANGE_RE.search(formula)
    if not m:
        return formula
    a = REF_RE.match(m.group(1))
    prefix = a.group(1) + a.group(2)
    return formula[:m.start()] + prefix + "7000:" + prefix + "7016" + formula[m.end():]


def trap_a(name, xml):
    """缓存篡改：公式不动，把第一个有缓存的公式单元格缓存改成 12345。"""
    hit = []

    def rep(m):
        if hit:
            return m.group(0)
        ref = re.search(r'r="([^"]+)"', m.group(1)).group(1)
        hit.append(ref)
        print("  [A] %s!%s  缓存 → 12345（公式保持不变）" % (name, ref))
        return m.group(1) + m.group(2) + "<v>12345</v>"

    return CACHED_RE.sub(rep, xml)


def trap_b(name, xml):
    before = xml
    out = _patch_first_formula(xml, _shift_start_row)
    if out != before:
        m = re.search(r"<f\b[^>]*>([^<]*)</f>", out)
        print("  [B] %s  公式改为 %s（缓存未同步）" % (name, m.group(1) if m else "?"))
    return out


def trap_c(name, xml):
    before = xml
    out = _patch_first_formula(xml, _to_out_of_range)
    if out != before:
        m = re.search(r"<f\b[^>]*>([^<]*)</f>", out)
        print("  [C] %s  公式改为 %s" % (name, m.group(1) if m else "?"))
    return out


def main():
    ap = argparse.ArgumentParser(description="给 xlsx 注入三类验证陷阱")
    ap.add_argument("src", nargs="?", help="源 xlsx；省略则用内置 demo 样本")
    ap.add_argument("--out", default=".", help="输出目录（默认当前目录）")
    a = ap.parse_args()

    outdir = Path(a.out)
    outdir.mkdir(parents=True, exist_ok=True)

    if a.src:
        src = Path(a.src)
        if not src.is_file():
            sys.exit("源文件不存在: %s" % src)
    else:
        src = outdir / "_demo样本.xlsx"
        print("== 未给 --src，生成内置 demo 样本 ==")
        build_demo(src)

    made = []
    for key, (fname, desc) in TRAPS.items():
        dst = outdir / fname
        print("== 陷阱 %s：%s ==" % (key, desc))
        shutil.copy(src, dst)
        fn = {"A": trap_a, "B": trap_b, "C": trap_c}[key]
        rewrite(src, dst, fn)
        made.append(dst)
        print("   已生成 %s" % dst.name)

    print()
    print("提示：逐个跑四层，观察哪一层报 FAIL —— 只有 L2 应当全部命中。")
    print("注意：越界引用（C）在 SUM 这类函数上表现为**数值不符**；")
    print("      若原公式是 LARGE/INDEX，则会直接报 #NUM!/#REF!，属正常。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
