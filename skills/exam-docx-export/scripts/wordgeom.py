# -*- coding: utf-8 -*-
"""Word 版式几何估算器：无渲染器时用于定位「大片空白」与估算页数。

模型：
  A4 29.7cm 高，上下边距 2.2cm → 正文区 717.17pt
  左右边距 2.4cm → 正文宽 459.21pt
  段落高 = 行数 × 字号 × 1.30 × 行距倍数 + 段前 + 段后
  中文按 1em、西文按 0.5em 估宽；表格按列宽独立换行取行内最大值
"""
import sys
import zipfile
from docx import Document
from docx.shared import Pt, Length
from docx.oxml.ns import qn

CONTENT_H = 841.89 - 2 * 62.36          # 717.17
CONTENT_W = 595.28 - 2 * 68.03          # 459.21
LINE_K = 1.30                            # 单倍行高系数（宋体/雅黑经验值）
DEFAULT_SIZE = 10.5


def char_w(ch, size, bold=False):
    o = ord(ch)
    if o > 0x2E80 or ch in "　「」【】（）《》，。：；！？、":
        return size * (1.0 if not bold else 1.02)
    return size * 0.50


def text_w(s, size, bold=False):
    return sum(char_w(c, size, bold) for c in s)


def run_size(r, default=DEFAULT_SIZE):
    if r.font.size is not None:
        return r.font.size.pt
    return default


def doc_defaults(doc):
    """从 docDefaults 读默认行距/段后/字号——段落未显式设置时按此计算"""
    line, after, size = 1.0, 0.0, 10.5
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is None:
        return line, after, size
    pPr = dd.find(qn("w:pPrDefault"))
    if pPr is not None:
        pp = pPr.find(qn("w:pPr"))
        if pp is not None:
            sp = pp.find(qn("w:spacing"))
            if sp is not None:
                lv = sp.get(qn("w:line"))
                if lv and sp.get(qn("w:lineRule")) == "auto":
                    line = int(lv) / 240.0
                av = sp.get(qn("w:after"))
                if av:
                    after = int(av) / 20.0
    rPr = dd.find(qn("w:rPrDefault"))
    if rPr is not None:
        rp = rPr.find(qn("w:rPr"))
        if rp is not None:
            sz = rp.find(qn("w:sz"))
            if sz is not None and sz.get(qn("w:val")):
                size = int(sz.get(qn("w:val"))) / 2.0
    return line, after, size


DEFAULTS = [1.0, 0.0, 10.5]      # 由 docDefaults 填充：行距 / 段后 / 字号


def set_defaults(doc):
    """读 docDefaults（行距/段后/字号）+ 页面设置（正文区宽高）"""
    l, a, sz = doc_defaults(doc)
    DEFAULTS[:] = [l, a, sz]
    global CONTENT_W, CONTENT_H
    try:
        s = doc.sections[0]
        if s.page_width and s.left_margin is not None and s.right_margin is not None:
            CONTENT_W = (s.page_width - s.left_margin - s.right_margin).pt
        if s.page_height and s.top_margin is not None and s.bottom_margin is not None:
            CONTENT_H = (s.page_height - s.top_margin - s.bottom_margin).pt
    except Exception:
        pass                                  # 读不到就沿用文件头部的默认值
    return DEFAULTS


def para_height(p, avail_w=CONTENT_W, default_size=None, default_line=None,
                default_after=None):
    """返回 (高度pt, 行数)"""
    if default_size is None:
        default_size = DEFAULTS[2]
    if default_line is None:
        default_line = DEFAULTS[0]
    if default_after is None:
        default_after = DEFAULTS[1]
    pf = p.paragraph_format
    sb = pf.space_before.pt if pf.space_before is not None else 0
    sa = pf.space_after.pt if pf.space_after is not None else default_after
    ls = pf.line_spacing
    li = pf.left_indent.pt if pf.left_indent is not None else 0
    ri = pf.right_indent.pt if pf.right_indent is not None else 0
    avail = max(40.0, avail_w - li - ri)

    # 字号：取段内最大值作为行高基准
    sizes = [run_size(r, default_size) for r in p.runs if r.text]
    size = max(sizes) if sizes else default_size

    # 行距（注意 Length 是 int 子类，必须先判 Length）
    if ls is None:
        mult = default_line
        fixed = None
    elif isinstance(ls, Length):
        mult = None
        fixed = ls.pt
    elif isinstance(ls, (int, float)):
        mult = float(ls)
        fixed = None
    else:
        mult = default_line
        fixed = None

    # 宽度 → 行数（按 run 分别累计，粗略按空格切行）
    total = 0.0
    for r in p.runs:
        if not r.text:
            continue
        total += text_w(r.text, run_size(r, default_size), bool(r.font.bold))
    if total <= 0:
        lines = 1
    else:
        lines = 1
        acc = 0.0
        for r in p.runs:
            if not r.text:
                continue
            rs = run_size(r, default_size)
            for seg in r.text.split("\n"):
                if seg is not None and seg != r.text.split("\n")[0]:
                    lines += 1
                    acc = 0.0
                for ch in seg:
                    w = char_w(ch, rs, bool(r.font.bold))
                    if acc + w > avail:
                        lines += 1
                        acc = 0.0
                    acc += w

    if mult is None:
        lh = fixed
    else:
        lh = size * LINE_K * mult
    return lines * lh + sb + sa, lines


def cell_height(cell, avail_w):
    """单元格高度：末段的段后距不计入（Word 在单元格边界处丢弃）"""
    ps = cell.paragraphs
    h = 0.0
    for i, p in enumerate(ps):
        ph, _ = para_height(p, avail_w, default_after=0.0 if i == len(ps) - 1 else None)
        h += ph
    return h


def table_row_heights(t, avail_w=CONTENT_W, cm=None):
    """逐行高度（行可跨页断开，只有 cantSplit 阻止单行内部断开）"""
    ncol = len(t.columns)
    if cm:
        widths = [c / 2.54 * 72 for c in cm]
    else:
        widths = [avail_w / ncol] * ncol
    hs = []
    for row in t.rows:
        rh = 0.0
        for i, cell in enumerate(row.cells):
            if i >= len(widths):
                break
            inner = max(30.0, widths[i] - 8.5)
            rh = max(rh, cell_height(cell, inner))
        hs.append(rh + 5.0)
    return hs


def table_height(t, avail_w=CONTENT_W, cm=None):
    return sum(table_row_heights(t, avail_w, cm))


def block_units(doc, cm=None):
    """把正文拆成可换页的最小单元：段落 1 个；表格按行拆"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield ("p", Paragraph(child, doc._body))
        elif child.tag == qn("w:tbl"):
            yield ("tbl", Table(child, doc._body))


def simulate(doc, cm=None):
    """返回每页占用高度列表"""
    set_defaults(doc)
    pages = [0.0]
    for kind, obj in block_units(doc, cm):
        if kind == "p":
            h = para_height(obj, CONTENT_W)[0]
            pbb = bool(obj.paragraph_format.page_break_before)
            brk = any(b.get(qn("w:type")) == "page" for b in obj._p.iter(qn("w:br")))
            if pbb and pages[-1] > 0:
                pages.append(0.0)
            if pages[-1] + h > CONTENT_H and pages[-1] > 0:
                pages.append(0.0)
            pages[-1] += h
            if brk:
                pages.append(0.0)
        else:
            for rh in table_row_heights(obj, CONTENT_W, cm):
                if pages[-1] + rh > CONTENT_H and pages[-1] > 0:
                    pages.append(0.0)
                pages[-1] += rh
    return [p for p in pages if p > 0]


def walk_blocks(body):
    """产出 ('p', para) / ('tbl', table) 序列"""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield ("p", Paragraph(child, body))
        elif child.tag == qn("w:tbl"):
            yield ("tbl", Table(child, body))


def report(path, cm=None, verbose=True):
    doc = Document(path)
    set_defaults(doc)
    real = simulate(doc, cm)

    print(f"\n===== {path.split('/')[-1]} =====")
    print(f"正文区 {CONTENT_W:.1f} × {CONTENT_H:.1f} pt")
    print(f"估算页数：{len(real)}")
    for i, used in enumerate(real, 1):
        fill = used / CONTENT_H * 100
        bar = "█" * int(fill / 3)
        flag = "  ← 空白偏多" if fill < 78 else ("  ← 溢出" if fill > 100 else "")
        print(f"  P{i:02d}  {used:7.1f}pt  {fill:5.1f}%  {bar}{flag}")
    return real


if __name__ == "__main__":
    p = sys.argv[1]
    cm = [float(x) for x in sys.argv[2].split(",")] if len(sys.argv) > 2 else None
    report(p, cm)
