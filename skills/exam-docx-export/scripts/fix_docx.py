# -*- coding: utf-8 -*-
"""docx 后处理（六件事，全部幂等）：
   1) 给 SVG 图补一张高 DPI PNG 作为位图回退（Word/WPS/预览器都能显示，且保留 SVG 矢量）
   2) 图片与图注拆成两段（否则图注贴着图右侧折行）
   3) 表格行禁止跨页断开 + 封面/目录之间补分页符 + 删正文末尾空段落
   4) 表格列宽按内容自适应 + 表格拉到正文满宽 + 单元格内边距
   5) settings.xml 加 <w:updateFields w:val="true"/>，打开时自动刷新目录页码

用法：python fix_docx.py /path/to/file.docx
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

import pymupdf

DOCX = sys.argv[1] if len(sys.argv) > 1 else ""
if not DOCX:
    sys.exit("用法：python fix_docx.py /path/to/file.docx")
TMP = os.path.join(tempfile.gettempdir(), "exam_docx_fix")
SCALE = 5          # 5x → 约 360–410 DPI

# —— 表格列宽自适应参数 ——
FIT_TABLES = True  # 关掉则只做其余三项
EXP = 0.62         # 阻尼指数：0.5 更平均、1.0 完全按内容长度比例
FLOOR = 0.095      # 单列最小占比（防止短列被挤成一条缝）
CELLMAR = 108      # 单元格左右内边距 twips（5.4pt）
CHARW = 240        # 全角字宽 twips ≈ 字号 12pt（用于"表头必须单行放得下"的下限）
COVER_TBL_MARK = 'w:tblCaption w:val="covercards"'  # 封面数据卡标记，见下方 FIT_TABLES 分支


_PAT_P = re.compile(r"<w:p\b[^>]*/>|<w:p\b(?:(?!</w:p>).)*?</w:p>", re.S)


def _p_is_empty(g):
    if re.search(r"<w:t(?:\s[^>]*)?>[^<]", g):
        return False
    for k in ("<w:drawing", "<w:pict", "<w:br", "<w:tab", "<w:fldChar", "<w:instrText"):
        if k in g:
            return False
    return True


def drop_trailing_paragraphs(doc):
    """删除 </w:body> 前、sectPr 前的连续空段落（编辑器常留 1–2 个 → 多出一页空白页）"""
    i = doc.rfind("</w:body>")
    if i == -1:
        return doc, 0
    body, rest = doc[:i], doc[i:]
    si = body.rfind("<w:sectPr")
    if si == -1:
        return doc, 0
    head, sect = body[:si], body[si:]
    spans = [(m.start(), m.end(), m.group(0)) for m in _PAT_P.finditer(head)]
    keep = len(spans)
    while keep > 0 and _p_is_empty(spans[keep - 1][2]):
        keep -= 1
    removed = len(spans) - keep
    if not removed:
        return doc, 0
    cut = spans[keep][0] if keep < len(spans) else len(head)
    return head[:cut].rstrip() + sect + rest, removed


def ensure_cover_break(doc, toc_marker="目\u3000录"):
    """确保封面与目录之间有一个分页符。

    editor_sdk 的"封面 → doc_insert_page_break → doc_insert_toc"实际落成
    "封面 → 目录标题 → TOC 域 → 分页符"，即**分页符跑到目录后面去了**，
    封面和目录挤在同一页；等用户在 Word 里 F9 展开目录域（几十行）版面会被顶乱。
    """
    if doc.count('<w:br w:type="page"/>') >= 2:
        return doc, 0                      # 幂等
    i = doc.find(toc_marker)
    if i < 0:
        return doc, 0
    cur_start = doc.rfind("<w:p ", 0, i)
    if cur_start < 0:
        return doc, 0
    prev_end = doc.rfind("</w:p>", 0, cur_start) + len("</w:p>")
    if prev_end <= 0:
        return doc, 0
    if '<w:br w:type="page"/>' in doc[max(0, prev_end - 80):prev_end]:
        return doc, 0
    brk = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    return doc[:prev_end] + brk + doc[prev_end:], 1


def _vlen(s):
    """视觉宽度：CJK/全角记 1，ASCII 记 0.5"""
    return sum(1.0 if ord(ch) > 0x2E80 else 0.5 for ch in s)


def _txt(x):
    return "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", x, re.S))


def _text_width(doc):
    """从 sectPr 推正文满宽（twips）"""
    m = re.search(r'<w:pgSz w:w="(\d+)"', doc)
    pw = int(m.group(1)) if m else 11906
    mar = re.search(r"<w:pgMar\b[^>]*>", doc)
    l = r = 1134
    if mar:
        ml = re.search(r'w:left="(\d+)"', mar.group(0))
        mr = re.search(r'w:right="(\d+)"', mar.group(0))
        if ml:
            l = int(ml.group(1))
        if mr:
            r = int(mr.group(1))
    return pw - l - r


def _widths(t, textw):
    grid = [int(g) for g in re.findall(r'<w:gridCol w:w="(\d+)"', t)]
    nc = len(grid)
    if nc == 0:
        return None
    rows = re.findall(r"<w:tr\b.*?</w:tr>", t, re.S)
    cols = [[] for _ in range(nc)]
    heads = [""] * nc
    for ri, r in enumerate(rows):
        for j, c in enumerate(re.findall(r"<w:tc(?:\s[^>]*)?>.*?</w:tc>", r, re.S)[:nc]):
            s = _txt(c)
            cols[j].append(_vlen(s))
            if ri == 0:
                heads[j] = s

    # 1) 需求 = max(该列最长单元格, 表头×1.25)
    d = []
    for j, c in enumerate(cols):
        mx = max(c) if c else 1.0
        hd = _vlen(heads[j]) if heads[j] else 1.0
        d.append(max(mx, hd * 1.25, 1.0))
    p = [x ** EXP for x in d]
    sp = sum(p)
    w = [textw * x / sp for x in p]

    # 2) 抬底：单列不小于 FLOOR 占比，缺口由其余列按比例补
    floorw = FLOOR * textw
    for _ in range(8):
        low = [j for j in range(nc) if w[j] < floorw]
        if not low:
            break
        rest = [j for j in range(nc) if j not in low]
        rs = sum(w[j] for j in rest)
        if rs <= 0:
            break
        for j in low:
            w[j] = floorw
        for j in rest:
            w[j] = w[j] / rs * (textw - len(low) * floorw)

    # 3) 表头必须单行放得下（全角字宽 ≈ 字号 = CHARW twips）
    mins = []
    for j in range(nc):
        hv = _vlen(heads[j]) if heads[j] else 0.0
        mins.append(min(hv * CHARW + 2 * CELLMAR, textw * 0.6))
    sm = sum(mins)
    if sm > textw:                      # 下限总和超宽 → 等比压缩
        mins = [m * textw / sm for m in mins]
    for _ in range(8):
        need = sum(max(0.0, mins[j] - w[j]) for j in range(nc))
        if need < 1.0:
            break
        donors = [j for j in range(nc) if w[j] > mins[j] + 1]
        slack = sum(w[j] - mins[j] for j in donors)
        if slack < 1.0:
            break
        take = min(need, slack)
        for j in donors:
            w[j] -= (w[j] - mins[j]) / slack * take
        for j in range(nc):
            w[j] = max(w[j], mins[j])

    out = [int(round(x)) for x in w]
    out[-1] += textw - sum(out)          # 补齐舍入误差
    return out


def fit_table(t, textw):
    w = _widths(t, textw)
    if not w:
        return t
    gi = [0]

    def rg(m):
        v = w[gi[0]] if gi[0] < len(w) else w[-1]
        gi[0] += 1
        return '<w:gridCol w:w="%d"/>' % v

    t = re.sub(r'<w:gridCol w:w="\d+"/>', rg, t)

    def rrow(rm):
        row = rm.group(0)
        ci = [0]

        def rtc(cm):
            cell = cm.group(0)
            j = ci[0]
            ci[0] += 1
            v = w[j] if j < len(w) else w[-1]
            if "<w:tcW " in cell:
                return re.sub(r'<w:tcW w:w="\d+" w:type="\w+"/>',
                              '<w:tcW w:w="%d" w:type="dxa"/>' % v, cell, count=1)
            return cell.replace("<w:tcPr>",
                                '<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>' % v, 1)

        return re.sub(r"<w:tc(?:\s[^>]*)?>.*?</w:tc>", rtc, row, flags=re.S)

    t = re.sub(r"<w:tr\b.*?</w:tr>", rrow, t, flags=re.S)
    t = re.sub(r'<w:tblW w:w="\d+" w:type="\w+"/>',
               '<w:tblW w:w="%d" w:type="dxa"/>' % textw, t, count=1)
    if "<w:tblLayout" in t:
        t = re.sub(r'<w:tblLayout w:type="\w+"/>', '<w:tblLayout w:type="fixed"/>', t, count=1)
    else:
        t = t.replace("<w:tblPr>", '<w:tblPr><w:tblLayout w:type="fixed"/>', 1)
    mar = ('<w:tblCellMar><w:top w:w="60" w:type="dxa"/>'
           '<w:left w:w="%d" w:type="dxa"/>'
           '<w:bottom w:w="60" w:type="dxa"/>'
           '<w:right w:w="%d" w:type="dxa"/></w:tblCellMar>' % (CELLMAR, CELLMAR))
    if "<w:tblCellMar>" in t:
        t = re.sub(r"<w:tblCellMar>.*?</w:tblCellMar>", mar, t, count=1)
    else:
        t = t.replace("<w:tblPr>", "<w:tblPr>" + mar, 1)
    return t


def split_captions(doc, sid2name=None):
    """图与紧随其后的文字（图注 / 紧跟的小节标题）挤在同一段 → 拆成两段。

    不能用 `图[^<]{0,12}` 这种限定首字的正则：图注未必以「图」开头
    （本次是「数据来源：…」「上：…」「下：…」），会**静默命中 0 处**。
    改为：扫描每个 `</w:drawing></w:r>`，只要**同一段内**（下一个 `</w:p>` 之前）
    还有非空 `<w:t>`，就在图后断段，并把原段落的 `<w:pPr>` 复制给新段
    （否则紧跟的 Heading 会丢样式）。

    特例：若新段是 **Heading**（如 `11.1 考试形式与结构` 紧跟柱状图），
    要把图段的 `keepNext` 摘掉——否则「图 keepNext → H2 keepNext → 大表格」
    三级串联，图会被整块顶到下一页，上一页留下大片空白。
    """
    sid2name = sid2name or {}
    out, last, n, unchain = [], 0, 0, 0
    for m in re.finditer(r"</w:drawing></w:r>", doc):
        seg = doc[m.end():m.end() + 800]
        p_end = seg.find("</w:p>")
        head = seg[:p_end] if p_end >= 0 else seg
        t = re.search(r'<w:t(?:\s[^>]*)?>([^<]*)', head)
        if not (t and t.group(1).strip()):
            continue
        ps = None
        for pm in re.finditer(r"<w:p\b[^>]*>", doc[:m.end()]):
            ps = pm
        ppr = ""
        if ps is not None:
            pm = re.match(r"\s*<w:pPr>.*?</w:pPr>", doc[ps.end():], re.S)
            if pm:
                ppr = pm.group(0).strip()

        # 判断新段是不是 Heading → 是则摘掉图段的 keepNext
        is_head = False
        sm = re.search(r'<w:pStyle w:val="([^"]+)"/>', ppr)
        if sm:
            nm = sid2name.get(sm.group(1), "").lower()
            is_head = nm.startswith("heading")

        tail = doc[last:m.end()]
        if is_head:
            # 图段整体换成干净 pPr：去掉「heading 2」样式及其自带的 keepNext/keepLines。
            # 只摘直设 <w:keepNext> 不够 —— 样式定义里还有一份 keepNext，
            # 「图 → 标题 → 大表格」仍会三级串联，图照样被顶到下一页。
            seg = doc[ps.start():m.end()]
            m0 = re.match(r'(<w:p\b[^>]*>)\s*<w:pPr>.*?</w:pPr>', seg, re.S)
            if m0:
                seg = (m0.group(1)
                       + '<w:pPr><w:spacing w:before="0" w:after="0" '
                         'w:line="240" w:lineRule="auto"/></w:pPr>'
                       + seg[m0.end():])
                new_tail = doc[last:ps.start()] + seg
                unchain += 1
            else:
                new_tail = tail
            out.append(new_tail)
        else:
            out.append(tail)
        out.append("</w:p><w:p>" + ppr)
        last = m.end()
        n += 1
    out.append(doc[last:])
    return "".join(out), n, unchain


def normalize_image_spacing(doc):
    """图段行距改单倍。

    正文默认行距是 `line=330 lineRule=auto`（1.65 倍）。部分渲染器（腾讯文档转换、
    某些 WPS 版本）会把**内联图所在行**的行高也乘上这个倍数 —— 259pt 的图变成 427pt，
    于是「上一页明明还有 340pt 空白，图却整块被顶到下一页」。
    图段行距改 `line=240 lineRule=auto`（单倍）即可解除。
    """
    n = [0]

    def fix_p(m):
        p = m.group(0)
        if "<w:drawing>" not in p:
            return p
        head = p[:p.find(">") + 1]
        rest = p[len(head):]
        sp = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        if "<w:pPr>" not in rest:
            return head + "<w:pPr>" + sp + "</w:pPr>" + rest
        if "<w:spacing " in rest:
            new = re.sub(r"<w:spacing\b[^>]*/>", sp, rest, count=1)
        else:
            new = rest.replace("<w:pPr>", "<w:pPr>" + sp, 1)
        if new != rest:
            n[0] += 1
        return head + new

    doc = _PAT_P.sub(fix_p, doc)
    return doc, n[0]


def main():
    if os.path.exists(TMP):
        shutil.rmtree(TMP)
    os.makedirs(TMP)
    with zipfile.ZipFile(DOCX) as z:
        names = z.namelist()
        z.extractall(TMP)

    doc_path = os.path.join(TMP, "word/document.xml")
    rels_path = os.path.join(TMP, "word/_rels/document.xml.rels")
    ct_path = os.path.join(TMP, "[Content_Types].xml")
    set_path = os.path.join(TMP, "word/settings.xml")
    media = os.path.join(TMP, "word/media")

    doc = open(doc_path, encoding="utf-8").read()
    rels = open(rels_path, encoding="utf-8").read()

    rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rel_map.update(dict(re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)))
    # 正确解析：Id 在前 Target 在后
    rel_map = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt:
            rel_map[rid.group(1)] = tgt.group(1)

    used = set(int(x) for x in re.findall(r'Id="rId(\d+)"', rels))
    counter = max(used) if used else 0
    added = []

    def repl(m):
        nonlocal counter
        whole = m.group(0)
        ahead = doc[m.end():m.end() + 600]
        r = re.search(r'<asvg:svgBlip r:embed="([^"]+)"', ahead)
        if not r or 'r:embed' in whole:
            return whole
        svg_rid = r.group(1)
        tgt = rel_map.get(svg_rid, "")
        svg_path = os.path.join(TMP, "word", tgt)
        if not os.path.isfile(svg_path):
            print("  ! 找不到 SVG:", tgt)
            return whole
        counter += 1
        new_rid = "rId%d" % counter
        png_name = "image%d.png" % counter
        d = pymupdf.open(svg_path)
        pix = d[0].get_pixmap(matrix=pymupdf.Matrix(SCALE, SCALE), alpha=False)
        pix.save(os.path.join(media, png_name))
        added.append((new_rid, "media/" + png_name, pix.width, pix.height, tgt))
        return whole[:-1] + ' r:embed="%s">' % new_rid

    doc = re.sub(r'<a:blip>', repl, doc)

    # 图片与紧随其后的图注原本挤在同一段（图注会贴着图右侧折行），拆成两段
    sty = open(os.path.join(TMP, "word/styles.xml"), encoding="utf-8").read()
    sid2name = {}
    for sm in re.finditer(r"<w:style [^>]*>", sty):
        seg = sty[sm.end():sm.end() + 200]
        nm = re.search(r'<w:name w:val="([^"]*)"', seg)
        sid = re.search(r'w:styleId="([^"]+)"', sm.group(0))
        if nm and sid:
            sid2name[sid.group(1)] = nm.group(1)
    doc, nsplit, nunchain = split_captions(doc, sid2name)
    print("  图注段落拆分: %d（其中紧跟标题、已摘 keepNext: %d）" % (nsplit, nunchain))

    # 图段行距改单倍（否则部分渲染器把行距倍数乘到图高上，图被顶到下一页）
    doc, nimg = normalize_image_spacing(doc)
    print("  图段行距改单倍:", nimg)

    # 表格行禁止跨页断开（否则断页处会留下一行空白单元格）
    def _tr(m):
        tr = m.group(0)
        if "<w:cantSplit/>" in tr:
            return tr
        if "<w:trPr>" in tr:
            return tr.replace("<w:trPr>", "<w:trPr><w:cantSplit/>", 1)
        i = tr.find(">")
        return tr[:i + 1] + "<w:trPr><w:cantSplit/></w:trPr>" + tr[i + 1:]

    doc, ntr = re.subn(r"<w:tr\b[^>]*>.*?</w:tr>", _tr, doc, flags=re.S)
    print("  表格行 cantSplit:", ntr)

    # 封面与目录之间补分页符（editor_sdk 会把分页符插到目录后面）
    doc, nbrk = ensure_cover_break(doc)
    print("  封面/目录分页符:", nbrk)

    # 删掉正文末尾的空段落（否则会多出一整页空白页）
    doc, ndrop = drop_trailing_paragraphs(doc)
    print("  末尾空段落删除:", ndrop)

    # 表格列宽自适应（markdown 落进来的表默认全列等宽，长文本列会被压成一列一字）
    # 例外：带 <w:tblCaption w:val="covercards"/> 标记的表是封面数据卡（patch_cover.py 造的，
    # 列宽是手工算好的），必须原样放过。否则在「已有 docx 再体检」流程里重跑本脚本，
    # 会把 4 张卡压成等宽，封面版式静默走样。
    if FIT_TABLES:
        textw = _text_width(doc)
        cnt = [0]
        skip = [0]

        def _ft(m):
            t = m.group(0)
            if COVER_TBL_MARK in t:
                skip[0] += 1
                return t
            cnt[0] += 1
            return fit_table(t, textw)

        doc = re.sub(r"<w:tbl>.*?</w:tbl>", _ft, doc, flags=re.S)
        print("  表格列宽自适应: %d 张（正文满宽 %d twips）｜跳过封面卡 %d 张"
              % (cnt[0], textw, skip[0]))

    for new_rid, target, w, h, src in added:
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org'
                            '/officeDocument/2006/relationships/image" Target="%s"/>'
                            "</Relationships>" % (new_rid, target))
        print("  + %s → %s (%d×%d) 源 %s" % (new_rid, target, w, h, src))
    open(rels_path, "w", encoding="utf-8").write(rels)

    ct = open(ct_path, encoding="utf-8").read()
    if 'Extension="png"' not in ct:
        ct = ct.replace("<Types ", "<Types ", 1)
        ct = re.sub(r'(<Types[^>]*>)', r'\1<Default Extension="png" ContentType="image/png"/>',
                    ct, count=1)
        open(ct_path, "w", encoding="utf-8").write(ct)
        print("  + Content_Types 增加 png")

    st = open(set_path, encoding="utf-8").read()
    if "updateFields" not in st:
        st = re.sub(r'(<w:settings[^>]*>)', r'\1<w:updateFields w:val="true"/>', st, count=1)
        open(set_path, "w", encoding="utf-8").write(st)
        print("  + settings 增加 updateFields")

    open(doc_path, "w", encoding="utf-8").write(doc)

    # 重新打包
    out = DOCX
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(os.path.join(TMP, n), n)
        for _, target, _, _, _ in added:
            z.write(os.path.join(TMP, "word", target), "word/" + target)
    print("  已重写:", out, "新增位图", len(added), "张")


if __name__ == "__main__":
    main()
