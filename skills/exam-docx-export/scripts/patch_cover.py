# -*- coding: utf-8 -*-
r"""成品 docx 封面补版式 ＋ 目录域静态兜底（幂等，zip 层直改，不走 editor_sdk）

用法：python patch_cover.py <目标.docx> [备份.docx]

做两件事：
  1) 把封面整块换掉：主色细线（pBdr）＋ 双色大标题 ＋ 副标题 ＋ 1×4 数据卡表 ＋ 署名
     —— editor_sdk 落的封面只是 Title/Subtitle 两行纯文字，与 PDF 封面差距肉眼可见
  2) 把 TOC 域改写成多段落结构（begin 段 ＋ 静态章名段 × N ＋ end 段），
     并把 \o "1-2" 收成 \o "1-1"
     —— 域未求值时（腾讯文档预览 / Quick Look / Finder）显示干净章名目录而不是"未找到目录项。"

⚠️ 必须在 fix_docx.py / fix_titlepg.py **之后**跑，且之后不要再调 save_file。
⚠️ 封面间距（build_cover 的 sp1/sp2）给大了会让封面超出一页、**多出一整张空白页**；
   内容估算总高要 ≤ 可用高的 ~85%（A4+2cm 边距 = 14570 twips ≈ 728pt）。改完必须重渲染数页数。
⚠️ 默认字体色值/文案是"黑白红克制"版式的实例，换项目改 CARDS / build_cover() / CH 即可。
⚠️ 幂等：从 BAK 读、写回 DST。BAK 不存在时先自动备份一次。
"""
import os
import re
import shutil
import sys
import zipfile

SRC = sys.argv[1] if len(sys.argv) > 1 else ""
if not SRC:
    sys.exit("用法：python patch_cover.py <目标.docx> [备份.docx]")
DST = SRC
BAK = sys.argv[2] if len(sys.argv) > 2 else SRC.replace(".docx", ".bak.docx")

BOOK_TITLE = "2027 年国考考情手册"   # 封面区自检锚点，按当次项目改

CH = [
    "编制说明",
    "第一章　国考的制度底色",
    "第二章　历年招录全景（2019—2026）",
    "第三章　竞争态势八年演变",
    "第四章　2026 年度公告深度解读",
    "第五章　2027 年度前瞻",
    "第六章　招录层级与系统版图",
    "第七章　报考条件逐条拆解",
    "第八章　职位类别与试卷类别",
    "第九章　行测赋分标准：副省级卷",
    "第十章　行测赋分标准：地市级与行政执法类卷",
    "第十一章　行测模块分值权重",
    "第十二章　申论题型与分值分布",
    "第十三章　申论小题评分标准",
    "第十四章　申论作文评分标准",
    "第十五章　申论作答规范与卷面要求",
    "第十六章　全流程时间轴",
    "第十七章　面试考情",
    "第十八章　成绩合成与体检考察",
    "第十九章　广东考区专题",
    "第二十章　备考节奏与时间表",
    "附录A　历年招录参数总表",
    "附录B　行测赋分总表",
    "附录C　申论赋分总表与数据口径",
]

CARDS = [
    ("3.81", "万", "2026 年度计划招录"),
    ("371.8", "万", "报名过审人数"),
    ("98", " : 1", "过审人数与计划数之比"),
    ("32.96", "万", "广东报名人数 · 各省第一"),
]


def run(text, *, color="1A1A1A", sz=21, bold=False, italic=False):
    rpr = ['<w:rFonts w:hint="eastAsia"/>']
    if bold:
        rpr.append('<w:b w:val="1"/><w:bCs w:val="1"/>')
    if italic:
        rpr.append('<w:i w:val="1"/>')
    if color:
        rpr.append('<w:color w:val="%s"/>' % color)
    rpr.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz))
    return "<w:r><w:rPr>%s</w:rPr><w:t>%s</w:t></w:r>" % ("".join(rpr), text)


def para(inner, *, jc=None, after=0, before=0, line=240, pBdr=None):
    p = ["<w:pPr>"]
    if pBdr:
        p.append(pBdr)
    p.append('<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>'
             % (before, after, line))
    if jc:
        p.append('<w:jc w:val="%s"/>' % jc)
    p.append("</w:pPr>")
    return "<w:p>%s%s</w:p>" % ("".join(p), inner)


def card_cell(w, num, unit, label):
    n = run(num, color="A62B2B", sz=40, bold=True)
    if unit:
        n += run(unit, color="A62B2B", sz=22)
    p1 = para(n, jc="center")
    p2 = para(run(label, color="6E6E6E", sz=16), jc="center")
    return ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s%s</w:tc>'
            % (w, p1, p2))


def build_cards():
    widths = [2410, 2410, 2409, 2409]
    cells = "".join(card_cell(w, *c) for w, c in zip(widths, CARDS))
    borders = ('<w:tblBorders>'
               '<w:top w:val="single" w:color="1A1A1A" w:sz="8" w:space="0"/>'
               '<w:bottom w:val="single" w:color="1A1A1A" w:sz="8" w:space="0"/>'
               '<w:insideV w:val="single" w:color="D9D9D9" w:sz="6" w:space="0"/>'
               '</w:tblBorders>')
    return ('<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/>%s'
            '<w:tblLayout w:type="fixed"/><w:tblCaption w:val="covercards"/>'
            '<w:tblCellMar>'
            '<w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
            '<w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>'
            '</w:tblCellMar></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2410"/><w:gridCol w:w="2410"/>'
            '<w:gridCol w:w="2409"/><w:gridCol w:w="2409"/></w:tblGrid>'
            '<w:tr><w:trPr><w:cantSplit/></w:trPr>%s</w:tr></w:tbl>'
            % (borders, cells))


def build_cover(sp1=4600, sp2=2950):
    rule = ('<w:pBdr><w:bottom w:val="single" w:sz="18" w:space="2" '
            'w:color="A62B2B"/></w:pBdr>')
    top_rule = ('<w:pBdr><w:top w:val="single" w:sz="6" w:space="6" '
                'w:color="D9D9D9"/></w:pBdr>')
    out = []
    out.append(para(run("2027 年度 · 国家公务员考试", color="6E6E6E", sz=18),
                    after=120, pBdr=rule))
    out.append(para(run("2027 年国考考情", color="1A1A1A", sz=56, bold=True)))
    out.append(para(run("考情手册", color="A62B2B", sz=56, bold=True), after=180))
    out.append(para(run("行测 · 申论赋分标准逐项拆解　|　历年招录规模 · 竞争态势 · 报考门槛 · 全流程",
                        color="3C3C3C", sz=21), after=60))
    out.append(para(run("附 2027 年度时间节点前瞻", color="A62B2B", sz=21)))
    out.append(para("", after=sp1))
    out.append(build_cards())
    out.append(para("", after=sp2))
    out.append(para(run("广东中公教研", color="1A1A1A", sz=20), pBdr=top_rule))
    out.append(para(run("2026 年 9 月编制", color="6E6E6E", sz=18)))
    return "".join(out)


def build_toc():
    body = "".join(para(run(t, color="1A1A1A", sz=20), after=40) for t in CH)
    head = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText>TOC \\u \\o &quot;1-1&quot; \\z \\h '
            '\\tdkey wg2si6</w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>')
    tail = ('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    return head + body + tail


def para_start(doc, pos):
    """pos 处文本所在 <w:p 的段首（顺序无关，兼容任意属性排布）"""
    return doc.rfind("<w:p", 0, pos)


def main():
    if not os.path.exists(BAK):
        shutil.copy2(SRC, BAK)
        print("backup ->", BAK)
    z = zipfile.ZipFile(BAK)
    doc = z.read("word/document.xml").decode("utf-8")

    # --- 封面替换：<w:body> 之后 → 第一个分页符段之前 ---
    # ⚠️ 锚点按内容定位，不用 w14:paraId（每次生成都是随机值，必失效）
    # ⚠️ 替换起点必须是 <w:body> 之后：直接 new_cover + doc[m0:] 会把
    #    XML 声明＋<w:document> 根元素一起吞掉，document.xml 变成裸 <w:p>
    #    开头，python-docx/lxml 全部报 "Namespace prefix w not defined"
    mb = doc.find('<w:br w:type="page"/>')
    assert mb > 0, "找不到分页符"
    m0 = para_start(doc, mb)
    body_start = doc.find("<w:body>") + len("<w:body>")
    old_cover = doc[body_start:m0]
    assert BOOK_TITLE in old_cover, "封面区未找到书名: " + BOOK_TITLE
    new_cover = build_cover()
    doc = doc[:body_start] + new_cover + doc[m0:]
    print("cover: %d -> %d chars" % (len(old_cover), len(new_cover)))

    # --- 目录字段替换：按 instrText 内容定位 ---
    ti = doc.find("<w:instrText")
    assert ti > 0 and "TOC" in doc[ti:ti + 200], "找不到 TOC 域"
    t0 = para_start(doc, ti)
    t1 = doc.find("</w:p>", ti) + len("</w:p>")
    old_toc = doc[t0:t1]
    assert "TOC" in old_toc, old_toc[:200]
    doc = doc[:t0] + build_toc() + doc[t1:]
    print("toc: %d -> %d chars" % (len(old_toc), len(build_toc())))

    # --- 回写 ---
    tmp = DST + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            data = doc.encode("utf-8") if it.filename == "word/document.xml" \
                else z.read(it.filename)
            zo.writestr(it, data)
    z.close()
    shutil.move(tmp, DST)
    print("written:", DST, os.path.getsize(DST))


if __name__ == "__main__":
    main()
