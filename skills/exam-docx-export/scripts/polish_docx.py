# -*- coding: utf-8 -*-
"""成品 docx 收尾美化（幂等，zip 层直改，不走 editor_sdk）

用法：python polish_docx.py <目标.docx> [ink_hex] [red_hex]

做两件事：
  1) 把目录域占位串「未找到目录项。」换成 F9 提示语
     —— 域未刷新时（腾讯文档预览 / Quick Look / 部分 WPS）显示人话而不是报错文案
  2) 重写 styles.xml：宋体正文 + 微软雅黑标题 + H1 暗红细下边线
     —— editor_sdk 的默认样式很素（Normal 无字体字号），要还原"黑白红克制"的 PDF 调性

⚠️ 必须在 fix_docx.py / fix_titlepg.py **之后**跑，且之后不要再调 save_file。
⚠️ styleId 是 editor_sdk 随机生成的 → 本脚本按 `<w:name w:val="…"/>` 反查 id，不要硬编码。
"""
import re
import sys
import zipfile

P = sys.argv[1] if len(sys.argv) > 1 else ""
if not P:
    sys.exit("用法：python polish_docx.py <目标.docx> [ink_hex] [red_hex]")
INK = sys.argv[2] if len(sys.argv) > 2 else "1A1A1A"      # 墨（正文/标题）
RED = sys.argv[3] if len(sys.argv) > 3 else "A62B2B"      # 暗红（点缀）
HEI = "微软雅黑"      # 标题中文字体
SON = "宋体"          # 正文中文字体
LAT = "Times New Roman"

TOC_OLD = "未找到目录项。"
TOC_NEW = "打开文档后按 F9 或右键「更新域」生成目录（含页码与超链接）"


def style_id_of(styles_xml, name):
    """按 <w:name w:val="Normal"/> 反查该样式的 styleId"""
    for m in re.finditer(r'<w:style [^>]*>', styles_xml):
        seg = styles_xml[m.end():m.end() + 200]
        nm = re.search(r'<w:name w:val="([^"]*)"', seg)
        if nm and nm.group(1) == name:
            sid = re.search(r'w:styleId="([^"]+)"', m.group(0))
            if sid:
                return sid.group(1)
    return None


def build_styles(ids):
    """ids: {"Normal":..,"Title":..,"Subtitle":..,"heading 1":..,"heading 2":..,"heading 3":..}"""
    n = ids.get("Normal", "Normal")
    t = ids.get("Title", "Title")
    s = ids.get("Subtitle", "Subtitle")
    h1 = ids.get("heading 1", "Heading1")
    h2 = ids.get("heading 2", "Heading2")
    h3 = ids.get("heading 3", "Heading3")

    def fonts(east):
        return ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
                % (LAT, LAT, east, LAT))

    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>%(f_son)s<w:color w:val="%(ink)s"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="100" w:line="330" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>

<w:style w:type="paragraph" w:default="1" w:styleId="%(n)s"><w:name w:val="Normal"/>
<w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/>
<w:spacing w:before="0" w:after="100" w:line="330" w:lineRule="auto"/></w:pPr>
<w:rPr>%(f_son)s<w:color w:val="%(ink)s"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="%(t)s"><w:name w:val="Title"/>
<w:basedOn w:val="%(n)s"/><w:next w:val="%(n)s"/><w:uiPriority w:val="9"/><w:qFormat w:val="1"/>
<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/>
<w:spacing w:before="0" w:after="160" w:line="288" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
<w:rPr>%(f_hei)s<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="%(ink)s"/>
<w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="%(s)s"><w:name w:val="Subtitle"/>
<w:basedOn w:val="%(n)s"/><w:next w:val="%(n)s"/><w:uiPriority w:val="9"/><w:qFormat w:val="1"/>
<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/>
<w:spacing w:before="0" w:after="160" w:line="288" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
<w:rPr>%(f_hei)s<w:b w:val="0"/><w:bCs w:val="0"/><w:color w:val="%(red)s"/>
<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="%(h1)s"><w:name w:val="heading 1"/>
<w:basedOn w:val="%(n)s"/><w:next w:val="%(n)s"/><w:uiPriority w:val="9"/><w:qFormat w:val="1"/>
<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/>
<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="%(red)s"/></w:pBdr>
<w:spacing w:before="360" w:after="200" w:line="288" w:lineRule="auto"/>
<w:outlineLvl w:val="0"/></w:pPr>
<w:rPr>%(f_hei)s<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="%(ink)s"/>
<w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="%(h2)s"><w:name w:val="heading 2"/>
<w:basedOn w:val="%(n)s"/><w:next w:val="%(n)s"/><w:uiPriority w:val="9"/><w:qFormat w:val="1"/>
<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/>
<w:spacing w:before="240" w:after="120" w:line="288" w:lineRule="auto"/>
<w:outlineLvl w:val="1"/></w:pPr>
<w:rPr>%(f_hei)s<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="%(ink)s"/>
<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>

<w:style w:type="paragraph" w:styleId="%(h3)s"><w:name w:val="heading 3"/>
<w:basedOn w:val="%(n)s"/><w:next w:val="%(n)s"/><w:uiPriority w:val="9"/><w:qFormat w:val="1"/>
<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/>
<w:spacing w:before="180" w:after="90" w:line="288" w:lineRule="auto"/>
<w:outlineLvl w:val="2"/></w:pPr>
<w:rPr>%(f_hei)s<w:b w:val="1"/><w:bCs w:val="1"/><w:color w:val="%(ink)s"/>
<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
</w:styles>
""" % {"n": n, "t": t, "s": s, "h1": h1, "h2": h2, "h3": h3,
       "ink": INK, "red": RED,
       "f_son": fonts(SON), "f_hei": fonts(HEI)}


def extract_style(styles_xml, sid):
    """原样取回某个 styleId 的 <w:style> 块（用于兜底搬运本脚本未定义的样式）"""
    m = re.search(r'<w:style [^>]*w:styleId="%s"[^>]*/>' % re.escape(sid), styles_xml)
    if m:
        return m.group(0)
    m = re.search(r'<w:style [^>]*w:styleId="%s"[^>]*>.*?</w:style>'
                  % re.escape(sid), styles_xml, re.S)
    return m.group(0) if m else None


def main():
    z = zipfile.ZipFile(P)
    names = z.namelist()
    data = {k: z.read(k) for k in names}
    z.close()

    doc = data["word/document.xml"].decode("utf-8")
    n = doc.count(TOC_OLD)
    doc = doc.replace(TOC_OLD, TOC_NEW)
    data["word/document.xml"] = doc.encode("utf-8")
    print("目录占位替换:", n)

    cur = data["word/styles.xml"].decode("utf-8")
    ids = {}
    for nm in ("Normal", "Title", "Subtitle", "heading 1", "heading 2", "heading 3"):
        sid = style_id_of(cur, nm)
        if sid:
            ids[nm] = sid
    print("样式 id:", ids)
    missing = [k for k in ("Normal", "Title", "Subtitle", "heading 1", "heading 2")
               if k not in ids]
    if missing:
        print("  ! 缺样式，保持原样不动:", missing)
    else:
        new_sty = build_styles(ids)
        # 安全闸：正文里所有 <w:pStyle w:val="…"/> 都必须能在新 styles.xml 里找到定义。
        # 否则标题会静默掉样式（整篇变成正文外观），而且极难反查。
        used = set(re.findall(r'<w:pStyle w:val="([^"]+)"/>', doc))
        defined = set(re.findall(r'w:styleId="([^"]+)"', new_sty))
        lost = sorted(used - defined)
        if lost:
            # 本脚本未定义的样式（heading 4、TOC 各级、表格样式…）原样搬过来，避免整篇掉版
            carried = [extract_style(cur, sid) for sid in lost]
            carried = [b for b in carried if b]
            if carried:
                new_sty = new_sty.replace("</w:styles>", "".join(carried) + "</w:styles>")
                print("  搬运原有样式 %d 个: %s" % (len(carried), lost))
                defined = set(re.findall(r'w:styleId="([^"]+)"', new_sty))
                lost = sorted(used - defined)
        if lost:
            raise SystemExit("  ! 中止：正文引用了新 styles.xml 里不存在的样式 %s" % lost)
        data["word/styles.xml"] = new_sty.encode("utf-8")
        print("styles.xml 已重写（正文引用 %d 个样式全部有定义）" % len(used))

    with zipfile.ZipFile(P, "w", zipfile.ZIP_DEFLATED) as zo:
        for k in names:
            zo.writestr(k, data[k])

    zz = zipfile.ZipFile(P)
    print("zip testzip:", zz.testzip(), "| entries:", len(zz.namelist()))
    print("已更新:", P)


if __name__ == "__main__":
    main()
