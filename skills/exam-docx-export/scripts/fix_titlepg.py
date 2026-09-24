# -*- coding: utf-8 -*-
"""给成品 docx 加「首页不同」：封面不显示页眉与页码（与 PDF 封面一致）

用法：python fix_titlepg.py <目标.docx>

三个坑（都踩过）：
  1. editor_sdk **已经生成** `footerReference w:type="first"`，且指向带 PAGE 域的
     `footer1.xml`。只加 `<w:titlePg/>` 而不重指，封面页码照旧出现。
  2. editor_sdk 产出的 `<w:footerReference r:id="…" w:type="first"/>` 属性顺序是
     **r:id 在前、w:type 在后**。正则写成 `<w:footerReference w:type="first" r:id="…"/>`
     会**静默匹配不到**（不报错、封面页码仍在）→ 必须用顺序无关的写法。
  3. 新增部件后要做「悬空关系 / 未声明引用」两道校验，否则 Word 报"内容有问题"。
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

P = sys.argv[1] if len(sys.argv) > 1 else ""
if not P:
    sys.exit("用法：python fix_titlepg.py <目标.docx>")
_T = tempfile.gettempdir()
TMP = os.path.join(_T, "_titlepg_repack")
OUT = os.path.join(_T, "_titlepg_out.docx")

EMPTY_HDR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
             '<w:p/></w:hdr>')
EMPTY_FTR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
             '<w:p/></w:ftr>')

z = zipfile.ZipFile(P)
names = z.namelist()
data = {n: z.read(n) for n in names}
z.close()

doc = data["word/document.xml"].decode("utf-8")
if "<w:titlePg/>" in doc:
    print("已有 titlePg，跳过（若要重跑请先删掉 <w:titlePg/>）")
    raise SystemExit(0)


def _free(prefix, ext):
    i = 1
    while ("word/%s%d.%s" % (prefix, i, ext)) in names:
        i += 1
    return "%s%d.%s" % (prefix, i, ext)


hdr_name = _free("header", "xml")
ftr_name = _free("footer", "xml")
print("新建部件:", hdr_name, ftr_name)

rels = data["word/_rels/document.xml.rels"].decode("utf-8")
used = set(re.findall(r'Id="([^"]+)"', rels))
rid_h = next("rId%d" % i for i in range(200, 400) if "rId%d" % i not in used)
used.add(rid_h)
rid_f = next("rId%d" % i for i in range(200, 400) if "rId%d" % i not in used)

data["word/" + hdr_name] = EMPTY_HDR.encode("utf-8")
data["word/" + ftr_name] = EMPTY_FTR.encode("utf-8")

rels = rels.replace(
    "</Relationships>",
    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/header" Target="%s" />'
    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/footer" Target="%s" /></Relationships>'
    % (rid_h, hdr_name, rid_f, ftr_name))
data["word/_rels/document.xml.rels"] = rels.encode("utf-8")

ct = data["[Content_Types].xml"].decode("utf-8")
add = ""
if ('PartName="/word/%s"' % hdr_name) not in ct:
    add += ('<Override PartName="/word/%s" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.header+xml"/>' % hdr_name)
if ('PartName="/word/%s"' % ftr_name) not in ct:
    add += ('<Override PartName="/word/%s" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.footer+xml"/>' % ftr_name)
ct = ct.replace("</Types>", add + "</Types>")
data["[Content_Types].xml"] = ct.encode("utf-8")

# ── 顺序无关：把 first 类型的页脚/页眉引用改指到新建的空部件 ──
m = re.search(r'<w:footerReference\b[^>]*w:type="first"[^>]*/>', doc)
if m:
    print("原 first 页脚引用:", m.group(0), "→ rId", rid_f)
    doc = doc.replace(m.group(0),
                      '<w:footerReference w:type="first" r:id="%s"/>' % rid_f)
else:
    print("  ! 文档没有 first 页脚引用，将在 sectPr 里补一个")
    doc = doc.replace("</w:sectPr>",
                      '<w:footerReference w:type="first" r:id="%s"/></w:sectPr>' % rid_f)

m = re.search(r'<w:headerReference\b[^>]*w:type="default"[^>]*/>', doc)
if m:
    doc = doc.replace(m.group(0),
                      '<w:headerReference w:type="first" r:id="%s"/>%s'
                      % (rid_h, m.group(0)))
else:
    print("  ! 未找到 default 页眉引用，first 页眉未加")

doc = doc.replace("</w:sectPr>", "<w:titlePg/></w:sectPr>")
data["word/document.xml"] = doc.encode("utf-8")

order = list(names)
for extra in ("word/" + hdr_name, "word/" + ftr_name):
    if extra not in order:
        order.append(extra)

shutil.rmtree(TMP, ignore_errors=True)
os.makedirs(TMP, exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zo:
    for n in order:
        if n in data:
            zo.writestr(n, data[n])
print("zip testzip:", zipfile.ZipFile(OUT).testzip(),
      "| entries:", len(zipfile.ZipFile(OUT).namelist()))

zz = zipfile.ZipFile(OUT)
nn = set(zz.namelist())
d2 = zz.read("word/document.xml").decode("utf-8")
r2 = zz.read("word/_rels/document.xml.rels").decode("utf-8")
ids = set(re.findall(r'Id="([^"]+)"', r2))
usedr = set(re.findall(r'r:(?:id|embed)="([^"]+)"', d2))
print("未声明引用:", sorted(usedr - ids))
tgts = re.findall(r'Target="([^"]+)"(?![^>]*External)', r2)
print("悬空关系:", [t for t in tgts if ("word/" + t.lstrip("/")) not in nn])
zz.close()

os.replace(OUT, P)
print("已更新:", P)
