# -*- coding: utf-8 -*-
"""把 part1.html + part2.html 转成干净 Markdown（供 Word 使用）"""
import os, re, sys, tempfile
from lxml import html as LH

# 可用环境变量覆盖，默认落在临时目录，不写死任何本机路径
BUILD = os.environ.get("EXAM_BUILD") or os.path.join(
    tempfile.gettempdir(), "exam_docx_work", "build")
OUT = os.environ.get("EXAM_MD_OUT") or os.path.join(
    os.path.dirname(BUILD), "word", "manual.md")

CHART_ORDER = ["c2_towns", "c3_structure", "c4_clusters", "c1_targets"]


def txt(el):
    """取元素纯文本，压缩空白"""
    s = el.text_content()
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[^\S\u3000]*\n[^\S\u3000]*", " ", s)
    s = re.sub(r"[^\S\u3000]{2,}", " ", s)
    return s.strip()


def inline(el):
    """把元素内的行内内容转成 markdown 行内文本（保留 strong/b/em/code）"""
    parts = []
    if el.text:
        parts.append(el.text)
    for ch in el:
        tag = ch.tag if isinstance(ch.tag, str) else ""
        inner = inline(ch)
        if tag in ("strong", "b"):
            inner = "**%s**" % inner.strip() if inner.strip() else ""
        elif tag in ("em", "i"):
            inner = "*%s*" % inner.strip() if inner.strip() else ""
        elif tag == "br":
            inner = "　"
        elif tag == "span":
            cls = ch.get("class") or ""
            style = ch.get("style") or ""
            # 隐藏的白色锚点，丢弃
            if re.search(r"color:\s*#(?:fff|ffffff)\b", style, re.I) and "6pt" in style:
                inner = ""
            elif "tag" in cls:
                inner = "【%s】" % inner.strip()
            elif "g" in cls.split():
                inner = "**%s**" % inner.strip()
        elif tag in ("div", "p", "ul", "table", "h4"):
            # 块级元素出现在行内位置，少见；用空格隔开
            inner = " " + inner + " "
        parts.append(inner)
        if ch.tail:
            parts.append(ch.tail)
    s = "".join(parts)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[^\S\u3000]+", " ", s)
    s = re.sub(r"\*\*[ \t\r\n]*\*\*", "", s)
    s = re.sub(r"\*\*([^*]+?)\s*\*\*", lambda m: "**%s**" % m.group(1).strip(), s)
    return s.strip()


def table_md(tbl):
    rows = []
    for tr in tbl.iter("tr"):
        cells = [inline(td).replace("|", "／") for td in tr if td.tag in ("td", "th")]
        if cells:
            rows.append(cells)
    if not rows:
        return []
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |",
           "|" + "|".join([" --- "] * ncol) + "|"]
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return out


def tl_md(tl):
    """时间轴容器 → 转成「年份 × 字段」两维表。

    坑：`.tl` 里没有任何一段可读文本，每个字段都埋在叶节点 `div.it > span.lk/.lv`
    里，靠通用的"段落/列表/纯文本 div"分支一条都收不到 —— 整条时间轴在 Word 版里
    会**静默消失**（HTML 版还在，肉眼对比才发现）。必须显式展开成表。
    """
    cols = (tl.findall(".//div[@class='col']") + tl.findall(".//div[@class='col next']")
            or [c for c in tl.iter("div") if "col" in (c.get("class") or "").split()])
    keys, data = [], []
    for c in cols:
        yr = c.find("div[@class='yr']")
        items = []
        for it in c.findall(".//div[@class='it']"):
            lk = it.find("span[@class='lk']")
            lv = it.find("span[@class='lv']")
            k = inline(lk) if lk is not None else ""
            items.append((k, inline(lv) if lv is not None else ""))
            if k and k not in keys:
                keys.append(k)
        data.append((inline(yr) if yr is not None else "", dict(items)))
    if not keys or not data:
        return []
    rows = [["年份"] + keys]
    for yr, d in data:
        rows.append([yr] + [d.get(k, "") for k in keys])
    out = ["| " + " | ".join(rows[0]) + " |",
           "|" + "|".join([" --- "] * len(keys and rows[0])) + "|"]
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    out.append("")
    return out


def blocks(el, out):
    """递归遍历，输出 markdown 行"""
    for ch in el:
        tag = ch.tag if isinstance(ch.tag, str) else ""
        cls = (ch.get("class") or "").split()

        if tag in ("script", "style"):
            continue

        # --- 图表占位 ---
        if tag == "div" and "figure" in cls:
            ph = ch.find(".//div[@class='chartph']")
            if ph is not None:
                out.append("[[CHART:%s]]" % ph.get("data-k"))
            caps = [c for c in ch.iter("div") if "cap" in (c.get("class") or "")]
            cap = txt(caps[0]) if caps else ""
            if cap:
                out.append("*%s*" % cap)
            out.append("")
            continue

        # --- 章标题 ---
        if tag == "div" and "chapter-head" in cls:
            h2 = ch.find(".//h2")
            num = ch.find(".//div[@class='num']")
            pre = ""
            if num is not None:
                n = txt(num)
                CN = "零一二三四五六七八九十"
                if n.isdigit():
                    i = int(n)
                    if i <= 10:
                        pre = "第%s章　" % CN[i]
                    elif i < 20:
                        pre = "第十%s章　" % CN[i - 10]
                    else:
                        pre = "第%s十章　" % CN[i // 10]
            if h2 is not None:
                out.append("# " + pre + inline(h2))
                out.append("")
            continue

        # --- 时间轴（.tl 容器：文字全在叶子 div，容器本身没有直接文本 —— 少写这支整段全丢） ---
        if tag == "div" and "tl" in cls:
            rows = [r for r in ch.findall("div")
                    if "row" in (r.get("class") or "").split()]
            for row in rows:
                d = row.find(".//div[@class='d']")
                if d is None:
                    d = row.find(".//div[@class='tl-d']")
                m = row.find(".//span[@class='m']")
                e = row.find(".//div[@class='e']")
                if e is None:
                    e = row.find(".//div[@class='tl-b']")
                dt = txt(d) if d is not None else ""
                ev = inline(e) if e is not None else ""
                if m is not None and txt(m):
                    ev = "〔%s〕%s" % (txt(m), ev)
                if dt or ev:
                    out.append("- **%s**　%s" % (dt, ev))
            out.append("")
            continue

        # --- 提示框 / 框架框 ---
        if tag == "div" and ("tip" in cls or "warn" in cls or "key" in cls):
            ttl = ch.find(".//span")
            ttl_txt = txt(ttl) if ttl is not None else ""
            whole = inline(ch)
            if ttl_txt and whole.startswith(ttl_txt):
                whole = whole[len(ttl_txt):].strip()
            if ttl_txt:
                out.append("> **%s** %s" % (ttl_txt, whole))
            elif whole:
                out.append("> " + whole)
            out.append("")
            continue

        if tag == "div" and "framework" in cls:
            ft = ch.find(".//div[@class='fttl']")
            if ft is not None:
                out.append("> **%s**" % txt(ft))
                for p in ch.findall("p"):
                    s = inline(p)
                    if s:
                        out.append("> " + s)
                out.append("")
            continue

        # --- 步骤条 ---
        if tag == "div" and "steps" in cls:
            sts = [s for s in ch.findall("div")
                   if "st" in (s.get("class") or "").split()]
            # 首条以「N 月」开头的（月度大事记）改用项目符号，避免与序号混淆
            use_bullet = bool(sts and re.match(r"\s*\d{1,2}\s*月", txt(sts[0])))
            n = 1
            for st in sts:
                s = inline(st).replace("\n", "　")
                s = re.sub(r"[ \t]*　[ \t]*", "　", s).strip()
                if use_bullet:
                    out.append("- " + s)
                else:
                    out.append("%d. %s" % (n, s))
                    n += 1
            out.append("")
            continue

        # --- 速记卡（12 格）---
        if tag == "div" and "quick" in cls:
            rows = []
            for card in ch.findall("div"):
                v = card.find(".//div[@class='v']")
                k = card.find(".//div[@class='k']")
                if v is not None and k is not None:
                    rows.append((txt(k), txt(v)))
            if rows:
                out.append("| 项目 | 数值 |")
                out.append("| --- | --- |")
                for k, v in rows:
                    out.append("| %s | **%s** |" % (k, v))
                out.append("")
            continue

        # --- 卡片 ---
        if tag == "div" and "card" in cls:
            blocks(ch, out)
            continue

        # --- 表格 ---
        if tag == "table":
            out.extend(table_md(ch))
            out.append("")
            continue

        # --- 标题 ---
        if tag == "h2":
            out.append("# " + inline(ch))
            out.append("")
            continue
        if tag == "h3":
            out.append("## " + inline(ch))
            out.append("")
            continue
        if tag == "h4":
            # 卡片内 h4 可能带 float 的负责人信息
            leader = ""
            for sp in ch.iter("span"):
                if "float" in (sp.get("style") or ""):
                    leader = txt(sp)
            head = ch.text or ""
            for sp in ch.iter("span"):
                if "float" not in (sp.get("style") or ""):
                    head += sp.text_content()
            head = re.sub(r"\s+", " ", head).strip()
            out.append("### " + head)
            out.append("")
            if leader:
                out.append("**负责人：** %s" % leader)
                out.append("")
            continue

        # --- 列表 ---
        if tag == "ul":
            for li in ch.findall("li"):
                s = inline(li)
                s = re.sub(r"\s+", " ", s).strip()
                if s:
                    out.append("- " + s)
            out.append("")
            continue

        # --- 段落 ---
        if tag == "p":
            s = inline(ch)
            if s:
                out.append(s)
                out.append("")
            continue

        # --- 纯卡片组 / 其他容器：递归 ---
        if tag in ("div", "section", "body", "html"):
            if "toc" in cls:
                continue
            if "readme" in cls:
                blocks(ch, out)
                continue
            if "docend" in cls:
                out.append("---")
                out.append("")
                out.append("**%s**" % txt(ch).split("广东中公")[0].strip())
                out.append("")
                out.append("**广东中公教研**")
                out.append("")
                continue
            if "grid2" in cls or "grid3" in cls:
                blocks(ch, out)
                continue
            if "tl" in cls:
                out.extend(tl_md(ch))
                continue
            if "t" in cls or "cn" in cls or "num" in cls or "cap" in cls:
                continue
            blocks(ch, out)
            continue


def main():
    parts = []
    for fn in ("part1.html", "part2.html"):
        with open("%s/%s" % (BUILD, fn), encoding="utf-8") as f:
            raw = f.read()
        raw = re.sub(r"<!--CHART:([A-Za-z0-9_]+)-->",
                     lambda m: '<div class="chartph" data-k="%s"></div>' % m.group(1),
                     raw)
        doc = LH.document_fromstring(raw)
        out = []
        blocks(doc.body, out)
        parts.append("\n".join(out))

    md = "\n".join(parts)
    md = md.replace("# 关于本手册", "# 编制说明　关于本手册", 1)
    md = re.sub(r"\n{3,}", "\n\n", md)
    md = md.strip() + "\n"

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(md)

    print("written:", OUT)
    print("chars:", len(md))
    print("charts:", md.count("[[CHART:"))
    print("h1:", len(re.findall(r"^# ", md, re.M)))
    print("h2:", len(re.findall(r"^## ", md, re.M)))
    print("h3:", len(re.findall(r"^### ", md, re.M)))
    print("tables:", len(re.findall(r"^\| ", md, re.M)))


if __name__ == "__main__":
    main()
