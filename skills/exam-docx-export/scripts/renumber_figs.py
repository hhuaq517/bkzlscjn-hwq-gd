#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按章号重编图号：合并章节后，图注里的旧章号需回正为当前章号。
   规则：图 {章号}-{章内序号}，章内序号从 1 起按正文出现顺序计。"""
import re
import sys
import shutil
from docx import Document

SRC = sys.argv[1] if len(sys.argv) > 1 else "gk_new.docx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "gk_new.docx"
shutil.copy(SRC, SRC + ".bak")

CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
      "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def cn_num(s):
    """一 → 1 ；十一 → 11 ；附录A → 'A'"""
    s = s.strip()
    m = re.match(r"^附录\s*([A-Za-z])", s)
    if m:
        return m.group(1).upper()
    m = re.match(r"^第([一二三四五六七八九十]+)章", s)
    if not m:
        return None
    t = m.group(1)
    if t == "十":
        return 10
    if t.startswith("十"):
        return 10 + CN.get(t[1:], 0)
    if "十" in t:
        a, b = t.split("十")
        return CN.get(a, 0) * 10 + (CN.get(b, 0) if b else 0)
    return CN.get(t)


FIG = re.compile(r"^图\s*(\d+)\s*[-–]\s*(\d+)")


def main():
    doc = Document(SRC)
    chap = None
    seq = {}
    changed = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if p.style.name == "Heading 1":
            c = cn_num(t)
            if c is not None:
                chap = c
            continue
        m = FIG.match(t)
        if not m:
            continue
        if chap is None:
            continue
        seq[chap] = seq.get(chap, 0) + 1
        new = f"图 {chap}-{seq[chap]}"
        old = m.group(0)
        if old.replace(" ", "") == new.replace(" ", ""):
            continue
        # 改写 run0（图注的编号整体在首个 run 内）
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new, 1)
                changed.append((old, new, t[:46]))
                break
        else:
            # 兜底：编号被拆散，整段重建到首个 run
            r0 = p.runs[0]
            r0.text = new + t[len(old):]
            for r in p.runs[1:]:
                r.text = ""
            changed.append((old, new, t[:46] + "  [重建]"))

    doc.save(OUT)
    print(f"已保存 {OUT}")
    print(f"重编 {len(changed)} 处：")
    for o, n, ctx in changed:
        print(f"   {o:10s} → {n:10s}   {ctx}")


if __name__ == "__main__":
    main()
