#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 出厂巡检：填充率 / 越界 / 空白页 / 缩略图总览

用法：
    python check.py book.pdf [--sign 署名关键字] [--top 48] [--skip-cover]
                             [--thumb overview.pdf] [--cols 6] [--rows 5]

判定建议：
    填充率 < 10%  -> "孤行页"，必须修（内容搬家 > 删内容 > 拉伸留白）
    越界块        -> 压住页脚或横向溢出，常见原因是 position:relative 的
                     提示框被跨页分片，给它加 page-break-inside:avoid
"""
import argparse
import glob
import os
import sys

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24.3
except ImportError:
    import fitz


def content_box(doc, sign, skip_cover, top):
    """用页脚文本块反推正文内容框的上下边界。"""
    start = 1 if skip_cover else 0
    for i in range(start, len(doc)):
        for b in doc[i].get_text("blocks"):
            if sign and sign in b[4]:
                return b[1] - 4, top
    # 找不到署名时退回经验值（A4 纵向，15mm 下边距）
    return 297 / 25.4 * 72 - 46, top


def scan(doc, sign, bottom, top, right_edge):
    rows = []
    for i, p in enumerate(doc):
        blocks = [b for b in p.get_text("blocks") if b[4].strip()]
        body = [b for b in blocks if b[1] > top and not (sign and sign in b[4])]
        if not body:
            rows.append((i + 1, 0.0, 0, []))
            continue
        fill = (max(b[3] for b in body) - top) / (bottom - top) * 100
        over = []
        for b in body:
            if b[3] > bottom + 1:
                over.append(("下溢", round(b[3] - bottom, 1), b[4].strip()[:24]))
            if right_edge and b[2] > right_edge + 1:
                over.append(("右溢", round(b[2] - right_edge, 1), b[4].strip()[:24]))
        rows.append((i + 1, fill, len(body), over))
    return rows


def contact_sheet(doc, out_path, cols, rows_n, dpi=90):
    """拼一张缩略图总览 PDF，便于一眼看整体版式与异常空白。"""
    W, H = 595, 842
    out = fitz.open()
    per = cols * rows_n
    idx = 0
    while idx < len(doc):
        page = out.new_page(width=W, height=H)
        cw, ch = W / cols, H / rows_n
        for k in range(per):
            if idx + k >= len(doc):
                break
            r, c = divmod(k, cols)
            rect = fitz.Rect(c * cw + 4, r * ch + 4,
                             (c + 1) * cw - 4, (r + 1) * ch - 4)
            page.insert_image(rect, pixmap=doc[idx + k].get_pixmap(dpi=dpi))
        idx += per
    out.save(out_path, deflate=True, garbage=3)
    out.close()


def find_browser():
    cands = glob.glob(os.path.expanduser(
        "~/Library/Caches/ms-playwright/chromium_headless_shell-*/"
        "chrome-headless-shell-mac-arm64/chrome-headless-shell"))
    cands += [
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/usr/bin/chromium", "/usr/bin/google-chrome",
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--sign", default="", help="页脚署名关键字，用于反推内容框底边")
    ap.add_argument("--top", type=float, default=48.0, help="页眉下沿 y（pt）")
    ap.add_argument("--skip-cover", action="store_true", help="第 1 页是封面则跳过")
    ap.add_argument("--thumb", default="", help="输出缩略图总览 PDF 路径")
    ap.add_argument("--cols", type=int, default=6)
    ap.add_argument("--rows", type=int, default=5)
    a = ap.parse_args()

    doc = fitz.open(a.pdf)
    bottom, top = content_box(doc, a.sign, a.skip_cover, a.top)
    right_edge = doc[0].rect.width - 42.5  # 15mm 右边距
    rows = scan(doc, a.sign, bottom, top, right_edge)

    print(f"{a.pdf}  共 {len(doc)} 页   内容框 y={top:.0f}~{bottom:.0f}")
    print("-" * 58)
    thin = over = 0
    for pno, fill, nblocks, ov in rows:
        flag = ""
        # ⚠️ 判据是 fill < 10，**不能写成 0 < fill < 10**：
        #    完全空白的页 body 为空 → fill 恰为 0.0，被 `0 <` 排除后就静默漏检，
        #    而空白页恰恰是"孤行页"里最该修的一种（实测：末页全空仍返回退出码 0）。
        if fill < 10:
            flag = "  <-- 孤行页"
            thin += 1
        if ov:
            flag += "  <-- 越界"
            over += 1
        print(f"  p{pno:<3} 填充 {fill:5.1f}%  块 {nblocks:<3}{flag}")
        for kind, dx, txt in ov:
            print(f"         {kind} {dx}pt : {txt}")
    print("-" * 58)
    print(f"孤行页 {thin} 个 / 越界页 {over} 个")
    if a.thumb:
        contact_sheet(doc, a.thumb, a.cols, a.rows)
        print(f"缩略图总览 -> {a.thumb}")
    doc.close()
    return 1 if (thin or over) else 0


if __name__ == "__main__":
    sys.exit(main())
