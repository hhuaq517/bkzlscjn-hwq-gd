#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
轻量管线：单文件 HTML -> PDF -> 页脚署名 + 页码盖章 -> 可选目录页码回填

适用：没有 Playwright、或不需要"封面全出血单独渲染"的简单文档。
封面页（第 1 页）自动跳过盖章。

用法：
    python stamp.py in.html out.pdf [--sign "署名文字"] [--title "页眉文字"]
                    [--skip-stamp] [--toc-map anchors.json]

--toc-map: JSON，{"目录占位符": "正文锚点首句"}；先定位占位符所在页，
           再用锚点首句定位真实页码，最后替换占位符并重新渲染一次。
"""
import argparse
import glob
import json
import os
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24.3
except ImportError:
    import fitz


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
    sys.exit("未找到无头浏览器（chrome-headless-shell / Edge / Chrome）")


def file_url(path):
    """file:// 路径中的空格和中文必须编码，否则打不开。"""
    return "file://" + urllib.parse.quote(os.path.abspath(path))


def to_pdf(html_path, pdf_path, browser):
    cmd = [browser, "--headless", "--disable-gpu", "--no-pdf-header-footer",
           "--no-sandbox", f"--print-to-pdf={pdf_path}", file_url(html_path)]
    subprocess.run(cmd, check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def stamp(pdf_path, sign="", title="", skip_first=True):
    """盖章页脚：细分隔线 + 左署名 + 右页码。china-ss 为内置中文黑体。"""
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc):
        if skip_first and i == 0:
            continue
        r = page.rect
        y = r.height - 30
        page.draw_line(fitz.Point(r.width * 0.06, y - 12),
                       fitz.Point(r.width * 0.94, y - 12),
                       color=(0.78, 0.81, 0.85), width=0.6)
        if sign:
            page.insert_textbox(fitz.Rect(r.width * 0.06, y, r.width * 0.62, y + 14),
                                sign, fontname="china-ss", fontsize=8,
                                color=(0.42, 0.47, 0.53), align=0)
        label = f"{i} / {len(doc) - 1}" if skip_first else f"{i + 1} / {len(doc)}"
        page.insert_textbox(fitz.Rect(r.width * 0.62, y, r.width * 0.94, y + 14),
                            label, fontname="china-ss", fontsize=8,
                            color=(0.42, 0.47, 0.53), align=2)
    doc.saveIncr()
    doc.close()


def toc_fill(html_path, pdf_path, mapping):
    """用正文锚点首句定位真实页码，返回 {占位符: 页码}；未命中为 0。"""
    doc = fitz.open(pdf_path)
    result = {}
    for placeholder, anchor in mapping.items():
        hit = 0
        for i, page in enumerate(doc):
            if page.search_for(anchor):
                hit = i + 1
                break
        result[placeholder] = hit
    doc.close()
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html")
    ap.add_argument("out")
    ap.add_argument("--sign", default="", help="页脚左侧署名文字")
    ap.add_argument("--title", default="", help="页眉文字（可选，需 HTML 预留）")
    ap.add_argument("--skip-stamp", action="store_true", help="只渲染不盖章")
    ap.add_argument("--toc-map", default="", help="JSON 文件路径")
    ap.add_argument("--browser", default="")
    a = ap.parse_args()

    browser = a.browser or find_browser()
    out = os.path.abspath(a.out)
    os.makedirs(os.path.dirname(out), exist_ok=True)

    html_path = os.path.abspath(a.html)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_pdf = os.path.join(tmp, "raw.pdf")
        to_pdf(html_path, tmp_pdf, browser)

        mapping = {}
        if a.toc_map:
            with open(a.toc_map, encoding="utf-8") as f:
                mapping = json.load(f)
            nums = toc_fill(html_path, tmp_pdf, mapping)
            bad = [k for k, v in nums.items() if not v]
            # 回填：把 HTML 里的占位符替换成页码，重渲染一次
            with open(html_path, encoding="utf-8") as f:
                src = f.read()
            for ph, pg in nums.items():
                src = src.replace(ph, str(pg) if pg else "—")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(src)
            print("TOC:", nums, ("未命中 -> " + ", ".join(bad)) if bad else "")
            to_pdf(html_path, tmp_pdf, browser)

        if not a.skip_stamp:
            stamp(tmp_pdf, a.sign, a.title)
        doc = fitz.open(tmp_pdf)
        doc.save(out, deflate=True, garbage=3)
        n = doc.page_count
        doc.close()

    print(f"[OK] {out}  共 {n} 页  (browser: {os.path.basename(browser)})")


if __name__ == "__main__":
    main()
