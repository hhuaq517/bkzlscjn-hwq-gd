#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""渲染讲义为 PDF：封面全出血页 + 正文（页眉页脚/页码） + 目录页码回填 + 合并。"""
import asyncio, re, sys, os, html as _html
import fitz  # PyMuPDF
from playwright.async_api import async_playwright

BUILD = os.path.dirname(os.path.abspath(__file__))
FONT_H = "'Hiragino Sans GB','STHeiti','Heiti SC','PingFang SC',sans-serif"

def cover_html(c):
    chips = "".join(f'<span class="chip">{_html.escape(x)}</span>' for x in c.get("chips", []))
    desc = "<br>".join(_html.escape(x) for x in c.get("desc", []))
    return f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><style>
*{{margin:0;padding:0;box-sizing:border-box;}}
@page{{size:A4;margin:0;}}
html,body{{width:210mm;height:297mm;overflow:hidden;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
.cover{{position:relative;width:210mm;height:297mm;overflow:hidden;color:#fff;
  background:radial-gradient(900px 520px at 84% -8%, #1E5B8A 0%, rgba(30,91,138,0) 62%),
  linear-gradient(158deg,#0B2740 0%,#12395C 56%,#0D2E4C 100%);}}
.grid{{position:absolute;inset:0;opacity:.15;
  background-image:linear-gradient(#ffffff22 1px,transparent 1px),linear-gradient(90deg,#ffffff22 1px,transparent 1px);
  background-size:26mm 26mm;}}
.glow{{position:absolute;width:150mm;height:150mm;right:-40mm;top:-30mm;border-radius:50%;
  background:radial-gradient(circle,#B8842B66 0%,rgba(184,132,43,0) 68%);}}
.glow2{{position:absolute;width:120mm;height:120mm;left:-46mm;bottom:10mm;border-radius:50%;
  background:radial-gradient(circle,#2C7A7240 0%,rgba(44,122,114,0) 70%);}}
.inner{{position:relative;height:100%;padding:38mm 24mm 24mm;display:flex;flex-direction:column;}}
.kicker{{font-family:{FONT_H};font-size:11pt;letter-spacing:6px;color:#E9D9B6;}}
.kicker::after{{content:"";display:block;width:58px;height:3px;background:#B8842B;margin-top:12px;}}
h1{{font-family:{FONT_H};font-size:38pt;line-height:1.28;margin-top:30px;font-weight:800;letter-spacing:1px;}}
h1 .sub{{display:block;font-size:21pt;font-weight:400;color:#CFE0EF;margin-top:16px;letter-spacing:2px;}}
.divider{{width:130px;height:5px;background:linear-gradient(90deg,#B8842B,#E9D9B6);margin:32px 0;border-radius:3px;}}
.desc{{font-family:{FONT_H};font-size:12pt;color:#D8E5F0;line-height:2.05;max-width:135mm;}}
.meta{{margin-top:auto;display:flex;gap:12px;flex-wrap:wrap;}}
.chip{{border:1px solid #ffffff55;border-radius:20px;padding:7px 17px;font-size:10pt;color:#E6EEF6;font-family:{FONT_H};}}
.band{{position:absolute;left:0;right:0;bottom:0;height:17mm;background:linear-gradient(90deg,#B8842B,#E9D9B6 58%,#B8842B);}}
.band span{{position:absolute;right:24mm;bottom:5.4mm;font-size:8.5pt;color:#3A2A0E;letter-spacing:3px;font-family:{FONT_H};}}
.band .left{{left:24mm;right:auto;letter-spacing:4px;}}
</style></head><body>
<div class="cover"><div class="grid"></div><div class="glow"></div><div class="glow2"></div>
<div class="inner">
  <div class="kicker">{_html.escape(c["kicker"])}</div>
  <h1>{_html.escape(c["title"])}<span class="sub">{_html.escape(c["subtitle"])}</span></h1>
  <div class="divider"></div>
  <div class="desc">{desc}</div>
  <div class="meta">{chips}</div>
</div>
<div class="band"><span class="left">{_html.escape(c.get("band_l",""))}</span><span>{_html.escape(c.get("band_r","广东中公教研"))}</span></div>
</div></body></html>"""

def read_toc_keys(body_path):
    with open(body_path, encoding="utf-8") as f:
        s = f.read()
    return re.findall(r'<span class="pg" data-s="([^"]*)">', s)

def find_pages(body_pdf, keys, skip_marker="导航页"):
    doc = fitz.open(body_pdf)
    skip = set()
    for i in range(len(doc)):
        if skip_marker and doc[i].search_for(skip_marker):
            skip.add(i + 1)
    out = []
    for k in keys:
        page_no = None
        for i in range(len(doc)):
            if (i + 1) in skip:
                continue
            if doc[i].search_for(k):
                page_no = i + 1
                break
        out.append(page_no if page_no else 0)
    doc.close()
    return out

async def run(cfg):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        # ---- 封面 ----
        cov = os.path.join(BUILD, cfg["name"] + "_cover.html")
        with open(cov, "w", encoding="utf-8") as f:
            f.write(cover_html(cfg["cover"]))
        await page.goto("file://" + cov)
        await page.wait_for_timeout(400)
        await page.pdf(path=cfg["cover_pdf"], width="210mm", height="297mm",
                       margin={"top":"0","bottom":"0","left":"0","right":"0"}, print_background=True)
        # ---- 正文 ----
        body_url = "file://" + os.path.abspath(cfg["body_html"])
        await page.goto(body_url)
        await page.wait_for_timeout(600)
        header = f"""<div style="width:100%;font-family:{FONT_H};font-size:7.5pt;color:#9AA9B8;padding:0 15mm;display:flex;justify-content:space-between;align-items:center;">
        <span style="letter-spacing:1px;">{cfg['header_l']}</span><span style="letter-spacing:2px;">{cfg['header_r']}</span></div>"""
        footer = f"""<div style="width:100%;font-family:{FONT_H};font-size:7.5pt;color:#9AA9B8;padding:0 15mm;display:flex;justify-content:space-between;align-items:center;">
        <span style="letter-spacing:2px;">广东中公教研</span>
        <span>第 <span class="pageNumber"></span> 页 / 共 <span class="totalPages"></span> 页</span></div>"""
        async def render(pdf_path):
            await page.pdf(path=pdf_path, format="A4", print_background=True,
                           margin={"top":"17mm","bottom":"16mm","left":"15mm","right":"15mm"},
                           display_header_footer=True, header_template=header, footer_template=footer)
        tmp = cfg["body_pdf"].replace(".pdf", "_pass1.pdf")
        await render(tmp)
        # ---- 回填目录页码 ----
        keys = read_toc_keys(cfg["body_html"])
        nums = find_pages(tmp, keys) if keys else []
        if nums:
            await page.evaluate("(nums)=>{const els=[...document.querySelectorAll('.toc .pg')];els.forEach((el,i)=>{if(nums[i])el.textContent=nums[i];});}", nums)
            print("TOC:", list(zip(keys, nums)))
        await render(cfg["body_pdf"])
        # ---- 合并 ----
        out = fitz.open()
        out.insert_pdf(fitz.open(cfg["cover_pdf"]))
        out.insert_pdf(fitz.open(cfg["body_pdf"]))
        out.save(cfg["out_pdf"], deflate=True, garbage=3)
        out.close()
        n = fitz.open(cfg["out_pdf"]).page_count
        print(f"[OK] {cfg['out_pdf']}  共 {n} 页")
        await browser.close()

if __name__ == "__main__":
    import json
    with open(os.path.join(BUILD, "docs.json"), encoding="utf-8") as f:
        docs = json.load(f)
    only = sys.argv[1] if len(sys.argv) > 1 else None
    for cfg in docs:
        if only and cfg["name"] != only:
            continue
        asyncio.run(run(cfg))
