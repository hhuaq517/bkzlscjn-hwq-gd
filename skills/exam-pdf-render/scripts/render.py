#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用印刷级 PDF 渲染器（Playwright 双通道）
  封面全出血页  +  正文（页眉/页脚/页码）  +  目录页码回填  +  合并

用法：
    python render.py docs.json [name]        # 只渲染指定 name，省略则全部

环境：python3（需 playwright + pymupdf，Chromium 随 playwright 一并安装）

docs.json 为数组，每项字段见文件末尾 SCHEMA 常量。
所有文案（署名 / 页眉 / 封面）均来自配置，脚本内不硬编码任何机构名。
"""
import asyncio
import html as _html
import json
import os
import re
import sys

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24.3
except ImportError:
    import fitz
from playwright.async_api import async_playwright

FONT_H = "'Hiragino Sans GB','STHeiti','Heiti SC','PingFang SC',sans-serif"

# ---------------------------------------------------------------- 配色预设
PALETTES = {
    "gov": {  # 政务 / 综合岗：藏青 + 琥珀金 + 松石绿
        "bg1": "#0B2740", "bg2": "#12395C", "bg3": "#0D2E4C",
        "highlight": "#1E5B8A", "accent": "#B8842B", "accent_soft": "#E9D9B6",
        "glow1": "#B8842B", "glow2": "#2C7A72",
        "kicker": "#E9D9B6", "sub": "#CFE0EF", "desc": "#D8E5F0",
        "band_a": "#B8842B", "band_b": "#E9D9B6", "band_text": "#3A2A0E",
    },
    "edu": {  # 教师 / 教育类：墨绿 + 暖橙
        "bg1": "#123B30", "bg2": "#1F5C4A", "bg3": "#164B3C",
        "highlight": "#2E7D65", "accent": "#C4703A", "accent_soft": "#EFD9C4",
        "glow1": "#C4703A", "glow2": "#7FA88C",
        "kicker": "#EFD9C4", "sub": "#CBE0D6", "desc": "#D9E8E1",
        "band_a": "#C4703A", "band_b": "#EFD9C4", "band_text": "#3A2412",
    },
    "med": {  # 医疗 / 卫健：青蓝 + 珊瑚
        "bg1": "#0A3C44", "bg2": "#106B7A", "bg3": "#0C525E",
        "highlight": "#17899B", "accent": "#D2604F", "accent_soft": "#F3D2CB",
        "glow1": "#D2604F", "glow2": "#5FA9B8",
        "kicker": "#F3D2CB", "sub": "#C6E4EA", "desc": "#D8EDEF",
        "band_a": "#D2604F", "band_b": "#F3D2CB", "band_text": "#3A1912",
    },
    "fin": {  # 金融 / 财经：深蓝 + 金
        "bg1": "#0E2144", "bg2": "#1B3A6B", "bg3": "#142C55",
        "highlight": "#2A5090", "accent": "#B58A2B", "accent_soft": "#EBDDB2",
        "glow1": "#B58A2B", "glow2": "#4E6E9E",
        "kicker": "#EBDDB2", "sub": "#CBD9EE", "desc": "#DCE5F2",
        "band_a": "#B58A2B", "band_b": "#EBDDB2", "band_text": "#33270C",
    },
    "plain": {  # 素雅 / 通用：石墨 + 砖红
        "bg1": "#23282E", "bg2": "#39414A", "bg3": "#2B3138",
        "highlight": "#55606C", "accent": "#A65D46", "accent_soft": "#E8D6CE",
        "glow1": "#A65D46", "glow2": "#7C8794",
        "kicker": "#E8D6CE", "sub": "#D2D8DE", "desc": "#DEE3E8",
        "band_a": "#A65D46", "band_b": "#E8D6CE", "band_text": "#33190F",
    },
}

DEFAULT_MARGIN = {"top": "17mm", "bottom": "16mm", "left": "15mm", "right": "15mm"}
DEFAULT_NAV_MARKER = "导航页"


# ---------------------------------------------------------------- 封面
def cover_html(cfg):
    cov = cfg["cover"]
    pal = dict(PALETTES.get(cfg.get("palette", "gov"), PALETTES["gov"]))
    pal.update(cfg.get("palette_override") or {})

    chips = "".join(
        f'<span class="chip">{_html.escape(x)}</span>' for x in cov.get("chips", [])
    )
    desc = "<br>".join(_html.escape(x) for x in cov.get("desc", []))
    band_l = _html.escape(cov.get("band_l", ""))
    band_r = _html.escape(cov.get("band_r", cfg.get("footer_l", "")))

    css = f"""*{{margin:0;padding:0;box-sizing:border-box;}}
@page{{size:A4;margin:0;}}
html,body{{width:210mm;height:297mm;overflow:hidden;
  -webkit-print-color-adjust:exact;print-color-adjust:exact;}}
.cover{{position:relative;width:210mm;height:297mm;overflow:hidden;color:#fff;
  background:radial-gradient(900px 520px at 84% -8%, {pal['highlight']} 0%,
    rgba(30,91,138,0) 62%),
    linear-gradient(158deg,{pal['bg1']} 0%,{pal['bg2']} 56%,{pal['bg3']} 100%);}}
.grid{{position:absolute;inset:0;opacity:.15;
  background-image:linear-gradient(#ffffff22 1px,transparent 1px),
    linear-gradient(90deg,#ffffff22 1px,transparent 1px);
  background-size:26mm 26mm;}}
.glow{{position:absolute;width:150mm;height:150mm;right:-40mm;top:-30mm;border-radius:50%;
  background:radial-gradient(circle,{pal['glow1']}66 0%,rgba(0,0,0,0) 68%);}}
.glow2{{position:absolute;width:120mm;height:120mm;left:-46mm;bottom:10mm;border-radius:50%;
  background:radial-gradient(circle,{pal['glow2']}40 0%,rgba(0,0,0,0) 70%);}}
.inner{{position:relative;height:100%;padding:38mm 24mm 24mm;
  display:flex;flex-direction:column;}}
.kicker{{font-family:{FONT_H};font-size:11pt;letter-spacing:6px;color:{pal['kicker']};}}
.kicker::after{{content:"";display:block;width:58px;height:3px;
  background:{pal['accent']};margin-top:12px;}}
h1{{font-family:{FONT_H};font-size:38pt;line-height:1.28;margin-top:30px;
  font-weight:800;letter-spacing:1px;}}
h1 .sub{{display:block;font-size:21pt;font-weight:400;color:{pal['sub']};
  margin-top:16px;letter-spacing:2px;}}
.divider{{width:130px;height:5px;
  background:linear-gradient(90deg,{pal['accent']},{pal['accent_soft']});
  margin:32px 0;border-radius:3px;}}
.desc{{font-family:{FONT_H};font-size:12pt;color:{pal['desc']};line-height:2.05;
  max-width:135mm;}}
.meta{{margin-top:auto;display:flex;gap:12px;flex-wrap:wrap;}}
.chip{{border:1px solid #ffffff55;border-radius:20px;padding:7px 17px;
  font-size:10pt;color:#E6EEF6;font-family:{FONT_H};}}
.band{{position:absolute;left:0;right:0;bottom:0;height:17mm;
  background:linear-gradient(90deg,{pal['band_a']},{pal['band_b']} 58%,{pal['band_a']});}}
.band span{{position:absolute;right:24mm;bottom:5.4mm;font-size:8.5pt;
  color:{pal['band_text']};letter-spacing:3px;font-family:{FONT_H};}}
.band .left{{left:24mm;right:auto;letter-spacing:4px;}}"""

    return f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<style>{css}</style></head><body>
<div class="cover"><div class="grid"></div><div class="glow"></div><div class="glow2"></div>
<div class="inner">
  <div class="kicker">{_html.escape(cov["kicker"])}</div>
  <h1>{_html.escape(cov["title"])}<span class="sub">{_html.escape(cov.get("subtitle", ""))}</span></h1>
  <div class="divider"></div>
  <div class="desc">{desc}</div>
  <div class="meta">{chips}</div>
</div>
<div class="band"><span class="left">{band_l}</span><span>{band_r}</span></div>
</div></body></html>"""


# ---------------------------------------------------------------- 目录回填
def read_toc_keys(body_path):
    """正文 HTML 里形如 <span class="pg" data-s="锚点文本"> 的条目。"""
    with open(body_path, encoding="utf-8") as f:
        s = f.read()
    return re.findall(r'<span class="pg" data-s="([^"]*)">', s)


def find_pages(body_pdf, keys, skip_marker=DEFAULT_NAV_MARKER):
    """在 PDF 文本层里找每个锚点首次出现的页码；命中导航标记页则跳过。"""
    doc = fitz.open(body_pdf)
    skip = set()
    if skip_marker:
        for i in range(len(doc)):
            if doc[i].search_for(skip_marker):
                skip.add(i + 1)
    out = []
    for k in keys:
        page_no = 0
        for i in range(len(doc)):
            if (i + 1) in skip:
                continue
            if doc[i].search_for(k):
                page_no = i + 1
                break
        out.append(page_no)
    doc.close()
    return out


# ---------------------------------------------------------------- 主流程
async def run(cfg):
    margin = cfg.get("margin") or DEFAULT_MARGIN
    body_html = os.path.abspath(cfg["body_html"])
    out_pdf = cfg["out_pdf"]
    os.makedirs(os.path.dirname(os.path.abspath(out_pdf)), exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()

        # ---- 封面（全出血，单独渲染）----
        cover_pdf = None
        if cfg.get("cover"):
            cov = os.path.join(os.path.dirname(os.path.abspath(cfg.get(
                "workdir", body_html))), cfg["name"] + "_cover.html")
            with open(cov, "w", encoding="utf-8") as f:
                f.write(cover_html(cfg))
            await page.goto("file://" + cov)
            await page.wait_for_timeout(400)
            cover_pdf = os.path.join(os.path.dirname(os.path.abspath(cov)),
                                     cfg["name"] + "_cover.pdf")
            await page.pdf(path=cover_pdf, width="210mm", height="297mm",
                           margin={"top": "0", "bottom": "0",
                                   "left": "0", "right": "0"},
                           print_background=True)

        # ---- 正文（带页眉页脚）----
        await page.goto("file://" + body_html)
        await page.wait_for_timeout(600)

        fl = _html.escape(cfg.get("footer_l", ""))
        hl = _html.escape(cfg.get("header_l", ""))
        hr = _html.escape(cfg.get("header_r", ""))
        pad = "padding:0 15mm;"

        # ⚠️ 页眉/页脚的 letter-spacing 必须 < 1px。
        # Chromium 在字距较大时会把一行文本按字符拆成多个 text run 落进 PDF 文本层，
        # 导致 `get_text()` 抽出「广 东 中 公 …」——视觉完全正常，但
        # 「署名出现次数 == 页数」这条合规巡检恒为 0，会误判成漏署名。
        header = (f'<div style="width:100%;font-family:{FONT_H};font-size:7.5pt;'
                  f'color:#9AA9B8;{pad}display:flex;justify-content:space-between;'
                  f'align-items:center;">'
                  f'<span style="letter-spacing:0.7px;">{hl}</span>'
                  f'<span style="letter-spacing:0.7px;">{hr}</span></div>')
        footer = (f'<div style="width:100%;font-family:{FONT_H};font-size:7.5pt;'
                  f'color:#9AA9B8;{pad}display:flex;justify-content:space-between;'
                  f'align-items:center;">'
                  f'<span style="letter-spacing:0.7px;">{fl}</span>'
                  f'<span>第 <span class="pageNumber"></span> 页 / '
                  f'共 <span class="totalPages"></span> 页</span></div>')

        async def render(pdf_path):
            await page.pdf(path=pdf_path, format="A4", print_background=True,
                           margin=margin, display_header_footer=True,
                           header_template=header, footer_template=footer)

        body_pdf = os.path.join(os.path.dirname(os.path.abspath(body_html)),
                                cfg["name"] + "_body.pdf")
        tmp = body_pdf.replace(".pdf", "_pass1.pdf")
        await render(tmp)

        # ---- 回填目录页码后二次渲染 ----
        keys = read_toc_keys(body_html)
        if keys:
            nums = find_pages(tmp, keys, cfg.get("nav_marker", DEFAULT_NAV_MARKER))
            await page.evaluate(
                "(nums)=>{const els=[...document.querySelectorAll('.toc .pg')];"
                "els.forEach((el,i)=>{if(nums[i])el.textContent=nums[i];});}", nums)
            bad = [k for k, n in zip(keys, nums) if not n]
            print(f"TOC: {len(keys)} 项", ("未解析 -> " + ", ".join(bad)) if bad else "全部命中")
            await render(body_pdf)

        # ---- 合并 ----
        out = fitz.open()
        if cover_pdf:
            out.insert_pdf(fitz.open(cover_pdf))
        out.insert_pdf(fitz.open(body_pdf))
        out.save(out_pdf, deflate=True, garbage=3)
        out.close()
        n = fitz.open(out_pdf).page_count
        print(f"[OK] {out_pdf}  共 {n} 页")
        await browser.close()
        return n


SCHEMA = """
docs.json 示例（数组，每项一本）：
[
  {
    "name": "d1",
    "palette": "gov",                    # gov|edu|med|fin|plain，默认 gov
    "palette_override": {},              # 可选，细改某个色值
    "cover": {                           # 省略或 null = 无封面
      "kicker": "2026 · 广州南沙",
      "title": "结构化面试理论讲义",
      "subtitle": "综合岗 · 编外人员集中招聘备考",
      "desc": ["搭建知识框架", "从“考什么”到“怎么答”"],
      "chips": ["第一分册 · 理论", "深度版"],
      "band_l": "广州南沙",
      "band_r": "页脚署名文字"           # 省略则用 footer_l
    },
    "header_l": "结构化面试理论讲义",     # 页眉左
    "header_r": "广州南沙 · 综合岗",      # 页眉右
    "footer_l": "署名文字",               # 页脚左（署名）
    "body_html": "/abs/path/d1_body.html",
    "out_pdf":  "/abs/path/out/01_xxx.pdf",
    "margin": {"top":"17mm","bottom":"16mm","left":"15mm","right":"15mm"},
                                          # 可选。⚠️ 只要 style.css 声明了非零
                                          # @page{margin:...}，本字段就被 CSS 顶掉、
                                          # 完全不生效（实测 17/18/30mm 产物几何一致）。
                                          # 改版心请改 style.css 的 @page，并把这里同步成
                                          # 同一个值仅作声明用。本字段只在 CSS 无 @page
                                          # margin 规则时才真正起作用。
    "nav_marker": "导航页"                # 可选，目录页白色标记，用于跳过误匹配
  }
]
"""


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print(SCHEMA)
        sys.exit(1)
    cfg_path = sys.argv[1]
    only = sys.argv[2] if len(sys.argv) > 2 else None
    with open(cfg_path, encoding="utf-8") as f:
        docs = json.load(f)
    for cfg in docs:
        if only and cfg["name"] != only:
            continue
        n = asyncio.run(run(cfg))
        if n is None:
            sys.exit(2)


if __name__ == "__main__":
    main()
