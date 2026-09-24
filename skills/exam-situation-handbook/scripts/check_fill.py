#!/usr/bin/env python3
"""PDF 填充率巡检——像素法快速定位"孤行页"（内容极少、大片留白）。

用法：
    python check_fill.py out.pdf [--footer 广东中公] [--dpi 72]

输出每页填充率百分比。判定标准：
    100%+  → 溢出（配合 overflow_check.py 一起看）
    55%—99% → 正常
    < 55%   → 偏空，考虑补内容或把别处内容搬来
    < 10%   → 孤行页，必须修

原理：按行扫描像素，找每页最后一个"有墨"的行，换算成 mm 与版心高度比较。
比逐页目检快得多，适合改完排版后秒级复查。
"""
import sys, os
import numpy as np

PAGE_H_MM = 297.0
TOP_MM = 14.5          # 版心上边距
BOT_MM = 277.5         # 版心下边距（含页脚区）—— 与 style.css 的 padding 对应


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    pdf = sys.argv[1]
    dpi = 72
    if '--dpi' in sys.argv:
        dpi = int(sys.argv[sys.argv.index('--dpi') + 1])

    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(pdf)
    scale = dpi / 72.0
    mm = 25.4 / dpi            # 1 像素 = 多少 mm

    print(f'{pdf}')
    print(f'共 {len(doc)} 页   dpi={dpi}\n')
    lows = []
    for i in range(len(doc)):
        bmp = doc[i].render(scale=scale, rotation=0)
        arr = np.asarray(bmp.to_pil().convert('L'))
        # 每行是否有墨（阈值 245，容忍极浅的表格线）
        rows = np.where((arr < 245).sum(axis=1) > 0)[0]
        if len(rows) == 0:
            print(f'  p{i+1:02d}   空页  ⚠')
            lows.append(i + 1)
            continue
        body = rows[rows * mm < BOT_MM]
        bottom = (body.max() * mm) if len(body) else 0.0
        fill = (bottom - TOP_MM) / (BOT_MM - TOP_MM) * 100
        flag = ''
        if fill < 10:
            flag = '  ⚠ 孤行页'
            lows.append(i + 1)
        elif fill < 55:
            flag = '  · 偏空'
            lows.append(i + 1)
        elif fill > 100:
            flag = '  ⚠ 溢出'
        print(f'  p{i+1:02d}  {fill:6.1f}%   底边 {bottom:6.1f}mm{flag}')

    print()
    if lows:
        print(f'需关注页：{lows}')
    else:
        print('✅ 全部页面填充率正常（>55%），无孤行页')


if __name__ == '__main__':
    main()
