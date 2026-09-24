#!/usr/bin/env python3
"""批量抓取招聘公告页 → 提取岗位表附件链接 → 下载到 ./att/

用法：
    python fetch_attachments.py urls.txt [att_dir]
      urls.txt   每行一个公告页 URL（可带 #备注）
      att_dir    下载目录，默认 ./att

关键点（踩过的坑）：
  * 附件链接多为相对路径 → urljoin
  * URL 含中文/空格 → 必须 quote() 逐段编码，否则 400/404
  * 政务站点多数要带 Referer 才给下载 → headers 里带上公告页 URL
  * 一个公告页可能有多个附件 → 优先取文件名含「岗位」的
"""
import sys, os, re, json
import urllib.request, urllib.parse
from urllib.parse import urljoin, unquote, quote

UA = ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/124.0 Safari/537.36')


def fetch(url, referer=None, binary=False):
    hdr = {'User-Agent': UA, 'Accept': '*/*'}
    if referer:
        hdr['Referer'] = referer
    req = urllib.request.Request(url, headers=hdr)
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def safe_url(u):
    """把含中文/空格的 URL 逐段编码（保留 :// 与已编码部分）"""
    parts = urllib.parse.urlsplit(u)
    path = quote(unquote(parts.path), safe='/%')
    query = quote(unquote(parts.query), safe='=&%')
    return urllib.parse.urlunsplit(
        (parts.scheme, parts.netloc, path, query, parts.fragment))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    urls = [l.strip() for l in open(sys.argv[1]) if l.strip() and not l.startswith('#')]
    att_dir = sys.argv[2] if len(sys.argv) > 2 else 'att'
    os.makedirs(att_dir, exist_ok=True)

    manifest = {}
    for idx, url in enumerate(urls, 1):
        url = url.split('#')[0].strip()
        try:
            html = fetch(url).decode('utf-8', 'ignore')
        except Exception as e:
            print(f'[{idx:02d}] 页面失败 {url} :: {e}')
            continue
        links = re.findall(r'href="([^"]+\.(?:xlsx|xls))"', html, re.I)
        links = [urljoin(url, l) for l in links]
        picked = [l for l in links if '岗位' in unquote(l)] or links
        if not picked:
            print(f'[{idx:02d}] 未找到岗位表附件 {url}')
            continue
        got = []
        for j, link in enumerate(picked, 1):
            name = unquote(os.path.basename(urllib.parse.urlsplit(link).path))
            if not name.lower().endswith(('.xlsx', '.xls')):
                name = f'att_{idx}_{j}.xlsx'
            out = os.path.join(att_dir, name)
            if os.path.exists(out) and os.path.getsize(out) > 2000:
                got.append(out); continue
            try:
                data = fetch(safe_url(link), referer=url, binary=True)
                if len(data) < 1000:
                    print(f'[{idx:02d}] 附件过小，疑似被拦：{name}')
                    continue
                with open(out, 'wb') as f:
                    f.write(data)
                got.append(out)
            except Exception as e:
                print(f'[{idx:02d}] 附件失败 {name} :: {e}')
        manifest[url] = got
        print(f'[{idx:02d}] {len(got)} 个附件 <- {url}')

    with open(os.path.join(att_dir, '_manifest.json'), 'w') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    print(f'\n共 {sum(len(v) for v in manifest.values())} 个附件 -> {att_dir}/')


if __name__ == '__main__':
    main()
