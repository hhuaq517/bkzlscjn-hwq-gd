#!/usr/bin/env python3
"""HTML 定页溢出检测（JS 注入法）——比像素法准得多，逐页给溢出像素值。

原理：注入脚本对每页比较「内容实际底边」与「内容框允许底线」，把结果写进
document.title，再用 Edge --dump-dom 把 DOM（含 title）取回来解析。

用法：
    python overflow_check.py build/index.html [--content .content] [--edge PATH] [--pad-bottom 12]

输出：
    OVFCHK::none                  → 零溢出，通过
    OVFCHK::11:6|14:28|24:3116    → 第 11 页溢 6px、第 14 页溢 28px、第 24 页溢 3116px

⚠️ 绝对不要用 `c.scrollHeight - c.clientHeight`：
   当 .content 没有固定高度（height:auto，靠内容撑开——本系列默认写法）时，
   scrollHeight 恒等于 clientHeight，结果**恒为 0**，任何溢出都测不出来，
   一律返回 OVFCHK::none 假通过。东莞教师册第一轮就因此漏掉 P12 的 7px 溢出。
   必须用 getBoundingClientRect 对比真实几何底边（本脚本做法）。

注意：
  * --pad-bottom 必须与 CSS 里 .page 的下内边距一致（本系列默认 12mm），否则底线算错
  * 必须用 --virtual-time-budget 给 JS 留执行时间，否则 title 可能还没写就被 dump
  * Edge 无头常不自己退出 → 必须包 sleep + kill
  * 溢出阈值设 2px（1px 是亚像素舍入噪声，不算问题）
"""
import sys, os, re, subprocess, tempfile, shutil

EDGE_CANDIDATES = [
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
]

INJECT = """
<script>
(function () {
  function run() {
    var MM = 96 / 25.4;                 // 1mm = 96/25.4 px
    var PAD_BOTTOM = __PAD__;           // 与 CSS .page 的下内边距一致（mm）
    var bad = [];
    var pages = document.querySelectorAll('.page');
    pages.forEach(function (p, i) {
      // 封面等"全出血装饰页"必须跳过：装饰圆/色块故意超出视口（overflow:hidden 已裁），
      // 用 `querySelector(sel) || p` 的兜底写法会把它们算成溢出（实测误报 1:272）
      if (p.classList.contains('cover')) return;
      var c = p.querySelector('__SEL__');
      if (!c) return;                       // 没有内容框的页不参与判定
      var limit = p.getBoundingClientRect().bottom - PAD_BOTTOM * MM;  // 内容允许到达的底线
      var maxB = c.getBoundingClientRect().bottom;
      // 连子元素一起取 max：.content 若被 CSS 定了高，内部溢出仍能被抓到
      Array.prototype.forEach.call(c.children, function (el) {
        var b = el.getBoundingClientRect().bottom;
        if (b > maxB) maxB = b;
      });
      var ovf = Math.round(maxB - limit);
      if (ovf > 2) bad.push((i + 1) + ':' + ovf);
    });
    document.title = 'OVFCHK::' + (bad.length ? bad.join('|') : 'none');
  }
  if (document.readyState === 'complete') { run(); }
  else { window.addEventListener('load', run); }
})();
</script>
"""


def find_edge(explicit=None):
    if explicit and os.path.exists(explicit):
        return explicit
    for c in EDGE_CANDIDATES:
        if os.path.exists(c):
            return c
    raise SystemExit('找不到 Edge / Chrome，请用 --edge 指定路径')


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return
    src = args[0]
    sel = '.content'
    pad = '12'          # 与 CSS .page 下内边距一致（mm）
    edge = None
    if '--content' in args:
        sel = args[args.index('--content') + 1]
    if '--pad-bottom' in args:
        pad = args[args.index('--pad-bottom') + 1]
    if '--edge' in args:
        edge = args[args.index('--edge') + 1]
    edge = find_edge(edge)

    html = open(src, encoding='utf-8').read()
    inject = INJECT.replace('__SEL__', sel).replace('__PAD__', pad)
    if '</body>' in html:
        out_html = html.replace('</body>', inject + '</body>')
    else:
        out_html = html + inject
    tmp = os.path.join(os.path.dirname(os.path.abspath(src)), 'check.html')
    with open(tmp, 'w', encoding='utf-8') as f:
        f.write(out_html)

    url = 'file://' + urllib_quote(tmp)
    cmd = (f'"{edge}" --headless=new --disable-gpu --no-pdf-header-footer '
           f'--virtual-time-budget=4000 --dump-dom "{url}"')
    # Edge 无头不自行退出 → 包 sleep + kill
    wrapped = f'( {cmd} & pid=$!; sleep 12; kill $pid 2>/dev/null )'
    try:
        dom = subprocess.run(['/bin/zsh', '-c', wrapped],
                             capture_output=True, text=True, timeout=90).stdout
    except subprocess.TimeoutExpired:
        dom = ''

    m = re.search(r'OVFCHK::([^<"\s]*)', dom)
    if not m:
        print('!! 未取到 OVFCHK 结果。检查：')
        print('   - Edge 路径是否正确')
        print('   - .page / .content 选择器是否与 HTML 一致')
        print('   - 是否被 kill 得太早（把 sleep 12 调大）')
        return 1

    res = m.group(1)
    if res == 'none':
        print('OVFCHK::none  ✅ 零溢出')
        return 0
    print('OVFCHK::' + res)
    print('\n溢出页明细（页号:溢出像素）：')
    for item in res.split('|'):
        pg, px = item.split(':')
        print(f'  第 {pg:>2} 页  溢 {px:>5} px  ≈ {int(px) * 25.4 / 96:.1f} mm')
    print('\n修法优先级：内容搬家 > 删内容 > 拉伸留白')
    return 1


def urllib_quote(p):
    from urllib.parse import quote
    return quote(os.path.abspath(p), safe='/')


if __name__ == '__main__':
    sys.exit(main())
