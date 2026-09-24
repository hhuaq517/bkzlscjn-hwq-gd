#!/usr/bin/env python3
"""批量解析招聘岗位表附件（.xlsx / .xls），输出 summary.json。

用法：
    # 1) 把所有岗位表放进 ./att/  （文件名任意，作为批次 key）
    # 2) python parse_positions.py [att_dir] [out_json]

输出 summary.json 结构：
    { "<批次key>": { "units": [ {sheet, unit, posts, people, cats, edus, locs, exam} ],
                     "total_posts": int, "total_people": int,
                     "exam_forms": [str] } }

设计要点（都是踩过的坑，别改）：
  * .xls 走 xlrd、.xlsx 走 openpyxl，统一成 [(sheet_name, [row_tuple])]
  * 表头行位置不固定 → 动态找同时含「岗位编码」和「招聘人数」的行
  * 人数列可能是 "控制数 5 人" 或合并单元格空白 → 用正则抓第一个整数并记录 ANOM
  * 「考试形式」写在表头上方的合并单元格里 → 必须 join 整行再判断
解析完务必与公告正文人数交叉核对。
"""
import sys, os, re, json, glob

try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    import xlrd
except ImportError:
    xlrd = None

ATT_DIR = sys.argv[1] if len(sys.argv) > 1 else 'att'
OUT_JSON = sys.argv[2] if len(sys.argv) > 2 else 'summary.json'


def norm(s):
    return re.sub(r'\s+', '', str(s)) if s is not None else ''


def read_sheets(path):
    """统一读成 [(sheet_name, [row_tuple, ...]), ...]"""
    out = []
    if path.lower().endswith('.xls'):
        if xlrd is None:
            raise RuntimeError('需要 xlrd 才能读 .xls')
        book = xlrd.open_workbook(path)
        for sh in book.sheets():
            out.append((sh.name, [tuple(sh.row_values(i)) for i in range(sh.nrows)]))
    else:
        if openpyxl is None:
            raise RuntimeError('需要 openpyxl 才能读 .xlsx')
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            out.append((ws.title, [tuple(r) for r in ws.iter_rows(values_only=True)]))
    return out


def find_col(hdr, *keys):
    for i, h in enumerate(hdr):
        for k in keys:
            if k in h:
                return i
    return None


ANOM = []


def to_int(v, ctx=''):
    """人数单元格 → int。容忍 '控制数 5 人' / '5人' / 合并单元格空白。"""
    s = norm(v).replace(',', '')
    if not s:
        return 0
    m = re.search(r'(\d+)', s)
    if m:
        if not re.fullmatch(r'\d+', s):
            ANOM.append((ctx, s))
        return int(m.group(1))
    ANOM.append((ctx, s))
    return 0


def main():
    result = {}
    files = sorted(f for f in glob.glob(os.path.join(ATT_DIR, '*'))
                   if not os.path.basename(f).startswith('.'))
    if not files:
        print(f'!! {ATT_DIR}/ 下没有附件', file=sys.stderr)
        return
    for path in files:
        key = os.path.basename(path).rsplit('.', 1)[0]
        units, exam_forms = [], []
        for name, rows in read_sheets(path):
            sheet_exam = []
            hdr_idx = None
            for i, r in enumerate(rows):
                j = ''.join(norm(x) for x in r)
                if '岗位编码' in j and '招聘人数' in j:
                    hdr_idx = i
                    break
            if hdr_idx is None:
                continue
            hdr = [norm(x) for x in rows[hdr_idx]]
            # 表头上方：抓「考试形式 / 开考比例」
            for r in rows[:hdr_idx]:
                cells = [str(x).strip() for x in r if x is not None and str(x).strip()]
                joined = ' '.join(cells)
                if '考试形式' in joined or '开考比例' in joined:
                    exam_forms.append(joined)
                    sheet_exam.append(joined)

            ci_unit = find_col(hdr, '招聘单位')
            ci_code = find_col(hdr, '岗位编码')
            ci_num = find_col(hdr, '招聘人数')
            ci_cat = find_col(hdr, '招聘类别', '岗位类别')
            ci_loc = find_col(hdr, '工作地点')
            ci_edu = find_col(hdr, '学历')
            n_post = n_people = 0
            cats, edus, locs = {}, {}, {}
            uname = None
            for r in rows[hdr_idx + 1:]:
                if ci_code is None or len(r) <= ci_code or not norm(r[ci_code]):
                    continue
                n_post += 1
                num = to_int(r[ci_num], f"{key}/{name}") if (ci_num is not None and len(r) > ci_num) else 0
                n_people += num
                if ci_unit is not None and len(r) > ci_unit and not uname:
                    uname = str(r[ci_unit]).strip()
                if ci_cat is not None and len(r) > ci_cat:
                    c = norm(r[ci_cat]); cats[c] = cats.get(c, 0) + num
                if ci_edu is not None and len(r) > ci_edu:
                    e = norm(r[ci_edu]); edus[e] = edus.get(e, 0) + num
                if ci_loc is not None and len(r) > ci_loc:
                    l = norm(r[ci_loc]); locs[l] = locs.get(l, 0) + num
            if n_post:
                units.append({'sheet': name, 'unit': uname or name, 'posts': n_post,
                              'people': n_people, 'cats': cats, 'edus': edus,
                              'locs': locs, 'exam': ' '.join(sheet_exam)})
        tp = sum(u['posts'] for u in units)
        tn = sum(u['people'] for u in units)
        result[key] = {'units': units, 'total_posts': tp,
                       'total_people': tn, 'exam_forms': exam_forms}
        print(f"{key}: {len(units):2d} units | {tp:4d} posts | {tn:4d} people")

    with open(OUT_JSON, 'w') as f:
        json.dump(result, f, ensure_ascii=False, indent=1)

    print('\n--- anomalies（非纯数字的人数单元格，需人工确认） ---')
    seen = set()
    for a in ANOM:
        if a[1] not in seen:
            seen.add(a[1])
            print('  ', a[0], '->', a[1])
    print(f'\nsaved {OUT_JSON}')


if __name__ == '__main__':
    main()
