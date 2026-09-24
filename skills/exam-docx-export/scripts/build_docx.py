# -*- coding: utf-8 -*-
"""HTML → Word(.docx) 生成驱动（editor_sdk 通道）

流程：新建空白文档 → 页面/段落默认样式 → 封面 → 分页符 → 目录
      → 正文（分块插入，图表处插入高清 PNG 主图 + SVG 矢量双嵌）
      → 页眉署名 → 页脚页码 → 保存

⚠️ 本脚本是**一次实际交付的驱动实例**（南沙区情手册），可当模板照改：
   COVER 文案、图表键名、页眉文字、章节标题都是那一次的实例值，换项目要逐处改。
   渲染/写入逻辑本身是通用的，不要改。

路径全部可用环境变量覆盖，默认值落在临时目录与当前工作目录，不写死任何本机路径：
   EXAM_EDSDK    editor_sdk 的 edsdk.py 路径
   EXAM_WORK     中间产物目录（默认 <临时目录>/exam_docx_work/word）
   EXAM_CHARTS   图表 SVG 目录（默认 <EXAM_WORK>/../build/charts）
   EXAM_TARGET   成品 .docx 落盘路径（默认 <当前目录>/out.docx）

用法：
   python build_docx.py
"""
import importlib.util
import json
import os
import re
import sys
import tempfile

EDSDK = os.environ.get("EXAM_EDSDK") or (
    "/Applications/WorkBuddy AI.app/Contents/Resources/app.asar.unpacked/"
    "resources/plugins/workbuddy-builtin/skills/tencent-local-office-edit/edsdk.py")
WORK = os.environ.get("EXAM_WORK") or os.path.join(
    tempfile.gettempdir(), "exam_docx_work", "word")
CHARTS = os.environ.get("EXAM_CHARTS") or os.path.join(
    os.path.dirname(WORK), "build", "charts")
MD = os.path.join(WORK, "manual.md")
TARGET = os.environ.get("EXAM_TARGET") or os.path.join(os.getcwd(), "out.docx")

spec = importlib.util.spec_from_file_location("edsdk", EDSDK)
edsdk = importlib.util.module_from_spec(spec)
spec.loader.exec_module(edsdk)


class RpcError(Exception):
    pass


def _quiet_die(msg, rpc_code=None):
    raise RpcError(msg)


edsdk._die = _quiet_die


def call(tool, **args):
    res = edsdk._rpc("tools/call", {"name": tool, "arguments": args})
    texts = [c.get("text", "") for c in (res.get("content") or [])
             if isinstance(c, dict) and c.get("type") == "text"]
    raw = "\n".join(texts)
    try:
        return json.loads(raw)
    except Exception:
        return {"_text": raw}


def call_r(tool, tries=8, **args):
    """带重试的调用：新建文档后编辑器偶发「document is not open」。"""
    import time
    last = None
    for i in range(tries):
        try:
            return call(tool, **args)
        except RpcError as e:
            last = e
            time.sleep(1.5)
    raise RuntimeError("调用 %s 连续失败: %s" % (tool, last))


def fid_of(res):
    if isinstance(res, dict):
        for k in ("file_id", "id"):
            if res.get(k):
                return res[k]
    m = re.search(r"file_id=([^\s,]+)", res.get("_text", ""))
    if m:
        return m.group(1)
    raise RuntimeError("无法取得 file_id: %r" % res)


def pos_of(res):
    for k in ("position", "last_edit_index", "end_index", "next_index"):
        v = res.get(k) if isinstance(res, dict) else None
        if isinstance(v, int):
            return v
    raise RuntimeError("无法取得插入位置: %r" % res)


def log(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- 资源准备
def chart_size_pt(path):
    import pymupdf
    d = pymupdf.open(path)
    r = d[0].rect
    return r.width, r.height


def render_png(svg_path, out_png, scale=5):
    import pymupdf
    d = pymupdf.open(svg_path)
    pix = d[0].get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
    pix.save(out_png)
    return pix.width, pix.height


MAXW = 462.0  # 正文可用宽度 595.3 - 2*56.7 = 481.9pt，留余量

def prepare():
    os.makedirs(os.path.join(WORK, "chunks"), exist_ok=True)
    meta = {}
    for k in ("c1_targets", "c2_towns", "c3_structure", "c4_clusters"):
        svg = os.path.join(CHARTS, k + ".svg")
        png = os.path.join(WORK, k + ".png")
        w, h = chart_size_pt(svg)
        disp_w = min(w, MAXW)
        disp_h = h * disp_w / w
        pw, ph = render_png(svg, png)
        meta[k] = {
            "svg": svg, "png": png,
            "w_px": round(disp_w * 96 / 72), "h_px": round(disp_h * 96 / 72),
            "dpi": round(pw / (disp_w / 72)),
        }
        log("  chart %-13s %.1f×%.1fpt → 显示 %d×%dpx, PNG %d×%d (≈%d dpi)"
            % (k, w, h, meta[k]["w_px"], meta[k]["h_px"], pw, ph, meta[k]["dpi"]))
    return meta


def split_md():
    with open(MD, encoding="utf-8") as f:
        md = f.read()
    parts = re.split(r"^\[\[CHART:([A-Za-z0-9_]+)\]\]\s*$", md, flags=re.M)
    # parts = [text0, key1, text1, key2, text2, ...]
    chunks, keys = [], []
    chunks.append(parts[0])
    i = 1
    while i < len(parts):
        keys.append(parts[i])
        chunks.append(parts[i + 1])
        i += 2
    files = []
    for n, c in enumerate(chunks):
        p = os.path.join(WORK, "chunks", "c%d.md" % n)
        with open(p, "w", encoding="utf-8") as f:
            f.write(c.strip() + "\n")
        files.append(p)
    log("  markdown 分块 %d 段，图表 %d 处：%s" % (len(files), len(keys), keys))
    return files, keys


COVER = """广州市南沙区情手册（2026 版）

事业单位 · 社区专职 · 编外人员备考通用

区位与区划　经济与产业　镇街全景　时政热点　五年蓝图　招考落点

数据截至 2026 年 9 月　|　共 12 章 + 附录

广东中公教研

目　录
"""


# ---------------------------------------------------------------- 主流程
def main():
    log("[1] 准备图表与分块")
    meta = prepare()
    chunks, keys = split_md()

    cover_path = os.path.join(WORK, "chunks", "cover.md")
    with open(cover_path, "w", encoding="utf-8") as f:
        f.write(COVER)

    log("[2] 新建文档")
    fid = fid_of(call_r("create_doc"))
    log("    file_id =", fid)

    log("[3] 页面与默认段落样式")
    call_r("doc_set_document_style",
         file_id=fid,
         page_style={"page_width": 595.3, "page_height": 841.9,
                     "top_margin": 56.7, "bottom_margin": 56.7,
                     "left_margin": 56.7, "right_margin": 56.7})
    call_r("doc_set_document_style",
         file_id=fid,
         default_paragraph_style={"line_spacing": 1.5, "line_spacing_rule": 1,
                                  "spacing_after": 5, "spacing_before": 0})

    log("[4] 封面")
    r = call_r("doc_insert_markdown", file_id=fid, idx=0,
             markdown="file://" + cover_path)
    pos_cover = pos_of(r)
    log("    封面插入完成，pos =", pos_cover)

    # 封面主标题 / 副标题套用内置 Title / Subtitle 样式（非 Heading，不进目录）
    log("[4b] 封面标题样式")
    for text, lvl in (("广州市南沙区情手册（2026 版）", 11),
                      ("事业单位 · 社区专职 · 编外人员备考通用", 12)):
        f = call_r("doc_find", file_id=fid, text=text)
        locs = f.get("locations") or []
        if locs:
            first = min(locs, key=lambda x: x["begin"])
            call_r("doc_modify_paragraph", file_id=fid,
                   ranges=[{"begin": first["begin"], "end": first["end"]}],
                   heading_lvl=lvl, jc="center")
            log("    样式 %s ← %s" % (lvl, text[:18]))
        else:
            log("    ! 未找到 %s" % text[:18])

    log("[5] 分页符 + 目录")
    call_r("doc_insert_page_break", file_id=fid, idx=pos_cover)
    call_r("doc_insert_toc", file_id=fid, idx=pos_cover, max_level=2)

    last = call_r("doc_get_last_operable_pos", file_id=fid)
    pos = last.get("position")
    log("    正文起点 pos =", pos)

    log("[6] 正文分块写入")
    for n, cp in enumerate(chunks):
        r = call_r("doc_insert_markdown", file_id=fid, idx=pos,
                 markdown="file://" + cp)
        pos = pos_of(r)
        log("    chunk%d ok → pos %d" % (n, pos))
        if n < len(keys):
            k = keys[n]
            m = meta[k]
            ri = call_r("doc_insert_image", file_id=fid, idx=pos,
                      image_path=m["svg"], w=m["w_px"], h=m["h_px"])
            pos = pos_of(ri)
            rp = call_r("doc_insert_paragraph_with_text", file_id=fid,
                      idx=pos, text="")
            pos = pos_of(rp)
            log("    图 %s 插入 ok → pos %d" % (k, pos))

    log("[6b] 套用内置标题样式（导航窗格 / 目录可识别）")
    st = call_r("doc_resolve_document_structure", file_id=fid, mode="outline",
                limit=0, text_preview_length=0)
    groups = {1: [], 2: [], 3: []}
    for nd in (st.get("nodes") or []):
        hl = nd.get("heading_level")
        if nd.get("type") == "Heading" and hl in groups:
            groups[hl].append({"begin": nd["start_index"], "end": nd["end_index"]})
    for lvl, name in ((1, "HEADING_1"), (2, "HEADING_2"), (3, "HEADING_3")):
        if groups[lvl]:
            call_r("doc_apply_named_style", file_id=fid,
                   ranges=groups[lvl], named_style_type=name)
            log("    %s × %d" % (name, len(groups[lvl])))

    log("[7] 页眉 / 页脚")
    call_r("doc_insert_header", file_id=fid,
         text="广州市南沙区情手册（2026 版）　|　广东中公教研")
    call_r("doc_set_page_number", file_id=fid, position="right", format="decimal")

    log("[8] 保存")
    res = call_r("save_file", file_id=fid, file_path=TARGET)
    log("   ", res.get("_text") or res)

    with open(os.path.join(WORK, "build_info.json"), "w", encoding="utf-8") as f:
        json.dump({"file_id": fid, "target": TARGET,
                   "charts": meta, "keys": keys}, f,
                  ensure_ascii=False, indent=2)
    log("完成 →", TARGET)


if __name__ == "__main__":
    main()
