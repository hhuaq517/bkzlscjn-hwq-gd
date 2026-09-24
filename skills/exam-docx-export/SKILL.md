---
name: exam-docx-export
description: 把已排好版的 HTML/PDF 资料（考情手册、讲义、题本）导出为**可编辑的 Word(.docx)**。走本地 editor_sdk（tencent-local-office-edit）通道：HTML→Markdown→分块写入→套内置标题样式→页眉页码→保存→Python 后处理重打包 zip（补高清位图回退、拆图注段落、强制目录刷新）。触发词：生成word版本、导出docx、PDF转Word、做成可编辑Word、纯文字Word、保留文字即可。
agent_created: true
---

# 资料 → Word(.docx) 导出管线

**本技能只管"把已有内容导出成一份能用的 Word"，不管内容怎么写。** 内容体系由 `exam-situation-handbook` / `exam-prep-pdf` 等负责；那一步的产物是 HTML（或 PDF），本技能从 HTML 出发。

## 〇、先判通道

| 情形 | 走哪 |
|---|---|
| **新建**一份 Word（无源材料，纯创作） | `tencent-docx`（有 doc-writer/formatter 子 agent） |
| **已有源材料**（HTML / PDF / 已排版的资料）要转成 Word | **本技能** → `tencent-local-office-edit`（editor_sdk） |
| 只要改一个已存在的 .docx 的几个字 | 直接 `tencent-local-office-edit`，不必走本管线 |

用户的常见说法："生成一份纯 word 文档的版本，保留文字即可，矢量图形非要留就保持高清" → **本技能**。

## 一、总流程

```
HTML 源
  │ ① lxml 解析 → 干净 Markdown（图表处留 [[CHART:xxx]] 占位）
  ▼
manual.md
  │ ② create_doc → 页面/默认段落样式 → 封面 → 分页符 → 目录域
  │    ③ 按 [[CHART:]] 分块，逐块 doc_insert_markdown(markdown="file://…")
  │       图表处 doc_insert_image 插 SVG
  │    ④ doc_resolve_document_structure(mode=outline, limit=0) → 套 Heading1/2/3
  │    ⑤ 页眉（署名）+ 页脚（页码）→ save_file
  ▼
原始 .docx（SVG 无位图回退 / 图注与图同段 / 目录域未刷新）
  │ ⑥ Python 后处理重打包 zip
  ▼
成品 .docx
```

**为什么必须有第 ⑥ 步**：editor_sdk 的 3 个已知缺陷只能靠改 OOXML 兜底（见第三节）。

## 二、各步要点

### ① HTML → Markdown
- 用 `lxml.html.document_fromstring`（`bs4` 不一定装了，`lxml` 通常有）
- 先把 `<!--CHART:xxx-->` 注释替换成 `<div class="chartph" data-k="xxx">`，否则注释在遍历时拿不到
- **`section` / `div` 容器都要递归**——只处理 `div` 会漏掉 `<section class="chapter">` 整章
- 类映射建议：
  | HTML | Markdown |
  |---|---|
  | `.chapter-head h2` | `# 第N章　标题`（`<div class="num">01</div>` → 中文序号） |
  | `h3` / `h4` | `## x.x` / `### 标题` |
  | `ul.clean li` | `- ` |
  | `.steps .st` | `1. `（若首条以「N 月」开头 → 改 `- `，避免"7. 9 月"混淆） |
  | `.tl`（时间轴容器） | `1. **<日期>**　<正文>`；**日期在 `.tl-d`、状态在 `span.m`、正文在 `.tl-b`，三层都是叶子 div，容器本身没有直接文本** —— 少写这一支，整条时间轴在 Word 里全丢 |
  | `.cards`（数据卡） | `- **<k>**：<v>` |
  | `table.tb` | Markdown 表格（单元格内 `\|` 换成 `／`） |
  | `.tip/.warn/.key` | `> **标签：** 正文` |
  | `.figure` | `[[CHART:key]]` + `*图 x-y　图注*` |
  | `.toc` | **跳过**（Word 用真目录域） |
- **空白折叠的坑（必看）**：Python str 正则里 `\s` **包含全角空格 U+3000**。`re.sub(r"\s+", " ", s)` 会把 `**A**　**B**` 变成 `**A** **B**`，随后被粗体清理规则 `\*\*\s*\*\*` 吞掉 → 两段粗体合并成一段。
  - 折叠用 `re.sub(r"[^\S\u3000]+", " ", s)`（排除 U+3000）
  - 粗体清理只用 `re.sub(r"\*\*[ \t\r\n]*\*\*", "", s)`
  - 需要"把全角空格附近的多余空白收掉"时用 `re.sub(r"[ \t]*　[ \t]*", "　", s)`，**不要**用 `[　\s]+ → 　`（会把 `9.9 平方千米` 变成 `9.9　平方千米`）
- 隐藏锚点要丢：`<span style="color:#fff;font-size:6pt">锚一章</span>` —— 判颜色要兼容 `#fff` 与 `#FFFFFF`（用 `re.search(r"color:\s*#(?:fff|ffffff)\b", style, re.I)`）

### ② editor_sdk 基础操作
```python
import importlib.util, json
spec = importlib.util.spec_from_file_location("edsdk", EDSDK_PY)
edsdk = importlib.util.module_from_spec(spec); spec.loader.exec_module(edsdk)

class RpcError(Exception): pass
edsdk._die = lambda msg, rpc_code=None: (_ for _ in ()).throw(RpcError(msg))

def call(tool, **args):                       # 直调 _rpc，比 CLI 好解析
    r = edsdk._rpc("tools/call", {"name": tool, "arguments": args})
    raw = "\n".join(c.get("text","") for c in (r.get("content") or []) if c.get("type")=="text")
    try:    return json.loads(raw)
    except: return {"_text": raw}
```
- `EDSDK_PY = "/Applications/WorkBuddy AI.app/Contents/Resources/app.asar.unpacked/resources/plugins/workbuddy-builtin/skills/tencent-local-office-edit/edsdk.py"`
- **必须加重试**：`create_doc` 之后立刻编辑偶发 `document is not open`，重试 8 次 ×1.5s
- **`edsdk.py schema <工具名>` 对 `doc_insert_toc` / `doc_apply_named_style` / `doc_set_document_style` 会报 `'list' object has no attribute 'get'`**（这几个工具的 inputSchema 是数组）→ schema 查不了 ≠ 不能调，直接 `--json` 调用即可
- `create_doc` 的返回是**人话文本**（`Created blank doc ... file_id=new_doc_xxx`），要正则抠 `file_id`；其余工具返回 JSON

### ③ 页面与内容写入
- 页面/默认段落样式**拆两次调用**：
  ```python
  call("doc_set_document_style", file_id=fid,
       page_style={"page_width":595.3,"page_height":841.9,
                   "top_margin":56.7,"bottom_margin":56.7,
                   "left_margin":56.7,"right_margin":56.7})     # A4 + 2cm
  call("doc_set_document_style", file_id=fid,
       default_paragraph_style={"line_spacing":1.5,"line_spacing_rule":1,
                                "spacing_after":5,"spacing_before":0})
  ```
  （**同时传两个字段会报 `document is not open`**）
- 长 Markdown 用 `markdown="file:///绝对路径"`，不要把整段塞进参数
- 分块写入：按 `[[CHART:]]` 切分，每块 `doc_insert_markdown` → 取返回的 `position`/`last_edit_index` 当下一次的 `idx`
  - **相邻两个图表会产生「空分块」**（两块 `[[CHART:a]]` / `[[CHART:b]]` 之间只有空行）→ 直接调 `doc_insert_markdown` 会报 `-32603 … InsertMarkdown: no content to insert` 并中断整轮。
  - 修法：写之前 `body = open(cp).read().strip()`，**空块只记日志、跳过写入**，但仍要在该位置插下一张图（配对型图表——上/下两张——正是这种排布）。
- 图片：`doc_insert_image(idx, image_path=<svg 绝对路径>, w=<px@96dpi>, h=<px@96dpi>)`
  - 尺寸 = pt × 96/72；正文可用宽度 = 595.3 − 2×56.7 = **481.9pt**，图表显示宽度留余量取 **≤462pt**
  - SVG 原始尺寸用 `pymupdf.open(svg)[0].rect` 读
- 封面：先写普通段落，再用 `doc_find`（返回的是 **`locations`**，不是 matches/results）拿 `begin/end`，配 `doc_modify_paragraph(ranges=[…], heading_lvl=11|12, jc="center")` 套 Title/Subtitle —— **Title/Subtitle 不是 Heading，不会进目录**

### ④ 套内置标题样式
```python
st = call("doc_resolve_document_structure", file_id=fid,
          mode="outline", limit=0, text_preview_length=0)   # ← limit=0 必须，否则截断
groups = {1:[],2:[],3:[]}
for nd in st["nodes"]:
    if nd.get("type")=="Heading" and nd.get("heading_level") in groups:
        groups[nd["heading_level"]].append({"begin":nd["start_index"],"end":nd["end_index"]})
for lvl,name in ((1,"HEADING_1"),(2,"HEADING_2"),(3,"HEADING_3")):
    if groups[lvl]:
        call("doc_apply_named_style", file_id=fid, ranges=groups[lvl], named_style_type=name)
```
- **不传 `limit=0` 会只返回前 ~150 个节点**（实测 14 个 H1 只回来 6 个），导致部分标题是样式、部分是直接格式 → 版式不一致
- editor_sdk 生成的 styleId 是随机的（`ugh5g9`…），但 `<w:name w:val="heading 1"/>` 是标准名，Word 认；字号 H1 18pt / H2 16pt / H3 14pt

### ⑤ 目录 / 页眉 / 页脚
- 顺序：封面 → `doc_insert_page_break(pos_cover)` → `doc_insert_toc(idx=pos_cover, max_level=2)` → `doc_get_last_operable_pos` 取正文起点
- **`doc_insert_footer`（署名）与 `doc_set_page_number`（页码）互斥** → 采用 **页眉放"书名｜署名" + 页脚放页码**
- 目录域占位文案是"未找到目录项。"，**编辑器与转换服务都不会求值** → 靠第 ⑥ 步的 `updateFields`
  - 两条兜底路线，**优先 A**：
    - **A（更干净）改多段落域**：把 `doc_insert_toc` 落的那个单段域改写成 Word 自己产出的结构 ——
      `begin` 段 ＋ **静态章名段 × N** ＋ `end` 段（三段以上都在同一域内）。域未求值时（腾讯文档预览、Quick Look、Finder 预览）直接看到一份**干净的章名目录**，没有任何提示语或报错文案；在 Word/WPS 里打开时又被真目录替换。
      顺手把 `\o "1-2"` 收成 **`\o "1-1"`**：与 PDF 版目录层级对齐（PDF 目录只列章），否则 Word 会多出 50+ 条节标题、目录膨胀到 2—3 页。
      章名清单从 `document.xml` 里按 `w:name="heading 1"` 反查 styleId 后抽取（**别去调 `doc_get_outline`，见第三节**）。
      静态段用 `w:after="40" w:line="240"` 的普通段落即可，不要加 tab leader（没有页码，点线会很怪）。
    - **B（旧法）** 把占位串替换成「**打开文档后按 F9 或右键「更新域」生成目录**」（`scripts/polish_docx.py` 已实现）

### ⑥ 后处理（`scripts/fix_docx.py`，通用）
对保存好的 .docx 做五件事（`python fix_docx.py <目标.docx>`）：
1. **补高清位图回退**：`doc_insert_image` 传 SVG 时只写 `<asvg:svgBlip>`，`<a:blip>` 没有 `r:embed` → Word 以外的查看器（Quick Look / Google Docs / 预览器）显示空图。
   为每张 SVG 用 `pymupdf` 渲染 `Matrix(5,5)`（≈360–410 dpi）PNG，写入 `word/media/`，加 relationship 与 `Default Extension="png"`，再把 `<a:blip>` 改成 `<a:blip r:embed="rIdPng">` —— **保留 svgBlip，实现 PNG+SVG 双嵌**，既高清又到处能看
2. **拆图注段落**：`doc_insert_image` 后补的 `doc_insert_paragraph_with_text(text="")` 不生效，图注会和图挤在同一段（贴着图右侧折行）。
   - ❌ **不要**用限定首字的正则：`r'(</w:drawing></w:r>)(<w:r>(?:(?!</w:r>).)*?<w:t>图[^<]{0,12})'` —— 图注未必以「图」开头（南网册是「数据来源：…」「上：…」「下：…」），**静默命中 0 处，一点报错都没有**
   - ✅ 改为扫描每个 `</w:drawing></w:r>`，只要**同一段内**（下一个 `</w:p>` 之前）还有非空 `<w:t>`，就在图后断段：
     `doc[:m.end()] + "</w:p><w:p>" + pPr + doc[m.end():]`
   - **必须把原段落的 `<w:pPr>…</w:pPr>` 复制给新段**，否则紧跟其后的 Heading 会掉样式（本次 `11.1 考试形式与结构` 就与 `written_struct` 图同段）
   - **图后紧跟 Heading 时，要把图段的 `<w:pPr>` 整个换成干净版**（`<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`，不留 pStyle）。只摘直设 `<w:keepNext>` 不够——图段带着 `heading 2` 样式，**样式定义里还有一份 `keepNext`**，「图 → 标题 → 后面的大表格」三级串联，图照样被顶到下一页，上一页留下半页空白（实测页填充率 ~54%）。判据：新段的 `<w:pStyle>` 经 `styles.xml` 映射后的 `w:name` 以 `heading` 开头
   - 图段顺带把行距改成 `line=240 lineRule=auto`（单倍）：正文默认 `line=330 auto`，部分渲染器会把内联图行高也乘 1.65，259pt 的图变成 427pt，本来放得下也被顶走
   - 同一张图**无图注**时图后直接是 `</w:p>`（下一张图 / 表格），不要误拆
3. **封面/目录之间补分页符**：调用顺序虽是"封面 → `doc_insert_page_break` → `doc_insert_toc`"，**实际落成的却是"封面 → 目录标题 → TOC 域 → 分页符"**——分页符跑到目录后面去了，封面和目录挤在同一页。
   - 当场看不出来（TOC 域是空的，占位只有一行），**等用户在 Word 里 F9 展开目录（14 章 + 38 节 ≈ 52 行）版面就被顶乱**
   - 修法：找目录标题段（默认 `目\u3000录`）→ 取其前一段的 `</w:p>` → 在此处插 `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`
   - 幂等判据：`doc.count('<w:br w:type="page"/>') >= 2` 就跳过
4. **删正文末尾空段落**：编辑器在落款后常留 1–2 个空段（含**自闭合 `<w:p .../>`**），会顶出**一整页空白页**。
   - 段落枚举正则必须**同时覆盖自闭合**：`<w:p\b[^>]*/>|<w:p\b(?:(?!</w:p>).)*?</w:p>`
   - 位置是 `</w:body>` 之前、**`<w:sectPr>` 之前**（sectPr 在最后一段之后，不在 body 最尾）
   - **别用 `<w:p\b[^>]*?(?:/>|>.*?</w:p>)\s*$` 配 `re.search`**：`re.S` 下 `.*?` 会跨段匹配到最左起点（上一段有内容）→ 判定为"非空"直接 break，结果一个也删不掉。要枚举全部段落再从尾部剥
5. **表格整形（三合一，最容易漏）**：`doc_insert_markdown` 落进来的表格**全部列等宽**，且 `tblCellMar` 只有 15 twips（0.75pt，文字贴着框线），表宽也只有正文宽的 86%。后果：一张"镇街｜行政村｜社区｜所辖村居名称"的表里，最后一列被压成**一列一个字**，行高虚高、整表难看。
   - **列宽按内容自适应**：每列需求 `d_j = max(该列最长单元格视觉宽度, 表头宽度×1.25)`；`w_j ∝ d_j^0.62`（阻尼指数，0.5 更平均 / 1.0 纯比例）；再对 `w_j < 9.5%` 的列抬底并重新归一
   - **表头必须单行放得下**的硬下限：`表头视觉宽度 × 240 twips + 2×108`（全角字宽 ≈ 字号 12pt = 240 twips）。缺这道，表头「镇（街道）」「社区居委会」会被折成两行
   - **拉到正文满宽**：`textw = pgSz.w − pgMar.left − pgMar.right`（A4 + 2cm 边距 = **9638 twips / 481.9pt**），写进 `<w:tblW>` 与各列
   - **`<w:tblLayout w:type="fixed"/>`**：不设 fixed 的话 Word 打开时会按 autofit 重排，辛苦算的列宽白费
   - **`tblCellMar` 左右改 108 twips（5.4pt）**、上下 60 twips
   - 视觉宽度函数：CJK/全角记 1、ASCII 记 0.5（`ord(ch) > 0x2E80`）
   - 实测收益：南沙区情手册 Word 版 **36 页 → 32 页**，且不再出现一列一字
6. **强制刷新域**：往 `word/settings.xml` 的 `<w:settings …>` 后插 `<w:updateFields w:val="true"/>`，Word/WPS 打开时自动填目录页码
7. **（可选）重写 `styles.xml` 贴齐 PDF 调性**：editor_sdk 给的默认样式很素（Normal 无字体字号、heading 1/2 各 18/16pt 纯黑）。要还原"素雅"版式时，直接在 zip 层整份替换 `styles.xml`：
   - `docDefaults` 补 `rPrDefault`（`Times New Roman` + 中文 `宋体`、`sz=21`、色 `1A1A1A`）与 `pPrDefault`（`line=330 auto`、`after=100`）
   - **styleId 是随机的**（本次 `czlvv5`=Normal / `va6utd`=heading 1 / `kipyav`=heading 2 / `wl9ruv`=Title / `lgc9ib`=Subtitle）→ 先用正则把 `<w:style … w:styleId="…">` 与紧随的 `<w:name w:val="…"/>` 配对列出来，再按 `w:name` 反查该改哪个 id
   - ⚠️ **每次 `build_docx.py` 重新生成，styleId 都会换一批**（同一份脚本两次跑出来就不一样）。所以**绝对不要把 styleId 写死在脚本里** —— 那样重跑一次，`document.xml` 引用的是新 id、`styles.xml` 里却只定义了旧 id，**整篇标题静默掉样式变成正文外观**，而且很难反查。`polish_docx.py` 已内置安全闸：重写前比对「正文引用的 pStyle」是否全部在新 styles.xml 里有定义，不满足直接中止
   - **`styles.xml` 一旦被改坏无法回滚**：改之前先把原始 styles.xml 留一份，或者干脆「改坏了就整份重建」
   - 标题改 `微软雅黑`，H1 加 `<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="A62B2B"/></w:pBdr>`（一条暗红细线，是"黑白红克制"最省的表现手法）
   - **Title / Subtitle 不要留 `<w:outlineLvl>`**：留着会让 `TOC \u` 把它们也收进目录
8. **封面补版式（`scripts/patch_cover.py`，按当次项目改文案）**：`build_docx.py` 落的封面只是 Title ＋ Subtitle 两行纯文字，跟 PDF 封面（红线／主色标题／数据卡／署名）差得远，用户一眼能看出来。zip 层把封面那几段整块换掉即可（换到**第一个 `<w:br w:type="page"/>` 之前**）：
   - eyebrow 小字（灰、`sz=18`）＋ **`<w:pBdr><w:bottom w:val="single" w:sz="18" w:space="2" w:color="<主色>"/>`** 一条主色细线（最省的主色表现手法）
   - 大标题分两行：第一行墨黑、第二行主色（各 `sz=56`、`b=1`）
   - 副标题灰 `sz=21` ＋ 一句主色提示行
   - **数据卡**：1×4 表，`<w:tblW>` 拉满正文宽（A4+2cm = **9638 twips**，四列 2410/2410/2409/2409），只留 `top`/`bottom`（黑 `sz=8`）与 `insideV`（浅灰 `sz=6`），数字 `sz=40` 主色加粗、单位 `sz=22` 主色、标签 `sz=16` 灰，全部 `w:jc=center`
   - 收尾：细灰上边框 ＋「署名」`sz=20` ＋「编制时间」`sz=18` 灰
   - **垂直间距用空段 `w:after` 顶**（两处：副标题→数据卡、数据卡→署名）
   - ⚠️ **封面一旦超过一页，会多出一整张空白页**（本次实测 37 页 vs 36 页：间距 5400/3800 twips 就溢出了，收到 4600/2950 才正好一页）。**留余量**：估算内容总高 ≤ 可用高的 ~85%（A4+2cm 边距 可用高 = 16838−2268 = **14570 twips ≈ 728pt**）。改完必须重渲染数页数
   - 封面表格会进 `doc_list_tables`，表数 +1 属正常

> ⚠️ **改完不要再用 editor_sdk 的 `save_file`**：它会按自己的内部模型重写 XML，把上面这些 zip 层后处理（列宽、cantSplit、PNG 双嵌、updateFields）全部抹掉。后处理必须走 `zipfile` 直接改 XML + 重打包。

**重打包必须保真 + 原子**：
1. 先 `z.extractall(TMP)`，再按**原 `namelist()` 逐条写回**（含 `word/`、`word/media/`、`docProps/`、`_rels/` 这些**目录项**），新增部件另加。
   > 踩坑：自己拼 namelist 时若用 `if nm.endswith("/"): continue` 跳过目录项，包内条目会从 29 掉到 23。`testzip()` 仍返回 None、Word 也照开，但没必要冒这个险——保持原样最稳。
2. **写到临时 zip，校验通过后再 `os.replace()` 覆盖成品**。
   > 血泪：`with zipfile.ZipFile(成品路径, "w")` 会**立刻截断成品**；只要后面任一步抛异常（例如新增部件忘了进写入循环），成品就被写成一个空包 / 半成品，且原文件已经没了。必须先写 `/tmp/xxx_out.docx`。
3. **写完做两道关系校验**（否则 Word 会报"无法打开，内容有问题"）：
   - **无悬空关系**：`document.xml.rels` 里每个 `Target` 都要在包内存在（跳过 `TargetMode="External"`）
   - **无未声明引用**：正文里所有 `r:id=` / `r:embed=` 都要能在 rels 里找到对应 `Id`
   > 本次加首页页眉时漏写 `header2.xml` 就制造过一个悬空 rId，靠这两道校验当场拦下。

### ⑦ 目检
- `doc_to_image`（editor_sdk 自带，把本地 docx 逐页渲染 PNG）——**远端腾讯文档转换服务，经常返回"服务繁忙" / 连接超时 / 不完整页数，必须重试循环**
- **判"渲染是否完整"的可靠办法：看末页正文墨迹**（去掉页眉页脚后 `(a<235).mean()`）。末页墨迹 ≈0 且页数比预期少 → 截断，重渲染；末页墨迹 >1% → 这次是完整的
  - 注意：**末页真空白**也可能是文档真的多了一页空白页（尾部空段没删干净），要和"截断"区分开——先去 XML 里数一遍末尾空段
- 逐页看：封面、目录页、每个图表页（图注是否在图下方独立成行）、**所有含长文本列的表格页**、最末页
- 文本层自检：字符数、段落数、表格数、图片数、残留 `**`/`#`/`[[CHART:` 是否为 0、章节是否齐全
- **改动前后做一次"文字内容全等"校验**：`''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', xml, re.S))` 两个包对比，字符串必须完全一致——证明只动了版式没动内容
- **只动局部时的逐页哈希法**：重渲染后对**未打算改动的页**做 `md5`（或 `cmp`）逐页比对。全等即证明改动确实被限制在目标页（本次补封面后第 3—36 页逐字节一致，只封面与目录变了）。比逐页看图快得多，也更能定位"多出一页空白页"这类问题
- **页数变化是首要信号**：渲染页数比预期 +1 就是"多了一张空白页"，先查封面/末尾空段，别急着看图
- 红线自检（沿用项目惯例）：第三方机构名、AI 痕迹词、外部域名

## 三、editor_sdk 已知缺陷速查

| 现象 | 原因 | 对策 |
|---|---|---|
| `document is not open` | `create_doc` 后立刻编辑 | 重试包装 |
| `document is not open`（另一处） | `doc_set_document_style` 同传 page_style + default_paragraph_style | 拆两次调用 |
| 只拿到部分标题 | `doc_resolve_document_structure` / `doc_get_outline` 默认截断 | 传 `mode="outline", limit=0, text_preview_length=0` |
| schema 命令报 `'list' object has no attribute 'get'` | 该工具 inputSchema 是数组 | 跳过 schema，直接 `--json` 调用 |
| 图注贴着图折行 | 图与图注同段 | 后处理拆段 |
| 目录显示"未找到目录项。" | 域未求值 | 后处理插 `updateFields` |
| 其它查看器看不到图 | SVG 无位图回退 | 后处理补 PNG 双嵌 |
| 表格断页处出现一行空白单元格 | 表格行被拆到两页 | 后处理给每个 `<w:trPr>` 加 `<w:cantSplit/>` |
| **表格全列等宽，长文本列一列一字** | `doc_insert_markdown` 落表默认等宽，`tblCellMar` 仅 15 twips | 后处理列宽自适应 + 满宽 + `tblLayout=fixed` |
| **封面与目录挤在同一页** | `doc_insert_page_break` 实际插到了目录**后面** | 后处理在目录标题前补分页符 |
| **末页多出一整页空白页** | 落款后留了 1–2 个空段落（含自闭合 `<w:p/>`） | 后处理从尾部剥空段（注意 sectPr 位置） |
| **封面带着重复书名的页眉和页码** | 编辑器只给了 `headerReference default`，没有 `titlePg`；而 PDF 原版封面是无页眉页码的整版设计 | 加 `<w:titlePg/>` + 新建空白首页页眉（`header2.xml`）+ 把 first 页脚置空。**两个坑**：① 编辑器其实**已经生成了 `footerReference w:type="first"`，且指向带 PAGE 域的 `footer1.xml`** —— 只加 `titlePg` 不重指，封面页码照旧出现；必须把该引用改成新建的空白页脚。② 编辑器产出的 `<w:footerReference r:id="…" w:type="first"/>` / `<w:headerReference r:id="…" w:type="default"/>` **属性顺序是 `r:id` 在前、`w:type` 在后**，正则写成 `<w:footerReference w:type="first" r:id="…"/>` 会**静默匹配不到**（不报错、封面页码仍在）。用 `<w:footerReference\b[^>]*w:type="first"[^>]*/>` 这种顺序无关写法 |
| 后处理成果被抹掉 | 又调了 `save_file` | 后处理必须走 zipfile，之后不要再 save_file |
| 页脚署名和页码只能留一个 | 两个工具互斥 | 署名进页眉 |
| **目录整页只有"未找到目录项。"** | TOC 域未求值，转换服务 / Quick Look / Finder 预览都不刷新 | 改**多段落域**：`begin` 段 ＋ 静态章名段 × N ＋ `end` 段，非更新型查看器看到干净目录；再配 `updateFields` |
| **Word 封面只是一行纯文字标题** | `build_docx.py` 的封面没有版式，与 PDF 封面差距肉眼可见 | 后处理 `patch_cover.py` 补主色细线／双色标题／数据卡表／署名 |
| **封面之后多出一整张空白页** | 封面内容（含空段 `w:after` 间距）超出一页，溢出的空段自己占了一页 | 收小间距；内容估算总高 ≤ 可用高（728pt）的 85%；改完重渲染数页数 |
| **`doc_get_outline` 对成品恒报 `document is not open`** | `open_file` 只是把盘上文件"开始"载入，未真正挂进编辑器；重试 8 次（间隔 3s）仍失败 | 别指望它拿页码/大纲；章名清单直接从 `document.xml` 按 `w:name="heading 1"` 反查 styleId 后抽取 |
| 重打包后包内条目变少 | 跳过目录项 | 按原 `namelist()` 全量写回 |
| **`polish_docx.py` 报"缺样式"后整份不动** | 白名单只认 `Normal/Title/Subtitle/heading 1/heading 2`；正文一旦用到 `heading 3`（卡片小标题常被 `###` 落成 h3）就判定"缺样式"直接放弃 | 已修：白名单加 `heading 3`，并新增 `extract_style()` 兜底——正文引用但本脚本未定义的 styleId，**从原 `styles.xml` 原样搬运过来**，不再整份放弃。安全闸保留：搬运后仍逐项核对 `used - defined` 必须为空 |
| **在"已有 docx 体检"流程里重跑 `fix_docx.py`，封面数据卡被压成等宽** | `fit_table` 会重算**所有**表的列宽，包括 `patch_cover.py` 手工排好的封面卡表 | 已修：`patch_cover.py` 给封面卡表打标记 `<w:tblCaption w:val="covercards"/>`；`fix_docx.py` 遇到该标记**原样放过**并单独计数（`跳过封面卡 N 张`）。自定义封面模板记得沿用这个标记 |
| 目录只有一行章名，右半边空 | 单级 `\o "1-1"` 撑不满一页 | 用两级：域改 `\o "1-2"`，静态兜底段同时抽 `heading 1`（加粗、无缩进，`sz=20`）与 `heading 2`（`<w:ind w:left="420"/>`、灰色 `3C3C3C`、`sz=18`）。按 `w:name` 反查两个 styleId 后逐段归类 |
| **整条 `.tl` 时间轴在 Word 版静默消失** | `.tl` 里没有可直接读的文本节点，每个字段都埋在叶节点 `div.it > span.lk/.lv`，通用的"段落/列表/纯文本 div"分支一条都收不到；HTML 版还在，**只有肉眼对比才发现** | 已修：`html2md.py` 加 `tl_md()`，把时间轴展开成「年份 × 字段」两维表（先扫全部年份收集列名，再按列名填值，缺失留空）。**交付前必检**：在 docx 文本层 grep 时间轴里的某个具体值（如 `预计 1 月下旬`），命中才算没丢 |

## 四、随附脚本（参考实现，按当次项目改路径/类名）

| 脚本 | 作用 |
|---|---|
| `scripts/html2md.py` | lxml 解析 HTML → Markdown，含上述全部类映射与空白折叠防坑。结构性容器（`.tl` 时间轴、`.bars` 条形图、`.stat` 数据卡、`.card`、`.tblab`）都有专门分支，**通用分支收不到的内容必须在这里显式展开**，否则 Word 版静默丢内容 |
| `scripts/build_docx.py` | editor_sdk 驱动：新建→样式→封面→目录→分块写入+插图→套标题样式→页眉页脚→保存 |
| `scripts/fix_docx.py` | **完全通用**的后处理器：`python fix_docx.py <目标.docx>` → 补高清位图回退 + 拆图注段落 + 表格列宽自适应/满宽/内边距 + updateFields。顶部 `FIT_TABLES`/`EXP`/`FLOOR`/`CELLMAR` 可按需调。**幂等**：对已处理过的成品再跑，文字/`<w:tcW>`/`<w:gridCol>`/条目数/分页符/`titlePg` 逐项一致（列宽由单元格文字重算，不读旧值）；带 `covercards` 标记的封面卡表自动跳过 |
| `scripts/fix_titlepg.py` | **完全通用**：`python fix_titlepg.py <目标.docx>` → 封面去页眉页码（`titlePg` + 空白首页页眉/页脚）。自动按 `w:name` 反查 styleId、顺序无关地重指 first 引用、自带两道关系校验 |
| `scripts/polish_docx.py` | **基本通用**：`python polish_docx.py <目标.docx> [ink] [red]` → 目录占位改 F9 提示 + 重写 `styles.xml`（宋体正文/微软雅黑标题/H1 暗红细线）。按 `w:name` 反查 styleId；`heading 3` 已纳入白名单，正文引用但未定义的样式会从原 `styles.xml` **原样搬运**后再做 `used - defined` 安全校验（不再整份放弃）。**换配色只需传两个色值**，如黑白册用 `1A1A1A 8A8A8A` |
| `scripts/patch_cover.py` | **参考实现**：`python patch_cover.py` → zip 层换掉封面（主色细线／双色大标题／副标题／1×4 数据卡表／署名）＋ 把 TOC 域改写成**两级**多段落静态兜底（`\o "1-2"`，H1 加粗 + H2 缩进灰字）。顶部 `CH` / `CARDS` / `build_cover()` 的文案与 `sp1`/`sp2` 间距按当次项目改。⚠️ 封面/目录定位**按内容不用 paraId**（每份文档随机，换项目必失效）：封面块＝**`<w:body>` 之后**到 `第一个 <w:br w:type="page"/> 段`之前，替换式必须是 `doc[:body_start] + new_cover + doc[m0:]`——若直接 `new_cover + doc[m0:]` 会把 XML 声明＋`<w:document>` 根元素一起吞掉，document.xml 变成裸 `<w:p>` 开头，python-docx/lxml 全部报 "Namespace prefix w not defined"（有 .bak 从备份重跑即可）；目录域＝含 `instrText …TOC` 的整段 |
| `scripts/wordgeom.py` | **完全通用**（需 `python-docx`）：`python wordgeom.py <目标.docx>` → 无渲染器时的版式几何估算：正文区自动取自 `sectPr`，逐页报占用高度/填充率，`<72%` 标「空白偏多」。用于「已有 docx 体检」，比真渲染快且不依赖网络。详见第六节 ⑥-1 |
| `scripts/renumber_figs.py` | **完全通用**：`python renumber_figs.py <目标.docx>` → 合并章节后按 H1 章号重排图注编号（`图 11-1 → 图 9-1`）。原地改、自动留 `.bak`。详见第六节 ⑥-2 |

**执行顺序**：`build_docx.py` → `fix_docx.py` → `fix_titlepg.py` → `patch_cover.py` → `polish_docx.py`（每一步都是 zip 层直改，**之后绝不能再调 `save_file`**）
> `patch_cover.py` 与 `polish_docx.py` 都碰目录域：**用了 patch_cover 的多段落兜底就不必再跑 polish_docx 的第 1 件事**（它找的是"未找到目录项。"，已不存在会自动跳过）；polish 的 `styles.xml` 那半仍然要跑。

`html2md.py` / `build_docx.py` 顶部的 `BUILD` / `CHARTS` / `MD` / `TARGET` / `COVER` 常量按当次项目改；`fix_docx.py` 无需改动。

### 自检：`fix_docx.py` 改完要做**双向测试**

只验证"重跑一遍结果不变"是不够的——那也可能说明它压根没干活。两次都要做：

1. **幂等性**：对已处理过的成品再跑一次 → 正文文字、`<w:tcW>`、`<w:gridCol>`、条目数、
   分页符数、`titlePg` 全部**逐项一致**；计数器除"表格张数/cantSplit 行数"外应报 0
   ```
   图注段落拆分: 0    封面/目录分页符: 0    末尾空段落删除: 0    新增位图 0 张
   ```
2. **修复力**：人为构造损坏件再跑 —— 删掉封面后的分页符 + 把某张表所有列改成等宽
   `<w:gridCol w:w="4153"/>` → 跑完应**精确还原**成正确值（本次实测首表还原为 `['5839','3799']`）
   ```
   封面/目录分页符: 1     ← 说明它真插了
   ```

自检脚本骨架：先 `zipfile` 读原包 → 改 `document.xml` 字符串 → 按原 `namelist` 重打成
`broken.docx` → 跑 `fix_docx.py` → 对比签名（正文文字 + tcW + gridCol + 条目数）。

**签名要连封面卡一起比**，否则最大的一处不幂等会被漏掉。抽法：
`re.search(r'<w:tblCaption w:val="covercards"/>.*?</w:tblGrid>(.*?)</w:tblGrid>', doc, re.S)`
—— 该表列宽必须稳定为 `2410/2410/2409/2409`。同时比整份 `document.xml` 的 md5，
"跑两次 sha 相同"是最省事的兜底判据。

## 五、交付

- 落盘位置：**当前工作空间**下的一单一个文件夹（`年份+地区+项目+产品名/`），Word 与同名 PDF 并列（2026-09-18 项目负责人明确，不要再写桌面或其他自建目录）
- 命名：`<资料全名>.docx`（与 PDF 完全同名，仅扩展名不同）
- 用 `present_files` 呈现，PDF 与 Word 一起传
- 中间产物（脚本、PNG、渲染页图）放 `/tmp/<项目>_work/word/`，交付后清理，**工作区文件夹只留最终件**
- 告知用户：目录是 Word 域，打开时自动刷新；若显示为空，右键目录 → 更新域（或 F9）

## 六、对**已有** docx 做版式体检与合并精简（不重新生成）

用户拿来一份现成 docx 说「排版不美观 / 有多余空缺 / 合并简化」时走这条，**不要**回头重跑渲染管线。

### ⑥-1 先量，再改：几何估算器 `scripts/wordgeom.py`

本机常见「无 Microsoft Word、无 LibreOffice、Mac 版 WPS 不支持 `--convert-to`」——
`doc_to_image`（唯一的 docx 渲染口）是**远端腾讯文档服务，会整段时间返回"服务繁忙"/`code 50000`/`437009`**，
此时**没有任何真渲染件可看**。用几何估算器替代目检：

```bash
python scripts/wordgeom.py <目标.docx>
# 正文区 459.2 × 717.2 pt
# 估算页数：27
#   P01  637.5pt  88.9%  █████████…         ← <72% 标「空白偏多」
```

模型：`段落高 = 行数 × 字号 × 1.30 × 行距倍数 + 段前 + 段后`；中文 1em、西文 0.5em 估宽；
表格**按列宽独立换行取行内最大**，并**逐行**判断换页。四个必须踩对的坑（错一个页数差 3–8 页）：

| 坑 | 正确做法 |
|---|---|
| 默认行距/字号/段后写死 | **从 `docDefaults` 读**（`w:pPrDefault/w:pPr/w:spacing`，`lineRule="auto"` 时 `w:line/240`）。本档实测默认 1.5 倍行距，按 1.0 算会把空白全算错 |
| 单元格内**末段**的 `space_after` 也计入行高 | Word **不计**——末段段后要扣掉，否则每张表多算几 pt，全篇虚增约 7 页 |
| 表格整张当不可分割块 | 只有 `<w:cantSplit/>` 才阻止**单行内部**断开；行与行之间**可以**跨页。必须逐行累加判页 |
| 只认 `page_break_before` | 编辑器产的文档更常用**段落内** `<w:br w:type="page"/>`。两种都要检测 |

正文区宽高**自动从 `sectPr` 推导**（`page_width - left - right`），换项目不用改常量。
字号/列宽/字距仍是经验值 → **只用来横向比"改前 vs 改后"与定位空白带，不要当绝对页数用**。

### ⑥-2 合并章节后**必须回正图号**

合并（如「行测 3 章 + 申论 4 章 → 2 章」）会让图注里的旧章号全错（第九章里出现「图 11-1」）。
`scripts/renumber_figs.py` 按 **H1 章号 + 章内出现顺序** 重排：

```bash
python scripts/renumber_figs.py <目标.docx>     # 原地改，自动留 .bak
# 图 11-1 → 图 9-1 / 图 12-1 → 图 10-1 / 图 13-1 → 图 10-2 …
```

前提：图注段落的**编号整体在首个 run 内**（本档如此，`run0="图 11-1"`）。脚本内置兜底：
定位不到就把编号 + 余文重建进首个 run。改完**扫一遍全篇 `图 \d+-\d+` 查重复**。

### ⑥-3 目录页码要用**同一个已修正的模型**重算

目录是「TOC 域 + 静态页码缓存」。**模型改一次（补了 `set_defaults`、修了 `cell_height`），
就必须重跑一次目录重建**，否则页码是旧的、整篇偏小 1–3 页。
自检：把 H1 逐个映射到模型页号，与目录缓存逐条比，要求 **N/N 全等**：

```
编制说明  模型 2  目录 2     第九章 行测赋分标准与模块权重  模型 11  目录 11
…
附录C　申论赋分总表与数据口径  模型 26  目录 26     ← 19/19 全对才算过
```

### ⑥-4 内容零丢失：**双向**比对

```python
# 去掉标题、只留 ≥20 字的正文段 + 全部表格签名（行列数 + 首行文本）
旧版未出现 → 逐条列出；表格签名逐张 `==` 才算一致
```
改动只应产生**有意差异**：合并掉的旧章标题、重编号的图注、合并重写的小节。
本次 149→146 段、41 张表逐张一致、11 条「未出现」**全部**可归因到上述三类 —— 这才敢交付。

### ⑥-5 交付与命名（并行会话冲突）

同一份资料**可能有另一个会话在并行改写同名文件**。判据：看 `/tmp/<项目>_docx/` 里
脚本/渲染图的**时间戳是否仍在前进**（本次相隔 2 小时后确认停滞）。
- 不停 → **不覆盖**它的目标文件，另存 `-排版优化版.docx` 待用户确认
- 已停 → 可按用户选择接管正式名，**但原件先 `cp` 进 `~/.Trash/<资料>_原件备份_<日期>/`**（不用 `rm`）
- 若拿不到 Word→PDF 通道（`doc_to_image` 限流 + 无本机 Office），**如实说明 PDF 仍是旧版**，
  给用户「WPS 打开 → F9 更新目录 → 导出 PDF」的替代路径，不要用截图拼一个假 PDF

## 七、无 Office 时的 docx → PDF 离线生成（`scripts/docx_to_pdf.py`）

当 **Word→PDF 全部通道都不通**（云端 `doc_to_image` 限流 + 本机无 MS Word / LibreOffice / WPS CLI），
用 reportlab 直接读 docx 结构重排一份矢量 PDF（文本层可提取、可合规 grep）：

```bash
python scripts/docx_to_pdf.py <源.docx> <输出.pdf>
```

要点（全部踩过）：
1. **中文字体**：`/System/Library/Fonts/Supplemental/Songti.ttc`（正文宋）+ `STHeiti Medium.ttc`（标题黑）
   用 `TTFont(..., subfontIndex=0)` 注册；Hiragino Sans GB 是 postscript outline，**reportlab 不支持**。
   `registerFontFamily("Song", bold="Hei")` 让 `<b>` 自动切黑体。
2. **缺字形是隐形杀手**：Songti 子集**没有 U+2212（−）也没有 U+2013（–）**，渲染成空白且提取不到。
   生成前用 `getFont(name).face.charToGlyph` 扫全文档，把缺字形字符映射到确定存在的等价物（−→ASCII -）。
   **只换缺的，别把 ＋ ＝ ÷ 这些好端端的也动了**（它们有字形）。
3. **目录页码两遍编译会不收敛**：目录行若用整段 markup，tab 后的旧页码混进左列，右列新页码被提取器拆走；
   左列必须**截断到 tab 前**（`cut_tab`）。页码写入用「迭代到不动点」跑 2–4 遍直到 h1_pages 稳定，
   校验用 `get_text(sort=True)` 按几何配对左右两列。
4. **表格**：列宽按 `gridCol` 比例缩放到正文宽；`repeatRows=1` 跨页重复表头；首行>1 行才加表头底纹
   （单行表多为卡片/数据条）。合并单元格靠 `row.cells` 的 `_tc` 去重。
5. **页码合规**：封面页 1 无页眉页脚（署名在封面），其余页页眉带「广东中公教研」
   → 「中公」出现次数 == 总页数，沿用项目红线校验。
6. **填充率窗口**：排除页眉线与页脚（A4@72dpi 取 `im[58:786]`），否则每页虚报 100%。
7. **目检替代**：当前会话看不了渲染图时，用程序化体检顶上——横向越界（blocks bbox）、
   图注与图是否同页、字体嵌入清单、逐段/逐单元格内容比对（注意：跨页段会被页眉页脚打断造成假阳性）。
