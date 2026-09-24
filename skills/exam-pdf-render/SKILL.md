---
name: exam-pdf-render
description: 把 HTML 排版成印刷级 PDF 的共享渲染管线——双引擎渲染、封面全出血、页眉页脚与署名页码、目录页码回填、填充率与越界自动巡检、缩略图目检。当考公考编/事业单位/国企类资料需要"输出最终 PDF"时使用；exam-prep-pdf、exam-situation-handbook、mock-exam-paper、exam-paper-restore 四个内容技能的排版环节统一走本管线。触发词：渲染PDF、生成PDF、HTML转PDF、排版导出、目录页码回填、PDF巡检、压页脚、孤行页。
agent_created: true
---

# 印刷级 PDF 渲染管线（共享层）

**本技能只管"把 HTML 变成能在印刷厂/打印机上直接用的 PDF"，不管内容怎么写。**
内容体系、考情调研、署名红线等由上层内容技能负责（`exam-prep-pdf` / `exam-situation-handbook` / `mock-exam-paper` / `exam-paper-restore`）。那四个技能做到"HTML 写完了"这一步之后，排版与导出统一按本文件执行，避免各自重复踩坑。

## 一、选管线

| | A 线 · 正式成品 | B 线 · 轻量单文件 |
|---|---|---|
| 脚本 | `scripts/render.py` | `scripts/stamp.py` |
| 引擎 | Playwright（Chromium） | 无头 chrome-headless-shell / Edge |
| 封面 | 全出血单独渲染再合并 | 无独立封面（首页跳过盖章） |
| 页眉页脚 | Playwright 模板，原生页码 | PyMuPDF 二次盖章（`china-ss` 内置中文字体） |
| 目录页码 | 白色锚点 + 二次渲染自动回填 | 占位符替换 + 重渲染 |
| 适用 | 正式交付、多分册、要封面要目录 | 快速出样、单文件、无 Playwright |

**默认走 A 线。** 只有"单文件 HTML 且环境没有 Playwright"时才用 B 线。

环境：
```
python3   # 需已装 playwright + pymupdf
```
字体：标题 `"Hiragino Sans GB","STHeiti"`，正文 `"Songti SC"`。

## 二、A 线流程

1. **建目录**：`<工作区>/<地区>/`（只放最终 PDF）、`_build/<地区>/`（HTML/CSS/JSON/中间 PDF，交付前清掉）。
2. **写 `style.css`**：直接复制 `scripts/style.css` 起步，改 `:root` 里的配色变量即可换风格。
3. **写 `dX_body.html`**：封面**不要**写进 body，单独用 `docs.json` 描述。
4. **写 `docs.json`**：字段见 `render.py` 末尾 `SCHEMA`，或运行 `python render.py` 不带参数查看。
5. **渲染**：
   ```bash
   python3 \
       scripts/render.py docs.json [name]
   ```
6. **巡检**：`python scripts/check.py out.pdf --sign "署名" --skip-cover --thumb _build/overview.pdf`
7. **目检**：打开缩略图总览逐页看，**这是底线，不能省**。
8. **清理**：确认无误后删掉 `_build/` 下所有中间产物，成品目录只留 PDF。

## 三、B 线流程

```bash
python scripts/stamp.py in.html out.pdf --sign "署名文字" --skip-stamp  # --skip-stamp=只渲染
python scripts/stamp.py in.html out.pdf --sign "署名" --toc-map anchors.json
```
- `file://` 路径含空格/中文**必须** urlencode（`stamp.py` 已处理）。
- 浏览器自动探测顺序：playwright 的 `chrome-headless-shell` → Edge → Chrome。
- `--toc-map` 的锚点用**章节正文首句**，不要用章节标题（标题会同时命中目录页导致误判）。

## 四、关键坑（每条都是实测翻车过的）

1. **CSS 的 `@page` margin 是权威值，`docs.json` 的 `margin` 字段会被它顶掉**。
   实测：`style.css` 写 `@page{margin:18mm 15mm 16mm 15mm}` 时，把 `docs.json` 的 `margin` 分别设成
   **17mm / 18mm / 30mm**，成品三个测量值完全一致 —— 正文内容顶边均 `19.5mm`、页眉 y 均 `16.5pt`、
   页脚 y 均 `826.1pt`。即：**只要 `style.css` 声明了非零 `@page` margin，`margin` 字段就不生效**，
   它只在 CSS 完全没有 `@page` margin 规则时才起作用（那时走 `render.py` 的 `DEFAULT_MARGIN`）。
   - 后果：照着 `style.css` 改版心时，若只改 CSS 不改 `docs.json`，**不会报错也不会报错位**，
     只会让人误判"改没生效"。**改版心一律以 `style.css` 为准**，`docs.json` 的 `margin` 只作声明性备注。
   - 随附 `style.css` 的默认配套是 `@page{size:A4;margin:18mm 15mm 16mm 15mm}`（封面补偿
     `.cover{height:262mm;margin:-18mm -15mm -16mm -15mm}` 与之配套，**改一边必须改另一边**）。
   - 反面情形：`@page{margin:0}`（封面 HTML 就是这么写的）会把版心顶到纸边，
     而 Playwright 的 header/footer 模板仍按 `margin` 参数排版 → 出现"页眉页脚正常、正文贴边"的错位。
     封面单独渲染时本来就该这样，正文 HTML **不要**写 `@page{margin:0}`。

2. **封面必须单独渲染再合并**。Playwright 的 header/footer 模板对所有页生效，无法只跳过封面。

3. **目录页码误匹配**：目录里的标题文字会被 `search_for` 命中，导致页码全填成目录页。两个办法：
   - 目录页放一个**白色**（`#FFFFFF`）小字标记"导航页"，搜索时跳过含该标记的页；**目录跨页时开头结尾各放一个**。
   - 图表/思维导图里的标签**不要**与章节标题用完全相同的字符串。

4. **`data-s` 必须等于正文实际渲染出来的文本**。`search_for` 在 PDF 文本层找不到就返回 0。最常见翻车：目录写"四、面试形式与成绩构成"，正文 `h2` 只有"面试形式与成绩构成" → 全部回填为 0。
   二级条目改用**白色锚点**定位：正文插 `<span style="color:#fff;font-size:6pt">场次锚一</span>`，`data-s="场次锚一"` —— 不可见但可搜。
   注意：`display:none` 的文本**不进 PDF 文本层**，搜不到；要"不可见但可搜"只能用白色。

5. **`position:relative` 的提示框必须加 `page-break-inside:avoid`**。`.tip/.warn/.key` 这类带 `::before` 绝对定位徽标的框，一旦被跨页分片，Chrome 会**按整页宽度而非内容宽度**渲染分片，横向溢出并压住页脚。

6. **渲染后必须跑填充率巡检**，比逐页目检快得多：
   ```bash
   python scripts/check.py book.pdf --sign "署名" --skip-cover --thumb _build/overview.pdf
   ```
   `填充 <10%` = 孤行页，`下溢/右溢` = 压页脚或横向溢出，两者都必须修到 0。

7. **孤行页修法优先级：内容搬家 > 删内容 > 拉伸留白。** 某章结尾溢出 1—3 行、导致下一页近空白时，最优解是把上一页有空档的**另一章的某个小节整段搬过来**（改编号即可）——一次同时修好"溢出页"和"偏空页"，且不损失内容。实测一次搬家即从 24 页降到 22 页、消除全部孤行页。

8. **分页密度**：把 `page-break-inside:avoid` 加在"题干+答案"整块上，会在页底留大片空白。改为只在**题干行**加 `page-break-after:avoid`、答案允许跨页，页数可压缩约 8%，观感更紧凑。

9. **同一文件多个 Edit 必须串行提交**，并行会互相覆盖丢更新。批量替换改用 Python 脚本一次原子改写。

10. **PDF 转换不要和 `rm` 串在一条命令里**，转换失败会连带把源文件删掉。

11. **`.grid2/.grid3` 卡片组必须加 `page-break-inside:avoid`**。grid 容器被分页切开时，Chrome 会在下一页顶部渲染出**只剩边框的空卡**（内容都留在上页，下页只留一个空圆角框，极易漏检）。修法：`.grid2,.grid3{page-break-inside:avoid;}` 配 `.card{page-break-inside:avoid;}`。

12. **章节强制分页会制造大量半空页**。`.chapter{page-break-before:always}` 让每章结尾留白，实测某册 37 页中有 6 页填充率低于 35%。**内容型手册**（区情册、资料册，非讲义）建议改 `.chapter{page-break-before:auto}`，配 `.chapter-head{page-break-inside:avoid;page-break-after:avoid}` + `h3,h4{page-break-after:avoid}`——实测 **37 页 → 27 页**且孤行页清零；再把 `body` 行高收到 1.74、字号 10.1pt、`p{margin:6px 0}`，可再压 1—2 页。

13. **说明页 / 目录页偏空时，补数据卡比拉大字号有效**。在编制说明页尾部加一排 12 格「核心数据速记」卡（`grid3` × 4 行，数字用 15pt 粗体、说明用 8.2pt 灰字），实测填充率 **41.4% → 71.5%**，且顺带提升了速查实用性。

14. **选项行 / 长文本行的「右溢 5—6pt」，根因是 `letter-spacing` + 行尾全角空格**。正文若设了 `body{letter-spacing:.2px}`，一行 50 字会累积约 10pt 宽度；再叠上选项之间用 `　　`（U+3000）分隔、而 CSS 行尾空格是「悬挂」不参与换行计算的，行宽就会稳定顶出内容框 5—6pt，`check.py` 报「右溢」。**修法（一次根治）**：把选择题选项从纯文本改成 flex 原子块——
    ```python
    # 构建期把 <div class="opts">A．x　　B．y</div> 拆成 span
    items = [x.strip() for x in re.split(r"[\u3000]+", body) if x.strip()]
    spans = "".join(f'<span class="op">{x}</span>' for x in items)
    ```
    ```css
    .qitem .opts{display:flex;flex-wrap:wrap;column-gap:1.7em;row-gap:.06em;
                 letter-spacing:0;font-size:10.15pt;}
    .qitem .opts .op{white-space:nowrap;}
    ```
    实测右溢从 2 处降到 0，且不再复发。**注意**：只要选项里单个选项不超过一行宽度（约 48 个汉字），`nowrap` 就是安全的。

15. **表单类页面（答题卡、得分统计表）会拖出「末页只留一个 docend」的孤行页**。表格行默认 `padding:7px 10px`，十来行的表格很容易把收尾块挤到下一页。修法：给这类表格加紧凑类 `table.tb.cpt td{padding:3.5px 8px;line-height:1.5;}` `table.tb.cpt th{padding:4.5px 8px;}`，实测省下 60—70pt，一次收拢。

16. **页眉/页脚的 `letter-spacing` 必须 < 1px**（`render.py` 里已从 2px 收到 **0.7px**）。Chromium 在字距较大时会把一行文本**按字符拆成多个 text run** 落进 PDF 文本层，抽出来就是「广 东 中 公 小 红 书 运 营 部」。**视觉完全正常**，但「署名出现次数 == 页数」这条合规巡检会**恒为 0**，白白怀疑自己漏了署名。同理，正文里 `letter-spacing:.2px` 以上要留意第 14 条的右溢。

> **⚠️ 关键补充：0.7px 只对"能合并 run 的提取器"才算解决 —— 合规校验必须用 `pymupdf`，不能用 `pypdfium2`。**
> 实测同一份 11 页 PDF（页脚 `letter-spacing:0.7px`）：
> - `pymupdf`（`import pymupdf as fitz`）→ `广东中公教研` **11 次＝页数** ✅，`中公` 也 11 次
> - `pypdfium2`（`get_textpage().get_text_range()`）→ 紧凑串 **0 次** ❌，只有「广 东 中 公 小 红 书 运 营 部」逐字带空格
>
> pypdfium2 不做 run 合并，因此它**永远**会把带字距的 CJK 页脚拆开。用它跑红线 grep 还会掩盖真实命中（如「中公教育」被拆成「中 公 教 育」搜不到）。
> **结论**：红线 grep（机构名 / 域名 / 关键词）与「署名次数 == 页数」两件事，一律用 `pymupdf`；`pypdfium2` 只用来**渲页面图做目检**（它渲染更快）。两者别混用。
> 兜底判据：若怀疑漏署名，用宽松正则 `广\s*东\s*中\s*公\s*小\s*红\s*书\s*运\s*营\s*部` 逐页数——视觉在就一定能匹配到。

17. **中途 section 的 `docend` 会变成「书中的『完』」**。多 section 拼装时，若每个 section 末尾都放 `docend`，前一个 section 的「完」会落在正文中间，紧跟着下一章标题，观感突兀。**只在最后一节保留 `docend`**；删掉中间那个的净收益还能顺带把末页拖尾收回来（实测 48 页 → 47 页、孤行页清零）。

18. **收尾自检的三条硬指标**：`check.py` 退出码 0（孤行页 0 / 越界页 0）；**全页填充率不低于 80%**（比 skill 默认的 10% 阈值严得多，实测能筛出「四分之一页留白」这类机器不报但眼睛看得见的问题）；目录 48 项级别的 `data-s` 全部命中。低于 80% 的页优先「补实质内容」（如补一段实务提示、一张对比表），而不是拉行距。

18. **目录条目会被 `body` 行高撑爆**。`.toc .l1/.l2/.part` 若继承 `body{line-height:1.84}`，24 条目录会跨成两页、第二页只有 27% 填充。**修法**：给目录条目单独设行高，`.toc .l1{line-height:1.35}` `.toc .l2{line-height:1.3}` `.toc .part{line-height:1.3}`，并把 `.toc h2` 的 `margin-bottom` 收到 11px 左右。实测 24 条目录稳定收进一页。

19. **多「编」结构的题本册，编之间不要强制分页**。给每个 `part` 加 `page-break-before:always` 会在每编末尾留下 25%—47% 的空白页（实测某册 6 处）。**改为 `page-break-before:auto`**，只保留「答案/解析编」另起页；实测 27 页 → 24 页且留白清零。编标题本身有渐变底条，接在上一编末尾观感不受影响。

20. **末尾「只剩 2 题」或「只剩一个 docend」的孤行页**，两条实测有效的解法：
   - **补实质内容**：题本册末尾若只剩一两道题，补一张「答题卡 + 得分统计表」（题号格子留空 + 正确率统计 + 分档自测建议）。实测填充率 **19.4% → 86.2%**，且对考生真有实用价值。
   - **并入表格**：`docend` 若怎么调 `margin-top` 都要单独占一页，就把它做成最后一个答案表的一行（`colspan` + `border:none`），或直接删掉——强行保留只会制造孤行页。

## 五、配色预设

`render.py` 内置 5 套，`docs.json` 里用 `"palette"` 选择，可用 `"palette_override"` 细改单个色值：

| 值 | 风格 | 主色 |
|---|---|---|
| `gov` | 政务 / 综合岗 | 藏青 `#12395C` + 琥珀金 `#B8842B` + 松石绿 `#2C7A72` |
| `edu` | 教师 / 教育 | 墨绿 `#1F5C4A` + 暖橙 `#C4703A` |
| `med` | 医疗 / 卫健 | 青蓝 `#106B7A` + 珊瑚红 `#D2604F` |
| `fin` | 金融 / 财经 | 深蓝 `#1B3A6B` + 金 `#B58A2B` |
| `plain` | 素雅通用 | 石墨 `#39414A` + 砖红 `#A65D46` |

**同一批次内配色必须统一，不同项目之间必须换一套。**

> **仅当客户明确要求"只用黑白"时才走下面这套**，别当默认用。做法是选 `plain` 打底 + `palette_override` 把所有色值压成灰阶（深色面用 `#23282E → #0D1012`，强调色全部换成 `#14171A` / 灰阶），正文侧 CSS 变量同步换成 `--ink:#14171A`、`--gray:#6B7378`、`--line:#D3D7DA` 一类的灰阶体系，图表改用 `g1..g4` 四级灰而不引入彩色。2027 广东事业单位统考 / 2027 广东南网第一批 / 2027 国考三册用过。

## 六、脚本速查

```
scripts/render.py   A 线：Playwright 渲染 + 封面合并 + 目录回填（读 docs.json）
scripts/stamp.py    B 线：无头浏览器 + PyMuPDF 盖章页脚页码 + 占位符回填
scripts/check.py    巡检：填充率 / 越界 / 孤行页 / 缩略图总览
scripts/style.css   样式模板：改 :root 配色变量即可换风格
```

`check.py` 退出码 0 = 干净，1 = 有孤行页或越界页。
