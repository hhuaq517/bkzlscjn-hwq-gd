# 备考资料生产线 · Skill Bundle

一套面向**考公 / 考编 / 事业单位 / 国企央企 / 公安辅警 / 社区专职**招聘考试的
**印刷级资料生产技能集**。六个技能协同完成一条链路：一句话需求 → 分流到对应内容产线
→ 统一走共享渲染层 → 输出可直接进印刷厂 / 打印机的 A4 PDF，或可编辑的 Word。

> 本仓库是**技能包**（skill bundle），不是可直接运行的应用。
> 安装方式见「[安装](#安装)」，六个技能装到同一个技能目录下即可互相路由。

---

## 一、这套技能解决什么问题

做一册备考资料，实际要处理四类事，每类都容易翻车：

| 类别 | 典型翻车 | 本套技能的对应能力 |
|---|---|---|
| **内容体系** | 考情手册写成了学习资料、模拟卷和真题还原搞混 | `exam-prep-suite` 总纲先判线再路由，附「易混点」清单 |
| **数据来源** | 拿网络经验当官方口径；岗位表用转载页截图 | 只认官网；附件必须下载原件解析；非公开数据一律标「测算」 |
| **版式与印刷** | 页脚压正文、表格截断、孤行页、目录页码全 0 | `exam-pdf-render` 共享渲染层 + 自动巡检 + 缩略图目检 |
| **交付与合规** | 出现竞品机构名、暴露数据源、留下 AI 痕迹 | 交付铁律 + 红线 grep 清单，收尾强制核对 |

---

## 二、六个技能

```
skills/
├── exam-prep-suite/          ← 总入口：判线、路由、交付铁律、跨产线共性坑
├── exam-pdf-render/          ← 共享渲染层：HTML → 印刷级 PDF（其余五个的下游）
├── exam-situation-handbook/  ← 考情手册（招考形势研究册）
├── exam-prep-pdf/            ← 备考讲义 / 题本
├── mock-exam-paper/          ← 模拟卷 / 冲刺卷 / 押题卷
├── exam-paper-restore/       ← 真题还原（模糊截图 → 完整试卷）
└── exam-docx-export/         ← 导出可编辑 Word(.docx)
```

### `exam-prep-suite` — 总入口

用户通常只丢一句话（「汕头辅警考情」「做几套模拟卷」）。本技能负责**判线 → 路由 → 盯交付铁律**，
不重复实现子技能里的排版细节。

路由表（节选）：

| 用户原话（典型） | 走哪条线 | 本质 |
|---|---|---|
| 「XX考情」「XX地区XX考试考情」 | `exam-situation-handbook` | 帮考生**看懂这场考试**、判断值不值得报 |
| 「做一份XX备考资料 / 讲义 / 题本」 | `exam-prep-pdf` | 帮考生**学会怎么考** |
| 「做几套模拟卷 / 冲刺卷 / 押题卷」 | `mock-exam-paper` | 出**新题** |
| 「还原这份试卷 / 把这张图做成试卷」 | `exam-paper-restore` | 把**已有**真题还原成卷 |
| 「转成 Word / 导出 docx」 | `exam-docx-export` | 已有资料 → 可编辑 `.docx` |
| 「排版 / 渲染 PDF / 目录页码」 | `exam-pdf-render` | 所有产线的共同下游 |

**三个易混点**（判错就整册重做）：

- **考情手册 ≠ 备考讲义**：前者是"招考形势研究册"（公告机制、竞争比、薪酬趋势），
  后者是"学习材料"（知识点、答题框架），数据源完全不同。
- **真题还原 ≠ 模拟卷**：前者还原已有题，后者新命题。
- **试讲（教师岗）≠ 结构化面试（综合岗）**：每次必须用官方公告核对面试形式再动笔。

### `exam-pdf-render` — 共享渲染层

**只管"把 HTML 变成能直接印刷的 PDF"，不管内容怎么写。** 两条线：

| | A 线 · 正式成品 | B 线 · 轻量单文件 |
|---|---|---|
| 脚本 | `scripts/render.py` | `scripts/stamp.py` |
| 引擎 | Playwright（Chromium） | 无头 `chrome-headless-shell` / Edge / Chrome |
| 封面 | 全出血单独渲染再合并 | 无独立封面（首页跳过盖章） |
| 页眉页脚 | Playwright 模板，原生页码 | PyMuPDF 二次盖章（内置中文字体） |
| 目录页码 | 白色锚点 + 二次渲染自动回填 | 占位符替换 + 重渲染 |
| 适用 | 正式交付、多分册、要封面要目录 | 快速出样、单文件、无 Playwright |

**默认走 A 线。** 附 `check.py` 做填充率 / 越界 / 孤行页巡检并生成缩略图总览 PDF。

### `exam-situation-handbook` — 考情手册

招考形势研究册，内置九种手册原型（单区单项目 / 跨区汇总 / 面试专项 / 部委直属事业单位 /
公安辅警 / 国企央企校招 / 政策性银行 / 社区专职 …）。附：

- `fetch_attachments.py` — 批量抓公告页 → 提取岗位表附件 → 下载（处理相对路径、中文 URL 编码、Referer 要求）
- `parse_positions.py` — 岗位表批量解析（`xlsx` / `xls` 双通道）汇总成 JSON
- `overflow_check.py` — HTML 定页溢出检测（JS 注入法，逐页给溢出像素值）
- `check_fill.py` — PDF 填充率像素法巡检

### `exam-prep-pdf` — 备考讲义 / 题本

考情调研 → 版式配色 → 图表与思维导图 → 分册编排 → 渲染。含独立 `render.py` 与 `style.css`。

### `mock-exam-paper` — 模拟卷 / 冲刺卷

模块配比锁定、跨套去重、时政时效（白名单制 + 年份黑名单断言 + 月份配额）、
答案分布拉平、专有名词跨套重复检测、Word 直出。

### `exam-paper-restore` — 真题还原

模糊截图 OCR（macOS Vision）、长截图缺页检测、学科卷专项（含手写 SVG 物理图形）、答案打匀。

### `exam-docx-export` — Word 导出

`HTML → Markdown → editor_sdk 写入 → zip 层后处理重打包`。九个脚本各管一段：

| 脚本 | 作用 |
|---|---|
| `html2md.py` | HTML → 干净 Markdown（供 Word 使用） |
| `build_docx.py` | 驱动 editor_sdk 生成 .docx（封面 / 目录 / 分块正文 / 图表双嵌 / 页眉页脚） |
| `fix_docx.py` | 后处理：SVG 补高 DPI PNG 位图回退、表格列宽自适应、行禁止跨页、`updateFields` |
| `fix_titlepg.py` | 首页不同（封面不显示页眉与页码） |
| `patch_cover.py` | 封面补版式 + 目录域静态兜底 |
| `polish_docx.py` | 重写 `styles.xml`（正文 / 标题字体 + H1 下边线） |
| `renumber_figs.py` | 图表编号重排 |
| `docx_to_pdf.py` | docx → PDF 离线生成（reportlab，两遍编译回填目录页码） |
| `wordgeom.py` | 逐行高度与分页几何计算 |

---

## 三、安装

把 `skills/` 下的**六个子目录**整个复制到用户级技能目录：

```bash
cp -R skills/* ~/.workbuddy-ai/skills/
```

装好后在对话里直接说需求即可触发，不需要记技能名：

| 你说 | 得到 |
|---|---|
| 「汕头辅警考情」 | 一册考情手册 PDF |
| 「做一份南沙综合岗面试讲义」 | 分册备考讲义 PDF |
| 「根据这个公告做三套模拟卷」 | 试卷 + 答案解析（PDF / Word 双版） |
| 「把这张试卷截图还原成 PDF」 | 印刷级完整试卷 |
| 「把这份资料转成 Word」 | 可编辑 `.docx` |

> **建议整套安装。** 只装总入口不装子技能，总入口会路由到不存在的技能。

---

## 四、环境要求

### Python

需要 **Python 3.10+**。建议为本套技能单建一个 venv，不要污染系统环境：

```bash
python3 -m venv ~/.venvs/exam-prep
source ~/.venvs/exam-prep/bin/activate
pip install -r requirements.txt
playwright install chromium          # A 线渲染引擎
```

**字体**（macOS 系统自带，无需额外安装）：

- 标题：`Hiragino Sans GB` / `STHeiti`
- 正文：`Songti SC`

Linux / Windows 下需自行替换 CSS 字体栈中的中文字体名。

### 无头浏览器

B 线脚本按以下顺序自动探测，找到第一个就用：

1. playwright 的 `chrome-headless-shell`（`~/Library/Caches/ms-playwright/`）
2. Microsoft Edge
3. Google Chrome
4. `/usr/bin/chromium`、`/usr/bin/google-chrome`

`stamp.py` 支持 `--browser PATH` 显式指定。

---

## 五、快速开始：渲染一本 PDF

以 `exam-pdf-render` 为例。

**1. 准备目录**

```
<工作区>/<地区>/          只放最终 PDF
_build/<地区>/            HTML / CSS / JSON / 中间 PDF，交付前清掉
```

**2. 复制样式模板起步**

```bash
cp skills/exam-pdf-render/scripts/style.css _build/gd/style.css
```

改 `:root` 里的 CSS 变量即可换配色，**不要动 `@page` 之外的结构**。

**3. 写正文 HTML**

封面**不要**写进 body —— 封面由 `docs.json` 的 `cover` 字段描述、单独渲染。
目录条目用 `<span class="pg" data-s="锚点文本">` 占位。

**4. 写 `docs.json`**

字段见 `render.py` 末尾的 `SCHEMA` 常量，或直接运行 `python render.py` 不带参数查看：

```json
[
  {
    "name": "d1",
    "palette": "gov",
    "cover": {
      "kicker": "2026 · 广州南沙",
      "title": "结构化面试理论讲义",
      "subtitle": "综合岗 · 编外人员集中招聘备考",
      "desc": ["搭建知识框架", "从“考什么”到“怎么答”"],
      "chips": ["第一分册 · 理论"],
      "band_l": "广州南沙",
      "band_r": "页脚署名文字"
    },
    "header_l": "结构化面试理论讲义",
    "header_r": "广州南沙 · 综合岗",
    "footer_l": "署名文字",
    "body_html": "/abs/path/d1_body.html",
    "out_pdf": "/abs/path/out/01_xxx.pdf"
  }
]
```

**5. 渲染**

```bash
python skills/exam-pdf-render/scripts/render.py docs.json
```

输出示例：

```
TOC: 3 项 全部命中
[OK] /path/out/01_verify.pdf  共 5 页
```

**6. 巡检**

```bash
python skills/exam-pdf-render/scripts/check.py out/01_verify.pdf \
    --sign "署名文字" --skip-cover --thumb _build/overview.pdf
```

退出码 `0` = 干净，`1` = 有孤行页或越界页。**填充率 < 10% 判定为孤行页，必须修到 0。**

**7. 目检**

打开 `_build/overview.pdf` 逐页看。**这一步是底线，不能省。**

**8. 清理**

确认无误后清掉 `_build/`，成品目录只留 PDF。

---

## 六、配色预设

`render.py` 内置五套，`docs.json` 里用 `"palette"` 选，`"palette_override"` 细改单色：

| 值 | 风格 | 主色 |
|---|---|---|
| `gov` | 政务 / 综合岗 | 藏青 `#12395C` + 琥珀金 `#B8842B` + 松石绿 `#2C7A72` |
| `edu` | 教师 / 教育 | 墨绿 `#1F5C4A` + 暖橙 `#C4703A` |
| `med` | 医疗 / 卫健 | 青蓝 `#106B7A` + 珊瑚红 `#D2604F` |
| `fin` | 金融 / 财经 | 深蓝 `#1B3A6B` + 金 `#B58A2B` |
| `plain` | 素雅通用 | 石墨 `#39414A` + 砖红 `#A65D46` |

**同一批次内配色必须统一，不同项目之间必须换一套，同一地市 / 同一系列沿用同一套。**

---

## 七、交付铁律

1. **一项目一文件夹**，文件夹内**只留最终 PDF**。脚本 / HTML / 临时图表全部放
   `_build/<项目>/` 或系统临时目录，交付前清干净。
2. **版式**：封面 + 目录（**带真实页码**）+ 分部分 / 分章节，印刷级 A4。
3. **内容**：多维度展开；**教师岗**附当地学校简介与在用教材版本，**非教师岗不附**。
4. **美化**：多种图表混用（时间轴 / 雷达图 / 流程图 / 思维导图 / 对比表 / 环形图 / 柱状 / 折线）。
5. **页脚**：极小的整理署名 + 「第 X 页 / 共 Y 页」。
6. **红线**（交付前必 grep）：
   - 除本机构外**不出现任何其他机构名**；
   - **不展示数据源**，不写"据 XX 网转载"；
   - 报名平台域名模糊化（`.gov.cn` 政府门户可保留，其余改写为"以公告原文为准"）；
   - **无 AI 痕迹**（不写"本文由 AI 生成""仅供参考"之类的自我指涉）；
   - 非公开数据（竞争比、薪资）可估算，但必须标注**"参考 / 测算"**。

---

## 八、常见问题

**Q：`render.py` 报 `ModuleNotFoundError: No module named 'playwright'`**
A：A 线需要 playwright。装 `pip install playwright && playwright install chromium`，
或改用 B 线 `stamp.py`（只需系统无头浏览器 + PyMuPDF）。
注意 `render.py` 在模块顶层 import playwright，所以**连不带参数查看 `SCHEMA` 也需要先装好**。

**Q：`TOC: N 项 未解析 -> ...`，目录页码填成 0**
A：`data-s` 的文本在 PDF 文本层里找不到。最常见原因是目录写了"四、面试形式与成绩构成"，
而正文 `h2` 只有"面试形式与成绩构成"。修法二选一：`data-s` 直接用正文标题原文；
或在正文插**白色锚点** `<span style="color:#fff;font-size:6pt">场次锚一</span>` 并让 `data-s` 指向它。
`display:none` 的文本不进 PDF 文本层，搜不到 —— 要"不可见但可搜"只能用白色。

**Q：目录页码全填成了目录页的页码**
A：目录里的标题文字本身会被 `search_for` 命中。在目录页开头与结尾各放一个
**白色**小字标记「导航页」（`nav_marker`，默认值即此），巡检时会跳过含该标记的页。

**Q：页脚署名"看起来有"，但合规 grep 数出来是 0**
A：页眉 / 页脚的 `letter-spacing` 必须 < 1px（`render.py` 里已是 0.7px）。
字距大时 Chromium 会把一行按字符拆成多个 text run，抽出文本变成「署 名 文 字」。
另外**红线校验必须用 `pymupdf`，不能用 `pypdfium2`** —— 后者不做 run 合并，
带字距的 CJK 页脚会被逐字拆开，恒为 0。

**Q：改了 `docs.json` 的 `margin`，版心一点没变**
A：正常。**CSS 的 `@page` margin 才是权威值**：只要 `style.css` 声明了非零
`@page{margin:...}`，`margin` 字段就被顶掉。实测把参数从 17mm 改到 30mm，
成品正文顶边（19.5mm）、页眉 y（16.5pt）、页脚 y（826.1pt）三个测量值完全不变。
**改版心请改 `style.css` 的 `@page`**，并把 `docs.json` 的 `margin` 同步成同一个值仅作声明用。

**Q：某页只剩几行，或者整页空白**
A：孤行页。修法优先级 **内容搬家 > 删内容 > 拉伸留白** —— 把上一页有空档的另一章
小节整段搬过来（改编号即可），一次同时修好"溢出页"和"偏空页"且不损失内容。
空白页同样会被 `check.py` 报为孤行页（填充率 0%）。

**Q：`.grid2` / `.grid3` 卡片组下一页顶部出现只剩边框的空卡**
A：grid 容器被分页切开。给 `.grid2,.grid3` 与 `.card` 都加 `page-break-inside:avoid`。

**Q：提示框（`.tip` / `.warn` / `.key`）跨页后横向溢出并压住页脚**
A：这类带 `::before` 绝对定位徽标的框必须加 `page-break-inside:avoid` ——
被跨页分片时 Chrome 会**按整页宽度而非内容宽度**渲染分片。

**Q：Windows / Linux 上字体不对**
A：本套技能的字体栈是 macOS 取向（`Hiragino Sans GB` / `STHeiti` / `Songti SC`）。
其他平台请在 CSS 里换成对应中文字体，或安装上述字体的等价替代。

---

## 九、目录结构

```
.
├── README.md                 本文件
├── LICENSE                   MIT（含附加说明）
├── CHANGELOG.md              版本与相对原包的改动清单
├── requirements.txt          依赖清单（按技能分段，标注可选）
├── .gitignore
└── skills/
    ├── exam-prep-suite/          总入口
    │   └── SKILL.md
    ├── exam-pdf-render/          共享渲染层
    │   ├── SKILL.md
    │   └── scripts/{render.py, stamp.py, check.py, style.css}
    ├── exam-situation-handbook/  考情手册
    │   ├── SKILL.md
    │   └── scripts/{fetch_attachments.py, parse_positions.py,
    │                overflow_check.py, check_fill.py}
    ├── exam-prep-pdf/            备考讲义 / 题本
    │   ├── SKILL.md
    │   └── scripts/{render.py, style.css}
    ├── mock-exam-paper/          模拟卷 / 冲刺卷
    │   └── SKILL.md
    ├── exam-paper-restore/       真题还原
    │   ├── SKILL.md
    │   └── scripts/vision_ocr.py
    └── exam-docx-export/         Word 导出
        ├── SKILL.md
        └── scripts/{html2md.py, build_docx.py, fix_docx.py,
                     fix_titlepg.py, patch_cover.py, polish_docx.py,
                     renumber_figs.py, docx_to_pdf.py, wordgeom.py}
```

---

## 十、许可证

MIT，见 [LICENSE](LICENSE)。

**附加说明**：仓库文档中出现的机构名称、品牌标识、页眉署名口径等，仅为说明交付规范而引用，
**不随本许可授权**。仓库**不附带任何考试真题原卷、招考数据或字体文件**；
文中提到的试卷样例、考情数据、公告原文等仅作格式与口径示例，
使用者如需相关素材应自行通过合法渠道取得并自行承担内容合规责任。
