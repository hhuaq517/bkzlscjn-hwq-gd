# 更新日志

本仓库采用「首发 + 逐版修订」的简单记录方式。版本号跟技能包整体走。

---

## v1.0.0 · 2026-09-24 · 首发

### 内容

收录六个协同工作的技能，构成一条完整的备考资料生产线：

- **`exam-prep-suite`** — 总入口。判线、路由、交付铁律、跨产线共性坑。
- **`exam-pdf-render`** — 共享渲染层。HTML → 印刷级 PDF，双引擎（Playwright / 无头浏览器），
  封面全出血单独渲染、页眉页脚与页码、目录页码回填、填充率与越界巡检、缩略图总览。
- **`exam-situation-handbook`** — 考情手册（招考形势研究册）。九种手册原型、公告抓取、
  岗位表批量解析、历年参数汇总。
- **`exam-prep-pdf`** — 备考讲义 / 题本。考情调研、版式配色、图表与思维导图、分册编排。
- **`mock-exam-paper`** — 模拟卷 / 冲刺卷。模块配比锁定、跨套去重、时政时效管控、
  答案分布拉平、Word 直出。
- **`exam-paper-restore`** — 真题还原。模糊截图 OCR、长截图缺页检测、学科卷专项、答案打匀。
- **`exam-docx-export`** — Word 导出。HTML → Markdown → editor_sdk 写入 → zip 层后处理重打包。

### 工程约定

- 补 `README.md`（安装 / 环境 / 快速开始 / 配色 / 交付铁律 / 常见问题）、
  `LICENSE`（MIT + 附加说明）、`.gitignore`、`requirements.txt`（按技能分段并标注可选依赖）。
- `requirements.txt` 区分「必需」与「可选」，并注明 `mock-exam-paper` / `exam-prep-suite` 零依赖。
- `.gitignore` 一律按文件名 / 目录名精确排除，**不使用 `*.html` / `*.png` / `*.pdf` 整类规则**
  —— 整类规则会静默吞掉模板与素材，且 `git add -A` 不报错。
- 中间产物默认落在系统临时目录，仓库树内不留任何运行产物。

---

## 相对原提交包的改动清单

原包共 **28** 个文件（含根 `README.md` 与根 `SKILL.md`）。
本仓库共 **32** 个文件。逐文件 MD5 比对结果：

| 项 | 数量 |
|---|---|
| 原包文件 | 28 |
| **逐字节一致（未改动）** | **12** |
| 有改动 | 16 |
| 新增（原包没有） | 4 |

逐字节一致、一行未动的 12 个文件：

```
skills/exam-docx-export/scripts/docx_to_pdf.py
skills/exam-docx-export/scripts/renumber_figs.py
skills/exam-docx-export/scripts/wordgeom.py
skills/exam-paper-restore/scripts/vision_ocr.py
skills/exam-pdf-render/scripts/stamp.py
skills/exam-pdf-render/scripts/style.css
skills/exam-prep-pdf/scripts/render.py
skills/exam-prep-pdf/scripts/style.css
skills/exam-situation-handbook/scripts/check_fill.py
skills/exam-situation-handbook/scripts/fetch_attachments.py
skills/exam-situation-handbook/scripts/overflow_check.py
skills/exam-situation-handbook/scripts/parse_positions.py
```

### A. 结构性调整（1 项）

| 改前 | 改后 | 原因 |
|---|---|---|
| 根目录 `SKILL.md`（总纲） | `skills/exam-prep-suite/SKILL.md` | 让六个技能在 `skills/` 下平级并列，安装只需一条 `cp -R skills/* ~/.workbuddy-ai/skills/`。**文件内容除下述去标识化外未改。** |

### B. 去标识化（内部人名 → 通用称呼）

| 文件 | 改前 | 改后 |
|---|---|---|
| `skills/exam-docx-export/SKILL.md` | 「<内部人名>的常见说法：」 | 「用户的常见说法：」 |
| `skills/exam-docx-export/SKILL.md` | 括号内署了内部人名，并点名了一个内部目录 | 「（2026-09-18 项目负责人明确，不要再写桌面或其他自建目录）」 |
| `skills/exam-paper-restore/SKILL.md` | 小节标题内署了内部人名 | 「卷面从简（2026-09-20 项目负责人明确要求…）」 |

> 上表「改前」列只描述被替换的内容类型，不复述具体人名 —— 否则等于把已脱敏的信息在日志里再公开一次。

### C. 可移植化（写死的本机绝对路径 → 通用写法）

原包中 `/Users/<本机用户名>/…` 一类的绝对路径共 12 处，全部处理；处理后仓库内 `/Users/` 命中数为 **0**。

**文档层（7 处）**

| 文件 | 改前 | 改后 |
|---|---|---|
| `exam-prep-suite/SKILL.md` | `/Users/…/.workbuddy/binaries/python/envs/default/bin/python` | `python3`（附依赖说明） |
| `exam-pdf-render/SKILL.md` | 同上（环境段 + 渲染命令） | `python3` |
| `exam-prep-pdf/SKILL.md` | 同上 | `python3` |
| `exam-prep-pdf/SKILL.md` | `/Users/<用户名>/Library/Caches/ms-playwright/…` | `~/Library/Caches/ms-playwright/…` |
| `exam-paper-restore/SKILL.md` | `/Users/…/.workbuddy-ai/…/python3`（隔离解释器） | 改为「确认解释器里装了 `pymupdf`」的通用表述 |
| `exam-situation-handbook/SKILL.md` | `/Users/…/.workbuddy-ai/…/python`（管理包隔离） | 改为「为本套技能单建一个 venv」 |
| `mock-exam-paper/SKILL.md` | `PY=/Users/…/python` | `PY=${PY:-python3}` |

**脚本层（5 个文件）**

| 文件 | 改前 | 改后 |
|---|---|---|
| `exam-docx-export/scripts/build_docx.py` | `TARGET = "/Users/…/Desktop/广州市南沙区情手册（2026版）.docx"`；`WORK`/`CHARTS` 写死 `/tmp/ns_qing_work/…`；`EDSDK` 写死应用路径 | 全部改为环境变量覆盖 + 临时目录 / 当前目录默认值：`EXAM_EDSDK` / `EXAM_WORK` / `EXAM_CHARTS` / `EXAM_TARGET` |
| `exam-docx-export/scripts/fix_docx.py` | 默认目标写死桌面路径；`TMP` 写死 `/tmp/ns_qing_work/word/_fix` | 目标改为**必填位置参数**（缺参给用法提示）；`TMP` 走 `tempfile.gettempdir()` |
| `exam-docx-export/scripts/fix_titlepg.py` | 同上；`TMP`/`OUT` 写死 `/tmp/_titlepg_*` | 同上 |
| `exam-docx-export/scripts/polish_docx.py` | 默认目标写死桌面路径 | 目标改为必填位置参数 |
| `exam-docx-export/scripts/patch_cover.py` | 默认目标写死桌面路径（用户名还是占位符 `xxx`，等于把"该填什么"留给使用者猜） | 目标改为必填位置参数 |
| `exam-docx-export/scripts/html2md.py` | `BUILD`/`OUT` 写死 `/tmp/ns_qing_work/…` | `EXAM_BUILD` / `EXAM_MD_OUT` 环境变量覆盖 + 临时目录默认值 |

> **渲染逻辑一行未动。** 上述改动只涉及路径解析与参数入口，不影响任何排版、分页、样式计算结果。

### D. 缺陷修复（1 项，实测暴露）

| 文件 | 问题 | 修法 |
|---|---|---|
| `exam-pdf-render/scripts/check.py` | 孤行页判据写的是 `if 0 < fill < 10`。**完全空白的页** `body` 为空 → 填充率恰为 `0.0`，被 `0 <` 排除后**静默漏检**，`check.py` 仍返回退出码 0。而空白页恰恰是最该修的一种孤行页。 | 改为 `if fill < 10`，与文件自身的文档字符串（`填充率 < 10% -> 孤行页`）以及 `check_fill.py` 的既有写法对齐。 |

复现与回归（实测）：

```
# 构造：在 5 页成品末尾追加一张空白页
$ python check.py bad.pdf --sign "页脚署名文字" --skip-cover --top 55
  p6   填充   0.0%  块 0    <-- 孤行页          ← 修复前此行无标记、退出码 0
孤行页 1 个 / 越界页 0 个                        ← 修复后退出码 1

# 回归：同一脚本跑无空白页的成品
$ python check.py out/01_verify.pdf --sign "页脚署名文字" --skip-cover --top 55
孤行页 0 个 / 越界页 0 个                        ← 退出码 0，无误报
```

### E. 文档一致性订正（1 项，实测暴露）

原文档（根 `SKILL.md` 与 `exam-pdf-render/SKILL.md`、`exam-prep-pdf/SKILL.md`）都写：

> `style.css` 的 `@page` margin 必须与 `render.py` 的 `margin` **完全一致**。
> 默认配套：`@page{…margin:17mm 15mm 16mm 15mm}` ↔ `{"top":"17mm",…}`

**实测结论与此不符**，两点：

1. **随附的 `style.css` 写的是 `18mm`，不是 `17mm`** —— 文档与随附模板不一致。
   （`style.css` 内部是自洽的：封面补偿 `.cover{margin:-18mm -15mm -16mm -15mm}` 正是按 18mm 算的。）
2. **`docs.json` 的 `margin` 字段根本不生效**：只要 `style.css` 声明了非零 `@page` margin，
   CSS 就是权威值。实测把 `margin` 参数依次设为 **17mm / 18mm / 30mm**，成品三个测量值完全相同 ——
   正文内容顶边均 `19.5mm`、页眉 y 均 `16.5pt`、页脚 y 均 `826.1pt`。

已按实测把三处文档改写成准确表述，并在 `render.py` 的 `SCHEMA` 里给 `margin` 字段加了
「被 CSS 顶掉」的注释。**`DEFAULT_MARGIN` 常量保持原值 17mm 未动**（不改行为）。

### F. 新增文件（4 个）

`README.md`（重写）、`LICENSE`、`.gitignore`、`requirements.txt`（后三个原包没有）。

### G. 公开前脱敏：移除「绕过文档站访问控制」的做法（1 节）

原包 `exam-paper-restore/SKILL.md` 有两处写明了从文档分享站取全文的具体绕过做法。
为避免在日志里把已脱敏的技术细节再公开一次，此处**只描述改动性质，不复述具体做法**：

| 位置 | 改动性质 | 处理 |
|---|---|---|
| 「换源：找回同一份试卷的公开原文 → 取高清页面图」整节 | 该节逐个站点写明了绕过其访问控制的具体步骤（含伪造爬虫身份、拼装需要凭据的预览接口、按文件名规律枚举下载） | **整节重写**为：只用站点自己提供的公开预览能力；明确写出**不要伪造爬虫身份、不拼装需要凭据的接口、不按文件名枚举下载**；站点只放出前几页时就只用这几页，确需完整原件走合法渠道取得 |
| 「学科知识卷专项 → 题源」小节 | 点名站点并给出「一次抓取就能拿到整卷题干＋选项＋答案＋逐题详解」的取源路径，并逐站比较抓取难度 | 改为「用公开题库按题干关键句检索核对原题」，并补一条「只用公开提供的内容、不绕过访问控制、确需原件走合法渠道」；站点难度对比删除 |

同时把该节里指向具体客户交付物与具体年份高考卷的表述泛化（「某地市教招生物卷」「某年某省选考」），
与仓库其余部分的案例泛化口径一致。

> **主体能力未受影响**：OCR 还原（`vision_ocr.py`）、长截图缺页检测、学科卷手写 SVG、
> 答案不打匀等全部保留。脱敏只动了「去哪拿全文」这一段。

脱敏后终检（逐词扫描，全部 0 命中）：本机绝对路径 · 内部人名 · 内部目录名 · GitHub 账号 ·
文档站域名与相关接口标识 · 任何凭据形态（`ghp_` / `github_pat_` / `password=` / `api_key=`）。

---

## 已知约束

1. **字体是 macOS 取向。** CSS 字体栈写的是 `Hiragino Sans GB` / `STHeiti` / `Songti SC`，
   其他平台需自行替换为中文字体。仓库**不附带任何字体文件**。
2. **`vision_ocr.py` 仅 macOS 可用。** 它调用系统 Vision 框架（pyobjc）。
3. **`build_docx.py` 是一次实际交付的驱动实例**（南沙区情手册），
   可当模板照改：封面文案、图表键名、页眉文字、章节标题都是那一次的实例值。
   写入逻辑本身通用。
4. **`build_docx.py` 依赖 editor_sdk**，随 WorkBuddy / CodeBuddy 内置的
   `tencent-local-office-edit` 技能提供；路径可用 `EXAM_EDSDK` 覆盖。
5. **`docx_to_pdf.py` 里的 `HDR`（页眉文字）与 `SIGN`（署名）是按项目改的常量**，
   换项目需自行修改。
6. **封面不计页。** A 线把封面单独渲染后合并，正文页码与目录页码都以正文为准
   （封面自带署名带、不显示页码），因此正文第 1 页在成品里是第 2 张纸。
7. **A 线的目录页码回填依赖 PDF 文本层**。若章节标题用图片或 `display:none` 呈现，
   回填会失败并打印 `未解析 -> …`。
8. 仓库**不含任何考试真题原卷、招考数据、职位表或第三方平台凭据**。

---

## 使用许可

MIT，见 [LICENSE](LICENSE)。文档中引用的机构名称与品牌标识不随本许可授权。
