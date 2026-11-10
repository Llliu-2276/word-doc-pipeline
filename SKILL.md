---
name: word-doc-pipeline
description: 将中文技术文档 / 课程设计文档从 Markdown 转换为符合排版规范的 Word（.docx），并提供写作约束检查、成品格式校验与批量构建：A4 页面、页眉页码、宋体/Times New Roman、固定 22 磅行距、首行缩进、三级标题、三线表、图注表注按章编号。当用户要编写、转换、检查或批量生成中文 Word 文档（课程设计说明书、可行性分析报告、需求/设计/测试文档、规范文档），需要把 Markdown→Word 文档管线迁移到新项目，或排查 docx 的页面、字体、行距、表格、图表注、页码问题时使用。
---

# 中文文档 Markdown → Word 流水线

## 用途

本 Skill 把「Markdown 写源文件、脚本转 docx、人工补封面目录、导出 PDF」这套文档工程固化下来：

- `scripts/`：转换与校验脚本，确定性执行，不依赖 Word。
- `references/`：文档工程约定、Word 成品参数、Markdown 写作约束、图表规范、工具链、验收清单。
- `assets/templates/`：新文档骨架、图表绘制说明模板、gitignore 追加片段。

转换器排版参数与 `references/02-Word格式规范.md` 一一对应；改格式先改脚本常量，再同步规范。

## 快速接入新项目

1. 在项目根创建 `docs/`，按文档类型分子目录（如 `docs/需求文档/`、`docs/设计文档/`）；每份文档一个 `.md` 源文件，图片放同目录 `images/`。
2. 把本 Skill 的 `scripts/` 复制到项目 `scripts/`（或保留 Skill 目录、用绝对路径调用）。
3. 设置项目名（页眉与校验共用，可省略则无页眉）：
   - Windows PowerShell：`$env:DOC_PROJECT_NAME="项目名"`
   - bash：`export DOC_PROJECT_NAME="项目名"`
4. 复制 `assets/templates/文档骨架.md` 起稿；需要画图时复制 `图表绘制说明模板.md` 到文档目录维护图清单。
5. 把 `assets/templates/gitignore.append` 内容追加到项目 `.gitignore`。

## 标准流程

1. **写 Markdown**：严格使用 `references/03-Markdown写作约束.md` 允许的语法子集。
2. **检查源文件**：`python scripts/check_md.py docs`（目录扫描默认跳过 `*绘制说明.md`、`*核查报告.md` 工作笔记）。
3. **转换单篇**：`python scripts/md2docx.py docs/<类型>/<文档>.md docs/<类型>/<文档>.docx`。
4. **校验成品**：`python scripts/check_docx.py docs/<类型>/<文档>.docx --project "项目名"`。
5. **批量转换**：文档较多时用 `python scripts/build_docs.py docs --project "项目名"`，转换后自动逐篇校验。
6. **Word 手工步骤**（脚本不做）：补封面、插入目录并「更新域」、检查分页与图表位置、另存 PDF。
7. **交付前验收**：按 `references/06-验收检查清单.md` 逐项勾选。

## 命令速查

| 命令 | 作用 |
| --- | --- |
| `python scripts/md2docx.py <in.md> <out.docx> [--project 名称] [--no-header] [--no-footer] [--page-size A4]` | 单篇转换（默认 A4、页眉项目名、页脚页码） |
| `python scripts/check_md.py <md 或目录>... [--all]` | 检查禁用语法、图注表注、图片缺失/宽度、图号连续性 |
| `python scripts/check_docx.py <docx 或目录>... [--project 名称]` | 检查纸张、边距、页眉页码、图片宽度、三线表、正文字体行距 |
| `python scripts/build_docs.py <docs 目录> [--project 名称] [--include "*.md"] [--skip-suffix ...] [--dry-run] [--no-check]` | 批量转换并校验 |

`check_md.py` / `check_docx.py` 存在错误时退出码为 1，可直接用于 CI 或提交前拦截。

## 硬性格式（返工项）

- 纸张 A4（21.0×29.7cm），上下边距 2.54cm、左右 3.18cm；python-docx 默认是 Letter，必须显式设置。
- 正文：宋体小四 12pt，西文/数字 Times New Roman，固定行距 22 磅，首行缩进 2 字符，两端对齐。
- 标题：一级黑体三号 16pt 不加粗、二级黑体小三 15pt 不加粗、三级黑体小四 12pt 加粗；段前段后各 0.5 行、无缩进。四级标题写作 `**4.4.1 标题**` 加粗段落。
- 图注/表注：宋体小五 9pt、居中；图注在图下，表注在表上；按章编号「图 4-1 说明」「表 4-1 说明」，附录可用「图 A-1」。
- 三线表：上下框线 1.5 磅、表头下框线 0.75 磅、无竖线与内横线、表头加粗、跨页重复标题行；表内文字 10pt。
- 图片：嵌入式、最大宽 14.5cm；图片段落 1.5 倍行距（固定行距会裁切图片），图注行距仍为 22 磅。
- 页眉可选项目名（宋体小五），页脚页码居中「- 1 -」。
- 封面与目录不进入 Markdown：封面在 Word 手工排版，目录用 Word 标题样式生成并更新域。

## 资源导航

| 资源 | 何时读取 |
| --- | --- |
| `references/01-文档工程与目录约定.md` | 规划 docs 目录、md/docx/pdf 成对规则、提交版命名、文档间变更传播 |
| `references/02-Word格式规范.md` | 调整排版参数、核对字体字号段落图表规格、解释脚本行为 |
| `references/03-Markdown写作约束.md` | 写文档前；查某语法是否支持、禁用语法如何替代 |
| `references/04-图表与表格规范.md` | 画图、截图、编号、写图表注、维护图表绘制说明时 |
| `references/05-工具链与环境.md` | 第一次安装依赖、换机器、排查 python-docx/字体/PDF 工具问题 |
| `references/06-验收检查清单.md` | 每次转 docx 后、导出 PDF 前、提交前 |
| `examples/示例文档.md` | 需要可转换的最小样例，验证脚本是否正常 |

## 脚本接口要点

- `md2docx.py` 支持的 Markdown 子集：`#/##/###`、管道表格、`-`/`1.` 列表、`**加粗**`、`---`、`![图注](images/figX-Y.png)`、表格上方 `**表 x-x 说明**`。表注与表格之间允许空行；其余不支持语法会告警或退化，不会中断转换。
- `md2docx.py` 只处理正文；封面与目录永远在 Word 中处理，重新生成 docx 后需重新补封面、更新目录域。
- `check_docx.py` 若不传 `--project`，只做通用检查，不校验页眉文字；页眉项目名也可由 `DOC_PROJECT_NAME` 提供。
- `build_docs.py` 递归查找 md，跳过 `资料/`、`附件/`、`.git/`、`node_modules/` 与工作笔记后缀；`--dry-run` 先看清单。

## 迁移到新项目时要改的配置

集中在 `scripts/md2docx.py` 顶部常量与函数：

- 字体：`FONT_HEI`、`FONT_SONG`、`FONT_LATIN`。
- 字号：`SIZE_H1/H2/H3/BODY/CAPTION/CELL`。
- 行距：`LINE_SPACING`（22 磅）、`IMG_LINE_SPACING`（1.5）。
- 页面：`PAGE_SIZES` 与 `setup_page()` 中的边距。
- 图片最大宽度：`MAX_IMG_WIDTH`（14.5cm）。
- 项目名：`--project` 或环境变量 `DOC_PROJECT_NAME`，不要在脚本里硬编码某个项目名。

同步修改 `references/02-Word格式规范.md` 与 `references/05-工具链与环境.md` 中的对应数值，避免规范与脚本脱节。

## 已知坑

- 表注必须独占一行且以 `**表 x-x` 开头；普通加粗段落不会被当作表注。
- 表格单元格内不支持换行、合并单元格与竖线 `|`；需要复杂表格时拆表或改为列表。
- 图片路径相对 md 文件解析；图片缺失只告警并跳过，转换不会失败，需看 stderr。
- 固定行距会裁切行内图片，图片段落已自动改为 1.5 倍行距，勿手动改回。
- 生成的 docx 中页脚页码是 Word 域，首次打开或打印前需更新域；目录同理。
- 校验脚本只能发现结构性错误，封面版式、图表位置、孤行孤页仍需人工目视。
