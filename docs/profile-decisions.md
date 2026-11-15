# Profile 通用化决策记录

> 决策日期：2026-11-11。依据：`docs/document-type-specs.md`（调研）与本轮逐项问答。
> 本文只记录"定了什么、为什么"；具体数值仍以 `profiles/*.toml` 为准。

## 一、决策清单

| # | 议题 | 决策 | 理由 |
| --- | --- | --- | --- |
| 1 | 第一期范围 | profile 框架 + 现有基线正规化 + `tech-doc` | 两个 profile 都贴近现状，先用低成本验证"脚本读配置"这条路；学位论文、公文分别放第二、三期 |
| 2 | 默认 profile | `ug-thesis-bit`，裸命令输出与今天完全一致 | 已验证的项目零改动；新类型显式 `--profile` |
| 3 | 配置格式 | TOML（`tomllib`，Python 3.11+ 标准库） | 零新依赖、支持注释、可标注每个数值的规范来源；本机 Python 3.12.10 |
| 4 | 结构件 | 本期只做目录域 + 分节页码；封面仍手工 | 目录域与分节已实测可用（`TOC \o "1-3" \h \z \u`、`oddPage`）；封面版式依赖各校模板，抽象成本高，推到学位论文期 |
| 5 | 缺字体 | profile 内 `fonts.fallback` 链替换 + stderr 明确告警，不中断 | 不硬失败（否则很多机器跑不起来），也不静默（否则送出不合规文件而不自知） |
| 6 | tech-doc 参数 | 按调研表格的惯例值：2.54/3.18、小四宋体、1.5 倍行距、章三号黑体/节四号黑体、图表名五号黑体、首页不编、三线表 | 国标只规定结构不规定排版，此为项目惯例；能同时压测倍数行距、图注黑体、首页不同三个新开关 |
| 7 | 配置模型 | schema 一次铺满，实现分期；未实现字段读入后显式报"未实现" | 避免 profile 文件结构改两遍；死字段要吵，不能静默忽略 |
| 8 | 校验器 | `check_docx.py` 读同一份 profile 派生期望值 | 单一事实源，根治"规范与脚本脱节"这个历史痛点 |
| 9 | profile 派生 | 支持 `base = "父 profile"` + 深度合并局部覆盖 | 各校只需覆盖差异项（约 10 行），避免拷贝 150 行完整文件 |
| 10 | 未核实值 | 标 `verified = false` + 注明候选值/分歧来源；运行提示"含 N 项未核实"，校验器对这些项只警告 | 不阻塞开发，又不让未验证数字被当成权威 |
| 11 | 国标版本 | 默认 GB/T 7713.1—2025、GB/T 7714—2025（截至 2026-11-11 均已生效），留 `references_standard = "2015"` 开关 | 按现行国标办事；学校模板仍写旧版时按开关回退。本期只影响规范描述与将来的参考文献校验 |
| 12 | 代码组织 | 新增 `scripts/profile.py`（加载/继承合并/字段校验），命令形式不变，只多 `--profile` | 保持"复制 scripts/ 进项目"的用法；不重构成 pip 包 |
| 13 | 文档同步 | 本期同步改 `SKILL.md` 与 `references/02-Word格式规范.md` | 旧文档把单一数值写成"硬性格式"，与多 profile 事实冲突，会误导后来者 |
| 14 | 提交 | 完成后提交并推送到 GitHub（含 `docs/` 调研与决策文档） | 沙箱内 push 需 `danger-full-access` 重试 |
| 15 | `tech-doc` 六处惯例值核实（2026-11-15） | 核实后仍保留 `verified = false`，把结论写进文件头注释与各块 `note` | GB/T 8567-2006 只规定章节结构与内容要素，正文明言“不规定具体的布局和字体”；这六项与中文文档通行层级同构但无该类型的标准依据，标 `true` 会把惯例冒充规范 |
| 16 | 校验闸门收紧与生成标记（2026-11-15） | `md2docx.py` 在 `docProps/custom.xml` 写 `wdp_*` 自定义属性（generator/profile/page_size/header/footer/toc/chapter_break 等）；`check_docx.py` 读取后决定豁免；缺页眉、缺页码域、缺 TOC 域、三线表缺边框、嵌入图与内联图数不符一律升级为 error；无标记的旧文档严格判定并给一条警告；标记里的 profile 与本次 `--profile` 不一致直接报错 | 旧闸门把这些只记为 warnings，而 `build_docs.py` 只把 errors 计入失败，等于漏页码的文档能过 CI。难点是 `--no-header/--no-footer/--no-toc` 是合法选项，必须把「profile 要求而缺失」与「生成时显式关闭」分开——docx 自己带着生成参数才能分清，顺带根治 K12（用错 profile 校验报假差异） |
| 17 | 第二期四开关（2026-11-15） | 实现 `page.gutter`、`page.mirror`、`structure.chapter_break`/`chapter_level`、`[[headings]] level = 4`，新增 `profiles/thesis-bit.toml` | 学位论文的刚性要求：装订线要留边、镜像边距让装订边总在内侧、每章另起右页、四级标题要真样式。`UNSUPPORTED` 随之删掉 gutter/mirror/hanging 三条，避免"schema 有、脚本不认"的死字段 |
| 18 | 章前分节与页码（2026-11-15） | 只在第二个及以后的章标题前分节；分节后显式删除克隆来的 `w:pgNumType` | `doc.add_section()` 会复制上一节的 sectPr，不清掉就变成"每章从 1 重编"。实测发现并修复，写进 02 与 06 防回归 |
| 19 | 镜像边距的空白页（2026-11-15） | 不改实现，改为提示 + 文档化（K14） | 有 `w:mirrorMargins` 时 Word 让新节从右页开始，正文前必然多一张不带页码的空白页；这是装订成册的正确版式，为了少一页关掉镜像反而会让奇偶页装订边错位 |
| 20 | 公文网格与行高（2026-11-15） | `body.line_spacing.mode = "grid"` 同时写 `w:docGrid`（字符格）与固定行距 = 版心高 ÷ 行数；段前段后显式为 0 | Word 实测 `w:docGrid` 只决定字符格宽度、不控制行高：只写网格时 16pt 仿宋行高 43 磅、每页 17 行；写固定 29 磅后才 22 行。另外 python-docx 默认模板 `docDefaults` 带 `w:after="200"`（10 磅），不清零会把行高顶成 39 磅——两点都是实测出来的，写进 02 §2.3 与 05 第九节 |
| 21 | 全线表（2026-11-15） | `table.style = "grid"` + `table.border.grid`，六条边全写 single；`check_docx.py` 按 style 分支校验 | 公文表格四面封口，三线表不成立。线宽取 0.5 磅（`sz=4`）属常见做法，标 `verified = false`；三线表分支保持原样，两种表混排由 profile 决定 |
| 22 | 奇偶页外侧页码（2026-11-15） | `page_number.position = "odd-even-outside"` + `format = "emdash"`，写 `section.footer`/`even_page_footer` 两份 PAGE 域并置 `w:evenAndOddHeaders` | 公文页码在外侧、奇偶页分置，且格式是「— 1 —」不是「- 1 -」。实现上两份页脚都要有域，否则偶数页没有页码；校验器同时核对 `evenAndOddHeaders` 与两份页脚 |
| 23 | 模板接管（2026-11-15） | 新增 `scripts/fill_template.py`：读入既有 .docx，替换 `{{占位符}}`，把 md 正文注入 `{{正文}}` 锚点；版式（页面/页眉页脚/页码/封面/目录）全部由模板负责，脚本只管内容排版；产物写 `wdp_mode=template`，`check_docx.py` 见标记即跳过版式校验 | 有些文档类型没有公开排版规范，版式就是单位给的那份 .docx（实验报告、课程设计、NSFC 申请书、各类申请表）。按 profile 重排会和单位模板打架，正确做法是「不碰模板版式、只填内容」。两条实现教训：①lxml 元素代理对象不稳定，不能用 `id(el)` 判断哪些节点是新增的，必须先快照 body 子节点再截断；②python-docx 的新段落插在 `sectPr` 之前，新增元素起点是快照长度 − 1，不是快照长度——两个错法都表现为「正文跑到模板末尾」 |

## 二、由此确定的技术形状

```
profiles/
  ug-thesis-bit.toml     # 现有基线，默认 profile（A4 / 2.54-3.18 / 22pt / 三线表 / - 1 -）
  tech-doc.toml          # GB/T 8567 软件文档惯例（1.5 倍行距 / 五号黑体图表名 / 首页不编）
  thesis-bit.toml        # 学位论文（装订线 1.0cm / 镜像边距 / 章前奇数页分节 / 四级标题 / 罗马前置页码）
scripts/
  profile.py             # TOML 加载 + base 继承 + schema 校验 + 未实现字段报错
  md2docx.py             # 读 profile 渲染；--profile 名或路径
  check_docx.py          # 读同一 profile 比对期望值
  check_md.py            # 语法与图注表注检查（与 profile 基本无关，仅图表编号策略相关）
```

## 三、分期

- **第一期（已完成）**：`profile.py` + 两份 profile + md2docx/check_docx 读配置 + 目录域与分节页码 + 文档同步。
- **第二期（已完成）**：学位论文——`thesis-bit` profile、前置罗马页码、正文从 1 起、每章另起右页（`chapter_break = "odd_page"`）、装订线与镜像边距、四级标题；同一轮把校验闸门收紧并引入生成标记（决策 16–19）。
- **第三期（已完成）**：公文——`cn-gov` profile、22×28 网格、仿宋_GB2312/小标宋 fallback、全线表、`一、（一）` 编号检查、奇偶页外侧「— 1 —」页码（决策 20–22）。首页上边距特例（12.6cm，需首页单独分节）未做，写入限制。悬挂缩进已随第二期用 `first_line_indent.mode = "hanging"` 解锁，尚缺一份使用它的 profile。
- **第四期（已完成）**：模板接管——`fill_template.py`、`{{占位符}}` 与 `{{正文}}` 锚点、`wdp_mode=template` 与校验豁免、`references/07-模板接管.md`（决策 23）。

## 四、仍待处理

- 文档第 9 节的待核实项：按决策 10 处理，不阻塞第一期。`tech-doc` 的六处惯例值已于 2026-11-15 核查完毕（见决策 15）：无该类型的标准依据，保持 `verified = false`，交付给具体单位时按对方模板覆盖。`thesis-bit` 的装订线 1.0cm、四级标题字号同样标 `verified = false`，未逐校核对。
- 环境侧仍未验证（2026-11-15 实测记录见 `references/05-工具链与环境.md` 第九节）：本机没有 WPS 与 LibreOffice，无法验证 WPS 的域/三线表渲染，也没有 Word 2016/2019 可测；无宋体/黑体环境下的版式同理。
- 模板接管只支持「无条件的占位符替换 + 一处正文锚点」：没有循环行（一个变量对应多行表格）、不改模板表格行数、不生成页码/目录域、重复运行会在锚点处重复注入。需要「一表多行」时手工在 Word 里补，或先把模板改成一行一条的静态结构。
- 字体 embedding 不在任何一期：python-docx 不支持，需操作 OOXML `w:embedRegular` 与字体混淆，成本远高于收益；缺字体只做 fallback + 告警。
