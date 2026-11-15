# word-doc-pipeline

把中文文档写成 Markdown，再用脚本确定性地产出符合排版规范的 Word（.docx）：A4、页眉页码、宋体/黑体 + Times New Roman、固定或倍数行距、首行缩进、三级标题、三线表、图注表注按章编号。同时提供写作约束检查与成品格式校验，转换器和校验器读同一份配置，规范与脚本不会脱节。

适合：课程设计说明书、可行性分析报告、需求/设计/测试文档、规范文档、技术文档——凡是要反复改正文、又不想每次在 Word 里重排格式的场景。

## 它解决什么问题

在 Word 里手改正文，格式会随内容漂移；改完一处又要通篇核对字号行距。这里的做法是：

- **`.md` 是唯一真源**：正文、表格、图注、表注只写 Markdown，`.docx` 与 PDF 都是派生物，永不反向修改。
- **排版参数只有一处**：`profiles/*.toml`。换文档类型 = 换一份 profile，不碰脚本。
- **生成与校验同源**：`md2docx.py` 按 profile 渲染，`check_docx.py` 按同一份 profile 反查成品，改规范不会漏掉校验。
- **脚本不做的三件事留给 Word**：补封面、更新目录域、另存 PDF——这三步依赖版式判断和最终渲染，硬做反而不可靠。

## 快速开始

```powershell
# 1. 依赖（Python 3.11+，需要 tomllib；3.12 自带）
python -m pip install python-docx

# 2. 选 profile（默认 ug-thesis-bit）
python scripts/profile.py --list

# 3. 检查 Markdown 是否符合转换子集
python scripts/check_md.py docs --profile tech-doc

# 4. 转换
python scripts/md2docx.py docs/需求文档/需求文档.md docs/需求文档/需求文档.docx --profile tech-doc

# 5. 校验成品（profile 必须与转换时一致，否则报的差异是假的）
python scripts/check_docx.py docs/需求文档/需求文档.docx --profile tech-doc --project "项目名"
```

补封面、`Ctrl+A` + `F9` 更新目录与页码域、另存 PDF，然后按 [06-验收检查清单.md](references/06-验收检查清单.md) 逐项勾选。

## 目录

| 路径 | 内容 |
| --- | --- |
| `scripts/` | 五个脚本：`profile.py` 配置加载与校验、`md2docx.py` 转换、`check_md.py` 源文件检查、`check_docx.py` 成品校验、`build_docs.py` 批量转换 + 校验 |
| `profiles/` | `ug-thesis-bit.toml`（默认基线）、`tech-doc.toml`（`base` 继承范例）、`thesis-bit.toml`（装订线 + 镜像边距 + 章前分节 + 四级标题） |
| `references/` | 六份规范：文档工程与目录约定、Word 格式、Markdown 写作约束、图表与表格、工具链与环境、验收清单 |
| `docs/` | 调研与决策记录（文档类型规范总览、profile 决策） |
| `assets/templates/` | 文档骨架、图表绘制说明模板、`gitignore` 追加片段 |
| `examples/` | 可转换的最小样例，验证脚本是否正常 |
| [SKILL.md](SKILL.md) | 完整使用手册：接入步骤、标准流程、命令速查、已知坑 |

## 命令

```
profile.py    --list | --show <名> | --check <名>
md2docx.py    <in.md> <out.docx> [--profile 名] [--toc|--no-toc] [--project "名称"]
                                 [--no-header] [--no-footer] [--page-size A4|Letter]
check_md.py   <md 或目录>... [--profile 名] [--all]
check_docx.py <docx 或目录>... [--profile 名] [--project "名称"]
build_docs.py <docs 目录> [--profile 名] [--toc] [--project] [--include 模式]
                          [--skip-suffix ...] [--dry-run] [--no-check]
```

`--profile` 省略时用 `ug-thesis-bit`；可用值由 `profiles/*.toml` 文件名决定。`check_md.py` / `check_docx.py` / `build_docs.py` 有错误时退出码为 1，可直接用于 CI。批量转换前先 `build_docs.py --dry-run` 看清单。

## profile 一览

| 项 | `ug-thesis-bit`（默认） | `tech-doc` | `thesis-bit` |
| --- | --- | --- | --- |
| 行距 | 固定 22 磅 | 1.5 倍 | 固定 22 磅（继承） |
| 章 / 节 / 三级标题 | 黑体 16 / 15 / 12，三级才加粗 | 黑体 16 / 14 / 12，均加粗 | 同默认 |
| 图注表注 | 宋体小五 9pt | 黑体五号 10.5pt | 同默认 |
| 首页页码 | 编 | 不编 | 不编（封面页） |
| 目录前置节 | 无 | 可选 | 有，罗马数字页码 |
| 装订线 / 镜像边距 | 无 | 无 | 1.0 cm / 开 |
| 章前分节 | 无 | 无 | 每章从奇数页开始 |
| 四级标题 | 加粗段落 | 加粗段落 | `####` 生成真 Heading 4 |
| 依据 | 北理工系毕业设计说明书规范 | GB/T 8567 只规定结构，排版为通行的惯例值 | 常见校级论文模板，未逐校核实 |

`tech-doc` 与 `thesis-bit` 都用 `base = "ug-thesis-bit"` 只覆盖差异项——要给某个学校或单位做版式，复制一份改十几行即可。惯例值一律标 `verified = false`，运行时提示、校验时降为警告，不会被当成权威数值。

## 已知限制

- 只支持 Markdown 子集：`#`/`##`/`###`（profile 定义了 level 4 时还有 `####`）、管道表格、`-`/`1.` 列表、`**加粗**`、`---`、`![](images/...)`、表格上方的 `**表 x-x 说明**`。代码块、行内链接、合并单元格、单元格内换行都不支持（详见 [03-Markdown写作约束.md](references/03-Markdown写作约束.md)）。
- 表格每行单元格数必须与表头一致，多于表头会直接中断转换。
- 图片缺失只打印警告并跳过，图注一起丢失，转换不会失败——必须看 stderr。
- 固定行距会裁切行内图片，图片段落已自动改为倍数行距，不要手动改回。
- 目录与页码是 Word 域，首次打开是空的，需要更新域；重新生成 `.docx` 后，封面、目录、PDF 全部要重做。
- 字体只写名字不嵌入，缺字体按 `fonts.fallback` 链替换并告警；正式交付应在装有目标字体的机器上复核。

## 分期

## 分期

第一期（profile 化、目录域与分节页码）与第二期（学位论文：装订线、镜像边距、章前分节、四级标题，以及校验器闸门收紧与生成标记）已完成；第三期公文（22×28 网格、仿宋_GB2312 字体回退、全线表、奇偶页外侧页码）、第四期模板接管尚未开始，取舍理由见 [docs/profile-decisions.md](docs/profile-decisions.md)。

## 许可

MIT，见 [LICENSE](LICENSE)。
