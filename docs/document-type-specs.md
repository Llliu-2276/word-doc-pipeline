# Word 文档类型格式规范总览

> 调研日期：2026-11-10。目的：为 `word-doc-pipeline` 从"单一排版基线"扩展为"按文档类型切换 profile 的通用技能"提供依据。
> 方法：5 路并行检索，优先取标准全文、高校教务处/研究生院官网原文、部委公开文件；所有数值标注来源，无法取到原文的一律列入待核实，不凭记忆填充。

## 0. 一句话结论

现有基线（A4 / 宋体小四 12pt / 固定 22 磅 / 三线表 / 页脚 `- 1 -`）不是"通用规范"，而是**北理工系毕业设计说明书**的一个实例被当成了默认值：其中 22 磅行距、三线表线宽都有真实依据，但左右 3.18cm 只是 Word「普通」页边距默认值，`- 1 -` 页码格式在 8 所高校规范里一次都没出现。通用化的做法不是推翻它，而是把它降级为 `profile: ug-thesis-bit`，再补齐缺失的开关。

## 1. 规范强度分层

不同文档类型的"规范"根本不是同一层东西，profile 设计必须先承认这一点：

| 层级 | 含义 | 例子 |
| --- | --- | --- |
| L1 强制性国标 | 有法可依、违反即不合规 | GB 5296.1—2012（消费品说明书，但只约束内容与可读性，不管字号） |
| L2 推荐性国标（实务强制） | 推荐性，但行业内按强制执行 | GB/T 9704—2012（党政机关公文格式）、GB/T 1.1—2020（标准编写）、GB/T 8567—2006（软件文档）、GB/T 7713 系列（学位论文/学术论文） |
| L3 机构模板 | 学校/期刊/企业模板，文本规范常与模板不一致 | 清华、上交、北航、北理工、东南等学位论文规范；期刊投稿指南；企业需求说明书模板 |
| L4 纯惯例 | 无任何文本依据，只有行业习惯 | 商业计划书、简历、总结/述职、调研报告 |
| L5 无排版规范（模板承载） | 规范只规定内容要素，格式由统一印制的表格/模板承载 | 实验报告、课程设计说明书、NSFC 申请书（系统模板） |

**设计含义**：L1/L2 的 profile 必须与标准条款一一对应（可校验）；L3 要允许"按校/按刊"派生；L4 给一个合理默认即可；L5 需要的不是排版生成，而是**模板接管模式**（在既有 .docx 里填内容），是另一条技术路径。

## 2. 跨类型参数对照矩阵

| 类型 | 页边距（上/下/左/右） | 装订线 | 正文 | 行距 | 标题体系 | 图表注 | 页码 | 表格 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 党政机关公文 GB/T 9704—2012 | 37/35/28/26 mm，版心 156×225 | 左侧装订、双面 | 3 号仿宋_GB2312（16pt） | **22 行×28 字撑满版心**（≈28.8–29.3 磅，国标无磅值） | 标题 2 号小标宋；层次 一、(黑体) （一）(楷体) 1./（1）(仿宋)，均 3 号 | 无图注规定 | 4 号宋体 `— 1 —`，外侧，奇偶分左右；无页眉 | 全线表 |
| 学位论文·清华 | 3.0/3.0/3.0/3.0 | 0 | 小四宋体 + TNR | 固定 20 磅 | 三号/四号/13pt/小四 黑体**不加粗**，前 24 后 18 | 图题 11pt，前 6 后 12 | 摘要起阿拉伯连续，TNR 五号居中 | 三线表 1 磅/0.5 磅 |
| 学位论文·上交 | 3.5/4.0/2.8/2.8 | 0.5 | 小四或五号 | 1.25 倍 | 三号/四号/小四 **粗黑体** | — | 待核实 | 三线表 |
| 学位论文·北航 | 2.5×4 | 0 | 小四 | 1.5 倍 | 章三号黑体；关键词黑体小四 | — | 前置 Ⅰ Ⅱ Ⅲ，正文 1 起 | 三线表 |
| 期刊·计算机学报 | 2.0×4 | — | 五号宋体 | — | 按刊 | 六号 | — | 三线表；双栏 |
| 毕业设计·北理工 | 3.5/2.6/3.0/2.6 | 0 | 宋体小四 | **固定 22 磅** | 三号/四号/小四 加粗 | 五号宋体 | 页脚五号宋体居中 | 三线表；字距 +0.5 磅 |
| 毕业设计·东南 | 2/2/2/2 | 0.5 | 宋体小四 | 18 磅 | — | — | 页眉右「第X页」 | — |
| 毕业设计·哈工程 | 2.8/2.8/2.5/2.5 | 0 | 宋体小四 | 22 磅固定 | 章小二黑体，段前分页 | — | — | — |
| 毕业设计·中北 | 3.0/2.5/3.0/2.0 | 0 | 宋体小四 | 1.5 倍 | 三号/四号/小四（三级不加粗） | 五号宋体 | 页脚居中 | 三线表「尽量」 |
| 毕业设计·深大 | 待核实 | — | 宋体**五号** | 单倍 | — | 小五 | — | — |
| 毕业设计·西邮 | 2.5×4 | 0.5 | 宋体小四 | 1.25 倍 | 三号/四号/小四 | 五号 | 页眉（双线） | — |
| 软件技术文档（GB/T 8567） | 无国标，惯例 2.54/3.18 | — | 小四宋体 | 1.5 倍 | 章三号黑体/节四号黑体 | 图名表名五号黑体 | 首页不编 | 三线表 |
| 标准文本（GB/T 1.1—2020） | 无公开数值 | — | **五号宋体（10.5pt）** | — | **章条标题五号黑体（与正文同字号）** | 图表题五号黑体；表内/注小五；图中字六号 | — | **封闭式表格** |
| 合同 | 惯例 2.54/3.17 | — | 四号仿宋（14pt） | 1.5 倍 | 「第一条」全文连续编号 | 表题在上 | 第X页 共Y页 | 全线表 |
| 可研报告 | 惯例 2.54/3.17 | — | 小四宋体 | 1.5 倍 | 第一章/1.1/1.1.1（发改委大纲） | — | 封面不编 | 全线表 |
| NSFC 申请书 | 系统模板 | — | 12pt，全文一致 | — | 立项依据/研究内容/研究基础三部分，≤30 页 | — | — | — |
| 实验报告 | **无排版规范** | — | 统一印制报告纸/模板 | — | 10 项固定结构 | — | — | — |
| 商业计划书/简历/总结 | 无 | — | 自由 | — | 无 | — | — | — |

## 3. 与现有基线的冲突清单

### P0：现在做不出来，且是硬需求

1. **单 section** —— 学位论文「前置罗马页码 + 正文从 1 起 + 每章另起右页」、公文「首页上边距 12.6/7.1cm」都无法表达。
2. **行距硬编码 22 磅** —— 无法覆盖固定 20/18 磅、1.5/1.3/1.25 倍、单倍、公文网格；22 磅撑在公文里会让页数与版心全部错位。
3. **页边距硬编码 2.54/3.18** —— 无装订线、无镜像（双面）、无不对称四边值。
4. **只有三线表** —— 公文、标准文本、合同需要全线表/封闭式表格。
5. **无封面、目录、修订记录、签署页、文档控制页** —— 学位论文、合同、软件文档、可研直接不可用。
6. **字号单档 12pt** —— 公文 16pt、合同 14pt、证券/期刊 10.5pt、标准小五 9pt 等无法切换。
7. **标题层级固定 16/15/12 且三级才加粗** —— 需每级独立 `font/size/bold/对齐/段距`；GB/T 1.1 甚至是「标题与正文同字号、只换字体」，现有模型无法表达。

### P1：能凑合但会产出不合规结果

8. 图注 9pt —— 北理工/中北/西邮/东华/内蒙工大都是五号 10.5pt，基线小一级。
9. 页码格式唯一 `- 1 -` —— 需支持纯数字、第X页共Y页、`— 1 —` 外侧奇偶、罗马数字前置、自定义起始页。
10. 编号体系固定 `1/1.1/1.1.1` —— 公文需 `一、（一）1.（1）`（最深 4 层，禁「第一章」），合同需「第一条」连续编号。
11. 图表编号「按章」是写死的 —— 需 `by_chapter` / `continuous` 二选一（公文与 GB/T 1.1 用全文顺序）。
12. 首行缩进已按字符计（正确），但缺「回行顶格」与参考文献「悬挂缩进 2 字符」两种模式。

### P2：体验与正确性

13. 无字体 fallback 与缺字体告警（公文必需 方正小标宋简体、仿宋_GB2312，本机均缺）。
14. 无字距加宽（北理工要求 +0.5 磅）。
15. 页眉横线（单线 0.5mm / 双线）无法配置。
16. 无「模板接管」模式（实验报告、部分课程设计的唯一正解）。

## 4. Profile 抽象设计

单一 YAML/JSON 配置档驱动，脚本只读 profile 不写死数值：

```yaml
name: cn-thesis-thu
standards: [GB/T 7713.1-2025, GB/T 7714-2025]
page:
  size: A4
  margins: {top: 3.0, bottom: 3.0, left: 3.0, right: 3.0}   # cm
  gutter: 0
  mirror: false
  header_distance: 2.2
  footer_distance: 2.2
  different_first_page: false
body:
  font_cn: 宋体
  font_en: Times New Roman
  size: 12
  line_spacing: {mode: fixed, value: 20}    # fixed | multiple | grid
  grid: {lines: null, chars: null}          # 公文用 {lines: 22, chars: 28}
  first_line_indent: {chars: 2, mode: first_line}  # first_line | hanging | none
  align: justify
  space_before: 0
  space_after: 0
  char_spacing: 0                           # 北理工 +0.5 磅
headings:
  - {level: 1, font: 黑体, size: 16, bold: false, align: center,
     space_before: {pt: 24}, space_after: {pt: 18}, page_break_before: false}
  - {level: 2, font: 黑体, size: 14, bold: false, align: left, ...}
numbering:
  scheme: decimal-three      # decimal-three | cn-gov(一、（一）1.（1）) | article(第一条) | chapter(第一章)
  max_depth: 3
  figure: {scheme: by_chapter, sep: "-", caption: below}
  table:  {scheme: by_chapter, sep: "-", caption: above}
table:
  style: three_line          # three_line | grid | none
  border: {outer: 12, header: 6}   # 0.125mm 单位（sz），12=1.5磅
  cell_font: 宋体
  cell_size: 10.5
  header_repeat: true
  row_height: null
figure:
  caption_size: 11
  caption_space_before: {pt: 6}
  caption_space_after: {pt: 12}
  max_width: 14.5            # cm
header_footer:
  header: {text: null, rule: none, font: 宋体, size: 9}
  page_number:
    position: footer-center   # footer-center | header-right | odd-even-outside
    format: plain             # plain | dash | cn(第X页共Y页) | roman
    font: Times New Roman
    size: 9
    start_at: 1
    restart_per_section: false
structure: [cover, abstract_cn, abstract_en, toc, body, references, appendix, acknowledgment]
special:
  signature_page: false
  seal_space: null
  amount_in_words: false
  doc_control_page: false
  revision_history: false
  classification: null
fonts:
  fallback:
    "仿宋_GB2312": [仿宋, FangSong, 宋体]
    "方正小标宋简体": [华文中宋, STZhongsong, 宋体]
    "楷体_GB2312": [楷体, KaiTi, 宋体]
```

## 5. 建议的 profile 清单

| profile | 覆盖 | 规范强度 |
| --- | --- | --- |
| `cn-gov` | 党政机关公文 15 文种 | L2 |
| `standard-gb` | 国家标准/团体标准文本（GB/T 1.1） | L2 |
| `tech-doc` | 软件需求/设计/测试/手册（GB/T 8567） | L2 |
| `thesis-*`（清华/上交/北航/上外/北工大/华科） | 学位论文 | L3 |
| `journal-*` | 期刊投稿 | L3 |
| `ug-thesis-*`（北理工/东南/哈工程/中北/深大/西邮） | 毕业设计说明书 | L3 |
| `contract` | 合同/协议 | L4 |
| `fs-report` | 可行性研究报告 | L4 |
| `proposal-nsfc` | 科研项目申请书 | L5（模板接管） |
| `lab-report` | 实验报告/课程设计 | L5（模板接管） |
| `business-doc` | 商业计划书、总结、述职、调研 | L4 |
| `resume` | 简历 | L4 |

## 6. 实现可行性分档

**A. 已实测可用（2026-11-10 探针验证）**

| 能力 | 验证结果 |
| --- | --- |
| 多 section + 另起右页 | `add_section(WD_SECTION.ODD_PAGE)` → `w:type=oddPage` ✓ |
| 页码重启与格式 | `w:pgNumType w:start="1" w:fmt="upperRoman"/"decimal"` ✓ |
| 镜像边距 + 装订线 | `w:mirrorMargins` + `w:gutter` ✓ |
| 首页不同 | `different_first_page_header_footer` ✓ |
| 目录域 | `TOC \o "1-3" \h \z \u` 写入成功 ✓ |
| 标题可被目录抓取 | `md2docx.py:118` 用 `doc.add_heading(level=...)`，即真 Heading 样式，仅被直接格式覆盖 ✓ |

结论：**分节与页码不是技术障碍**，缺的全是配置层与结构件生成。

**B. 需新增代码**：封面/结构件生成器、表样式切换（三线/全线）、编号体系映射、页码位置与奇偶页脚、行距三模式与文档网格、逐级标题表、悬挂缩进、图表编号策略、控制页与修订记录表、缺字体检测。

**C. 外部条件**：公文排版依赖 方正小标宋简体、仿宋_GB2312、楷体_GB2312，本机均未安装（python-docx 只写字体名不嵌字体，缺失时 Word 回退宋体）。需 fallback 链 + 生成时告警，或由用户装字体。

**D. 结构性限制**：实验报告、课程设计说明书、NSFC 申请书属"模板承载"型，正解是**模板接管**（读入既有 .docx，填占位符，保留原样式），不是从零生成——这是与其余 profile 不同的技术路径，需单独设计。

## 7. 时效提醒（会改变默认值）

- **GB/T 7713.1—2025** 学位论文编写规则：2025-08-01 发布、**2026-02-01 实施**，取代 2006 版（新增电子学位论文、原创性声明、使用授权说明，改版面与书脊）。
- **GB/T 7714—2025** 参考文献著录：**2026-07-01 起实施**，全部代替 2015 版。学位论文条目由"页码任选"改为**必标引文页码**；西文作者姓仅首字母大写（`Smith J`，不再是 `SMITH J`）；新增预印本 `[PP]`，`[EB]` 改指网页。多数高校模板仍写 2015 版，需配置开关。
- 公文：openstd 仅列 GB/T 9704—2012 且状态"现行"，未查到 2024/2025 修订立项公告——属"未见公开信息"，不等于确认无修订。

## 8. 调研内部的分歧点（已核对）

1. **三线表线宽 1.5/0.75 磅**：课程设计一路称只可溯源到电子科技大学研究生规范；技术文档一路核实 GB/T 7713.1—2006 明确"建议采用三线表，顶线底线 1.5 磅、栏目线 0.75 磅"。**以后者为准**，本科校级文件只是未写线宽。
2. **22 磅行距**：并非凭空而来，北理工、哈工程正文即 22 磅，西北农林写"最小值 22 磅"；但清华 20 磅、东南 18 磅、多校 1.5 倍，所以它是"某个流派"的值，不是通用值。
3. **左右 3.18cm**：即 Word「普通」页边距默认值，所查 8 校无一采用。

## 9. 待核实项汇总

- 公文：行距实测磅值（28/28.5/28.8/29.3/30 并存，无国标数值）；发文机关标志字号；份号字体；密级与期限间"★"；图/公式编排。
- 学位论文：清华三级标题 13pt 是否为排版误差；清华表题字号；各校页眉横线；上交页码罗马规则；华科页边距；GB/T 7713.1—2025 第 6.10 条是否规定具体字体字号（全文受版权保护）。
- 课程设计：课程设计说明书是否有给字号的校级文件；实验报告字体行距（由模板承载）；图表最大宽度条款；纯高职院校样本。
- 技术文档：GB/T 8567—2006 附录 DID 逐条条号（2006 版全文未取得，现用 1988 版结构）；GB/T 1.1—2020 表 F.1 逐项值（现引 2009 版附录 J）；GB/T 1.1 页边距/页码条款。
- 商务行政：GB 5296.1—2012 第 8.2.1 条文字高度点值表；GB/T 9969 现行版本；合同示范文本是否预置字体；各省申报书字体要求。

## 10. 来源

**标准**
- [GB/T 9704—2012 全文（上海财经大学转载）](https://jwc.sufe.edu.cn/f6/8c/c8107a128652/page.htm)｜[国家标准全文公开](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=F3CC9BEF482524C895FDA7A08BB4A70E)
- [GB/T 7713.1—2025 标准信息](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=36C05B5738C54D42B4B262525320B52F)｜[全文（上海海事大学）](https://gs.shmtu.edu.cn/_upload/article/files/54/63/492f2d84404b959505f3fc8299a2/509f006c-7625-4bca-bfe5-b1ce96b9f65d.pdf)
- [GB/T 7713.2—2022](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=0B963916637B8F34B295FCF4A51A1BE5)｜[GB/T 7713.3—2014](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=E2BA32B0F5E570E5844A9E7E2D42393F)
- [GB/T 8567—2006](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=84C42B6277D2714B7176B10C6E6B1A44)｜[GB 8567—1988 第 7–14 章结构](https://m.bzxz.net/bzxz/11660.html)
- [GB/T 9385—2008 全文 PDF](https://29229836.s21i.faiusr.com/61/ABUIABA9GAAgmYaq0QYousnYiwY.pdf)｜[GB/T 1.1—2009 附录 J 字号表](https://www.cmde.org.cn/gbpdf/GB1.1-2009.pdf)｜[GB/T 1.1—2020](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C4BFD981E993C417EF475F2A19B681F1)
- [GB 5296.1—2012](https://law.resource.org/pub/cn/ibr/gb.5296.1.c.2012.pdf)｜[GB/T 19678.1—2018](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=90696C0D60B8DCAE22538B689B7DEDC2)
- [GB/T 7714—2025 实施公告](http://tsg.xaau.edu.cn/info/1113/6081.htm)｜[《民法典》第 470 条](https://www.moj.gov.cn/pub/sfbgw/zwgkztzl/2025nianzhuanti/2025mfdxcy/2025mfdxcy_mfdql/202505/t20250507_518708.html)

**高校规范**
- 清华 [写作指南](https://www.dhs.tsinghua.edu.cn/wp-content/uploads/2023/12/2024031107044595.pdf)｜上交 [撰写指南](https://www.gs.sjtu.edu.cn/post/detail/Z3MzNDA0)｜北航 [规范](https://graduate.buaa.edu.cn/__local/4/27/6F/4E099CEBEC1B251AAE2796893F6_8A8AEEED_A4384.pdf)｜上外 [规范](https://graduate.shisu.edu.cn/_upload/article/files/d9/c2/e87f444e46b483a90d76f53c9523/91fe814e-e45f-4e1a-89dc-5130dfc2f1a1.pdf)｜北工大 [规范](https://graduate.bjut.edu.cn/beijinggongyedaxueyanjiushengxueweilunwenzhuanxieguifan.pdf)｜华科 [规定](https://mse.hust.edu.cn/info/1087/10059.htm)
- 北理工 [毕设规范](https://jwb.bit.edu.cn/tzgg/bf894804c42f42de93291543e6967800.htm)｜东南 [PDF](https://law.seu.edu.cn/_upload/article/files/43/39/e3cda4e443d7881c4f50e968361a/e29d9a40-ec35-4c46-8ac1-4617c783a839.pdf)｜哈工程 [DOC](https://icec.hrbeu.edu.cn/_upload/article/files/51/17/46871c814982be83a744f019022a/7ea17a95-0e06-459d-a369-21c8826f4485.doc)｜中北 [规范](https://jwc.nuc.edu.cn/info/1005/3354.htm)｜深大 [规范](https://jwb.szu.edu.cn/info/1358/2010.htm)｜西邮 [规范](https://jzxy.xupt.edu.cn/info/1094/1948.htm)｜广工商 [规范](https://jwc.gdbtu.edu.cn/contents/1168/33505.html)｜电子科大 [研究生规范](https://gr.uestc.edu.cn/xiazai/114/3917)

**实务与行业**
- [公文格式实施细则（行距 29.3 磅实例）](https://bgsh.glut.edu.cn/info/1027/1551.htm)｜[许昌学院（30 磅）](https://dwbgs.xcu.edu.cn/info/1193/1040.htm)｜[应急管理部报告格式](https://www.mem.gov.cn/gk/zfxxgkpt/fdzdgknr/202610/W020261016539448169956.doc)
- [发改委可研报告编写大纲 PDF](https://www.ndrc.gov.cn/xxgk/zcfb/ghxwj/202304/P020230407401908613786.pdf)｜[NSFC 2026 申请书改版](https://www.nsfc.gov.cn/p1/3381/2821/99242.html)｜[合同示范文本库](https://htsfwb.samr.gov.cn/View?id=2340996b-882d-47a4-b74d-c30784628737)
- [《计算机学报》格式](https://hhxueshu.com/h-nd-3975.html)｜[《中国人民大学学报》投稿指南](http://xuebao.ruc.edu.cn/CN/column/item19.shtml)
