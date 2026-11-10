# -*- coding: utf-8 -*-
"""中文技术/课程设计文档 Markdown → Word (.docx) 转换器。

用法:
    python scripts/md2docx.py <输入.md> <输出.docx> [--project "页眉项目名"] [--no-header] [--no-footer] [--page-size A4]

排版基线（与本 Skill 的 references/02-Word格式规范.md 一致）:
- A4（21.0×29.7cm），上下边距 2.54cm，左右 3.18cm
- 正文宋体小四（12pt），西文/数字 Times New Roman，固定 22 磅行距，首行缩进 2 字符
- 一级标题黑体三号（16pt）、二级黑体小三（15pt）、三级黑体小四（12pt）加粗，段前段后 0.5 行
- 图注/表注/页眉/页码宋体小五（9pt）
- 图片段落 1.5 倍行距（固定行距会裁切行内图片），图注在图片下方居中
- 三线表：上下框线 1.5 磅、表头下框线 0.75 磅、无竖线与内横线、表头加粗并跨页重复
- 页眉可选（项目名），页脚页码居中，格式「- 1 -」

支持的 Markdown 子集:
- # / ## / ### 标题（#### 不支持，会退化为加粗段落并告警）
- | a | b | 表格（首行表头，第二行 |---| 分隔）
- - 无序列表、1. 有序列表
- **加粗** 行内文本
- --- 分隔线（忽略）
- ![图注文字](图片相对路径) 嵌入图片（段落 1.5 倍行距）
- 表格正上方一行的 **表 x-x 说明** 渲染为表注（允许与表格之间有空行）
- 封面与目录不经过本脚本：docx 封面在 Word 中手工排版，目录在 Word 中插入并更新域
"""
from __future__ import annotations

import argparse
import os
import re
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# 字体常量
FONT_HEI = "黑体"          # 标题
FONT_SONG = "宋体"         # 正文/注释
FONT_LATIN = "Times New Roman"  # 西文与数字

# 字号（磅）
SIZE_H1 = 16    # 三号
SIZE_H2 = 15    # 小三
SIZE_H3 = 12    # 小四
SIZE_BODY = 12  # 小四
SIZE_CAPTION = 9  # 小五
SIZE_CELL = 10    # 表格内文字

LINE_SPACING = Pt(22)      # 固定 22 磅行距
IMG_LINE_SPACING = 1.5     # 图片段落 1.5 倍行距（固定值会裁切行内图片）
MAX_IMG_WIDTH = Cm(14.5)   # 图片最大宽度（A4 减左右 3.18cm 边距）

PAGE_SIZES = {
    "A4": (Cm(21.0), Cm(29.7)),
    "Letter": (Cm(21.59), Cm(27.94)),
}

# 已告警的语法类型，避免同一文件逐行刷屏
_WARNED = set()


def warn(msg):
    """向 stderr 输出告警（同一类型只提示一次由调用方控制）。"""
    print(f"警告: {msg}", file=sys.stderr)


def set_run_font(run, bold=False, size=SIZE_BODY, font=FONT_SONG):
    """设置 run 的中西文字体与字号。"""
    run.font.name = FONT_LATIN
    run.font.bold = bold
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font)


def set_para_format(p, indent=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                    size=SIZE_BODY, font=FONT_SONG, bold_base=False,
                    before_lines=None, after_lines=None,
                    line_spacing=LINE_SPACING):
    """设置段落格式：行距、缩进、对齐、段前段后（单位 0.01 行）。"""
    pf = p.paragraph_format
    pf.alignment = align
    pf.line_spacing = line_spacing
    pPr = p._p.get_or_add_pPr()
    if indent:
        ind = pPr.get_or_add_ind()
        ind.set(qn("w:firstLineChars"), "200")  # 首行缩进 2 字符
    if before_lines is not None or after_lines is not None:
        spacing = pPr.get_or_add_spacing()
        if before_lines is not None:
            spacing.set(qn("w:beforeLines"), str(before_lines))
        if after_lines is not None:
            spacing.set(qn("w:afterLines"), str(after_lines))


def fill_paragraph(p, text, size=SIZE_BODY, font=FONT_SONG, bold_base=False):
    """按 **加粗** 语法拆分文本写入段落。"""
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            set_run_font(p.add_run(part[2:-2]), bold=True, size=size, font=font)
        else:
            set_run_font(p.add_run(part), bold=bold_base, size=size, font=font)


def add_heading(doc, text, level):
    """黑体标题：一级三号、二级小三、三级小四加粗；段前段后 0.5 行。"""
    h = doc.add_heading(level=level)
    size = {1: SIZE_H1, 2: SIZE_H2, 3: SIZE_H3}[level]
    bold = level == 3
    fill_paragraph(h, text, size=size, font=FONT_HEI, bold_base=False)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.bold = bold
        set_run_font(run, bold=bold, size=size, font=FONT_HEI)
    set_para_format(h, indent=False, align=WD_ALIGN_PARAGRAPH.LEFT,
                    before_lines=50, after_lines=50)
    return h


def add_caption(doc, text, above_table=False):
    """图注/表注：宋体小五、居中、无缩进。"""
    p = doc.add_paragraph()
    set_para_format(p, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                    size=SIZE_CAPTION)
    fill_paragraph(p, text, size=SIZE_CAPTION, font=FONT_SONG)
    return p


def set_table_borders(table):
    """三线表：上下框线 1.5 磅，表头下框线 0.75 磅，无竖线与内横线。"""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "bottom"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "12")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    for edge in ("left", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def set_header_bottom_border(row):
    """表头行单元格下框线 0.75 磅。"""
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "auto")
        tcBorders.append(bottom)
        tcPr.append(tcBorders)


def set_repeat_header(row):
    """表格首行设为标题行，跨页自动重复。"""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


def add_table(doc, rows):
    """三线表：首行表头加粗、跨页重复；单元格文字宋体 10pt 不缩进。"""
    data_rows = [rows[0]] + rows[2:]  # 跳过 |---| 分隔行
    table = doc.add_table(rows=len(data_rows), cols=len(rows[0]))
    table.autofit = True
    set_table_borders(table)
    for r, cells in enumerate(data_rows):
        for c, text in enumerate(cells):
            cell = table.cell(r, c)
            p = cell.paragraphs[0]
            for old_run in list(p.runs):  # 清空默认空 run，避免残留空文本
                old_run._element.getparent().remove(old_run._element)
            p.paragraph_format.line_spacing = 1.0
            parts = re.split(r"(\*\*.+?\*\*)", text)
            for part in parts:
                if not part:
                    continue
                if part.startswith("**") and part.endswith("**"):
                    set_run_font(p.add_run(part[2:-2]), bold=True, size=SIZE_CELL)
                else:
                    set_run_font(p.add_run(part), bold=(r == 0), size=SIZE_CELL)
    set_repeat_header(table.rows[0])
    set_header_bottom_border(table.rows[0])
    return table


def add_image(doc, alt, path, md_dir):
    """嵌入图片（宽度不超页面），图注置于图片下方居中。"""
    img_path = path if os.path.isabs(path) else os.path.join(md_dir, path)
    if not os.path.isfile(img_path):
        warn(f"图片不存在，已跳过: {img_path}")
        return
    p = doc.add_paragraph()
    set_para_format(p, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                    line_spacing=IMG_LINE_SPACING)
    img = Image.from_file(img_path)
    width = img.width if img.width <= MAX_IMG_WIDTH else MAX_IMG_WIDTH
    p.add_run().add_picture(img_path, width=width)
    if alt:
        add_caption(doc, alt)


def add_page_field(paragraph):
    """在段落中插入 Word PAGE 域（页码），打开文档时自动更新。"""
    run = paragraph.add_run()
    set_run_font(run, size=SIZE_CAPTION, font=FONT_SONG)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, separate, text, end):
        run._r.append(el)


def setup_page(section, project=None, header=True, footer=True, page_size="A4"):
    """设置纸张、页边距、页眉与页脚页码。"""
    width, height = PAGE_SIZES.get(page_size, PAGE_SIZES["A4"])
    section.page_width = width
    section.page_height = height
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

    if header and project:
        hp = section.header.paragraphs[0]
        hp.text = ""
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_format(hp, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                        line_spacing=1.0)
        run = hp.add_run(project)
        set_run_font(run, size=SIZE_CAPTION, font=FONT_SONG)

    if footer:
        fp = section.footer.paragraphs[0]
        fp.text = ""
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_format(fp, indent=False, align=WD_ALIGN_PARAGRAPH.CENTER,
                        line_spacing=1.0)
        set_run_font(fp.add_run("- "), size=SIZE_CAPTION, font=FONT_SONG)
        add_page_field(fp)
        set_run_font(fp.add_run(" -"), size=SIZE_CAPTION, font=FONT_SONG)


def convert(md_path, docx_path, project=None, header=True, footer=True,
            page_size="A4"):
    """将 md 转换为 docx。"""
    with open(md_path, encoding="utf-8") as f:
        lines = f.read().splitlines()

    doc = Document()
    setup_page(doc.sections[0], project=project, header=header,
               footer=footer, page_size=page_size)

    # 文档默认样式：宋体 + Times New Roman，小四，固定 22 磅行距
    normal = doc.styles["Normal"]
    normal.font.name = FONT_LATIN
    normal.font.size = Pt(SIZE_BODY)
    normal.paragraph_format.line_spacing = LINE_SPACING
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), FONT_SONG)

    md_dir = os.path.dirname(os.path.abspath(md_path))
    first_h1 = None
    i = 0
    # 封面由作者在 Word 中手工排版，脚本不生成、不解析封面（md 中不写封面块）。

    while i < len(lines):
        line = lines[i].rstrip()
        s = line.strip()

        if s == "---" or not s:
            i += 1
            continue

        # 代码块：不支持，按普通段落输出并告警
        if s.startswith("```"):
            if "code" not in _WARNED:
                warn("不支持代码块 ```，已按普通段落输出（建议改为文字描述或表格）")
                _WARNED.add("code")
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                add_body(doc, lines[i].strip())
                i += 1
            i += 1
            continue

        # 引用块：不支持，去掉 > 前缀按正文输出
        if s.startswith(">"):
            if "quote" not in _WARNED:
                warn("不支持引用块 >，已去掉前缀按正文输出")
                _WARNED.add("quote")
            add_body(doc, s.lstrip("> ").strip())
            i += 1
            continue

        # 表注：表格上方的 **表 x-x 说明** 加粗行（允许与表格之间夹空行）
        m_cap = re.match(r"^\*\*(表[^*]+)\*\*\s*$", line)
        if m_cap:
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j + 1 < len(lines) and lines[j].strip().startswith("|") \
                    and re.match(r"^\|[\s:\-|]+\|?$", lines[j + 1]):
                add_caption(doc, m_cap.group(1))
                i = j
                continue

        # 表格：连续的以 | 开头的行
        if s.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:\-|]+\|?$", lines[i + 1]):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(cells)
                i += 1
            widths = {len(r) for r in rows}
            if len(widths) > 1:
                warn(f"表格列数不一致（{sorted(widths)}），请检查单元格内是否有竖线 |")
            if rows:
                add_table(doc, rows)
            continue

        # 图片：![图注文字](图片相对路径)
        m_img = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", line)
        if m_img:
            add_image(doc, m_img.group(1).strip(), m_img.group(2).strip(), md_dir)
            i += 1
            continue

        # 四级标题：不支持，退化为加粗段落
        if s.startswith("#### "):
            if "h4" not in _WARNED:
                warn("不支持四级标题 ####，已退化为加粗段落（建议写 **4.4.1 标题**）")
                _WARNED.add("h4")
            p = doc.add_paragraph()
            set_para_format(p, indent=False, align=WD_ALIGN_PARAGRAPH.LEFT)
            fill_paragraph(p, s[5:], bold_base=True)
            i += 1
            continue

        if s.startswith("### "):
            add_heading(doc, s[4:], level=3)
        elif s.startswith("## "):
            add_heading(doc, s[3:], level=2)
        elif s.startswith("# "):
            if first_h1 is None:
                first_h1 = s[2:]
            add_heading(doc, s[2:], level=1)
        elif s.startswith("- "):
            text = re.sub(r"^\[[ xX]\]\s*", "", s[2:])  # 兼容任务列表写法
            p = doc.add_paragraph(style="List Bullet")
            set_para_format(p, indent=False)
            fill_paragraph(p, text)
        elif re.match(r"^\d+\.\s", s):
            p = doc.add_paragraph(style="List Number")
            set_para_format(p, indent=False)
            fill_paragraph(p, re.sub(r"^\d+\.\s", "", s))
        else:
            add_body(doc, line)
        i += 1

    if first_h1:
        doc.core_properties.title = first_h1
    doc.save(docx_path)
    print(f"已生成: {docx_path}")


def add_body(doc, text):
    """正文段落：首行缩进 2 字符、固定 22 磅行距。"""
    p = doc.add_paragraph()
    set_para_format(p, indent=True)
    fill_paragraph(p, text)
    return p


def build_parser():
    parser = argparse.ArgumentParser(
        description="Markdown → Word（中文文档三线表/图注/页眉页码）")
    parser.add_argument("md", help="输入 Markdown 路径")
    parser.add_argument("docx", help="输出 docx 路径")
    parser.add_argument("--project", default=os.environ.get("DOC_PROJECT_NAME"),
                        help="页眉项目名；缺省读环境变量 DOC_PROJECT_NAME，未提供则不生成页眉")
    parser.add_argument("--no-header", action="store_true", help="不生成页眉")
    parser.add_argument("--no-footer", action="store_true", help="不生成页脚页码")
    parser.add_argument("--page-size", default="A4", choices=sorted(PAGE_SIZES),
                        help="纸张大小（默认 A4）")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    _WARNED.clear()
    convert(args.md, args.docx, project=args.project,
            header=not args.no_header, footer=not args.no_footer,
            page_size=args.page_size)
    return 0


if __name__ == "__main__":
    sys.exit(main())
