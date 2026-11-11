# -*- coding: utf-8 -*-
"""中文文档 Markdown → Word (.docx) 转换器（profile 驱动）。

用法:
    python scripts/md2docx.py <输入.md> <输出.docx> [--profile tech-doc] [--project "页眉项目名"]
                              [--toc] [--no-header] [--no-footer] [--page-size A4]

排版参数全部来自 profiles/*.toml（默认 profile: ug-thesis-bit，输出与改造前一致）:
- 页面：纸张、四边页边距、页眉页脚距边界、首页是否单独
- 正文：中西文字体、字号、行距（固定磅/倍数）、首行缩进、对齐、段前段后、字距
- 标题：每级独立 font/size/bold/align/段距/段前分页
- 表格：三线表线宽、单元格字体字号、表头加粗与跨页重复、表注字体字号
- 图注/表注：字体、字号、段距
- 页眉页脚：页眉文字与横线、页码位置与格式（plain/dash/cn/roman）、起始页
- 结构：目录分节 + TOC 域、正文页码起始值

支持的 Markdown 子集:
- # / ## / ### 标题（#### 不支持，会退化为加粗段落并告警）
- | a | b | 表格（首行表头，第二行 |---| 分隔）
- - 无序列表、1. 有序列表
- **加粗** 行内文本
- --- 分隔线（忽略）
- ![图注文字](图片相对路径) 嵌入图片（段落按 profile 的 image_line_spacing）
- 表格正上方一行的 **表 x-x 说明** 渲染为表注（允许与表格之间有空行）
- 封面不经过本脚本：封面在 Word 中手工排版；目录可用 --toc 生成分节 + 空域，打开后「更新域」填充
"""
from __future__ import annotations

import argparse
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import profile as profile_mod  # noqa: E402

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# 当前转换使用的 profile（convert() 内设置；helper 读取它，避免层层传参）
CFG = {}

# 已告警的语法类型，避免同一文件逐行刷屏
_WARNED = set()


def warn(msg):
    """向 stderr 输出告警（同一类型只提示一次由调用方控制）。"""
    print(f"警告: {msg}", file=sys.stderr)


def note(msg):
    """向 stderr 输出提示（不影响退出码，例如 profile 的未核实项与字体替换）。"""
    print(f"提示: {msg}", file=sys.stderr)


def _body_font():
    return CFG["body"]["font_cn"]


def _heading_cfg(level):
    for item in CFG.get("headings", []):
        if item.get("level") == level:
            return item
    return None


def current_line_spacing():
    """按 profile 的行距模式返回 python-docx 的 line_spacing 值。"""
    spec = CFG["body"]["line_spacing"]
    if spec["mode"] == "fixed":
        return Pt(float(spec["value"]))
    if spec["mode"] == "multiple":
        return float(spec["value"])
    raise profile_mod.ProfileError(f"行距模式 grid 未实现: {spec}")


def set_run_font(run, bold=False, size=None, font=None):
    """设置 run 的中西文字体与字号。"""
    size = CFG["body"]["size"] if size is None else size
    font = _body_font() if font is None else font
    run.font.name = CFG["body"]["font_en"]
    run.font.bold = bold
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font)


def set_para_format(p, indent=False, align=None, space_before=None, space_after=None,
                    line_spacing=None):
    """设置段落格式：行距、缩进、对齐、段前段后（支持 pt 与 0.01 行两种单位）。"""
    pf = p.paragraph_format
    pf.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY) if align else \
        ALIGN[CFG["body"]["align"]]
    pf.line_spacing = current_line_spacing() if line_spacing is None else line_spacing
    pPr = p._p.get_or_add_pPr()
    if indent:
        chars = int(CFG["body"]["first_line_indent"]["chars"])
        ind = pPr.get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(chars * 100))
    spacing = None
    if space_before or space_after:
        spacing = pPr.get_or_add_spacing()
    for spec, before in ((space_before, True), (space_after, False)):
        if not spec:
            continue
        if "lines" in spec and spec["lines"]:
            spacing.set(qn("w:beforeLines" if before else "w:afterLines"),
                        str(int(round(float(spec["lines"]) * 100))))
        elif "pt" in spec and spec["pt"]:
            spacing.set(qn("w:before" if before else "w:after"), str(int(round(float(spec["pt"]) * 20))))


def fill_paragraph(p, text, size=None, font=None, bold_base=False):
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
    """按 profile 的逐级标题参数添加标题（真 Heading 样式，供目录域抓取）。"""
    spec = _heading_cfg(level) or {
        "font": "黑体", "size": 16, "bold": False, "align": "left",
        "space_before": {"lines": 0.5}, "space_after": {"lines": 0.5},
    }
    h = doc.add_heading(level=level)
    fill_paragraph(h, text, size=spec["size"], font=spec["font"], bold_base=False)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.bold = bool(spec["bold"])
        set_run_font(run, bold=bool(spec["bold"]), size=spec["size"], font=spec["font"])
    set_para_format(h, indent=False, align=spec.get("align", "left"),
                    space_before=spec.get("space_before"),
                    space_after=spec.get("space_after"))
    if spec.get("page_break_before"):
        h.paragraph_format.page_break_before = True
    return h


def add_caption(doc, text, kind="figure"):
    """图注/表注：字体字号段距按 profile，居中、无缩进。"""
    if kind == "table":
        spec = CFG["table"].get("caption", {})
        font, size = spec.get("font", "宋体"), spec.get("size", 9)
        before, after = spec.get("space_before"), spec.get("space_after")
    else:
        font, size = CFG["figure"]["caption_font"], CFG["figure"]["caption_size"]
        before = CFG["figure"].get("caption_space_before")
        after = CFG["figure"].get("caption_space_after")
    p = doc.add_paragraph()
    set_para_format(p, indent=False, align="center",
                    space_before=before, space_after=after)
    fill_paragraph(p, text, size=size, font=font)
    return p


def set_table_borders(table):
    """三线表：上下框线、表头下框线按 profile 的八分之一磅值。"""
    border = CFG["table"]["border"]
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "bottom"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(int(border["outer"])))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    for edge in ("left", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def set_header_bottom_border(row):
    """表头行单元格下框线。"""
    size = str(int(CFG["table"]["border"]["header"]))
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), size)
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "auto")
        tcBorders.append(bottom)
        tcPr.append(tcBorders)


def set_repeat_header(row):
    """表格首行设为标题行，跨页自动重复。"""
    if not CFG["table"].get("header_repeat", True):
        return
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


def add_table(doc, rows):
    """表格：首行表头按 profile 决定加粗与跨页重复，单元格字体字号按 profile。"""
    cell_font = CFG["table"]["cell_font"]
    cell_size = CFG["table"]["cell_size"]
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
                    set_run_font(p.add_run(part[2:-2]), bold=True,
                                 size=cell_size, font=cell_font)
                else:
                    set_run_font(p.add_run(part), bold=(r == 0),
                                 size=cell_size, font=cell_font)
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
    set_para_format(p, indent=False, align="center",
                    line_spacing=float(CFG["body"]["image_line_spacing"]))
    img = Image.from_file(img_path)
    max_width = Cm(float(CFG["body"]["max_image_width"]))
    width = img.width if img.width <= max_width else max_width
    p.add_run().add_picture(img_path, width=width)
    if alt:
        add_caption(doc, alt, kind="figure")


def _field(paragraph, instr, placeholder="1", size=None, font=None):
    """在段落中插入一个 Word 域。"""
    run = paragraph.add_run()
    set_run_font(run, size=size, font=font)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr_el, separate, text, end):
        run._r.append(el)


def add_page_field(paragraph, size=None, font=None):
    """在段落中插入 Word PAGE 域（页码），打开文档时自动更新。"""
    _field(paragraph, " PAGE ", size=size, font=font)


def set_page_numbering(section, fmt=None, start=None):
    """设置节的页码格式（w:pgNumType）与起始值。"""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        sectPr.append(pgNumType)
    if fmt:
        pgNumType.set(qn("w:fmt"), fmt)
    if start is not None:
        pgNumType.set(qn("w:start"), str(int(start)))


def _write_page_number(paragraph, section_format):
    """按 profile 的页码格式写页脚/页眉段落。"""
    spec = CFG["header_footer"]["page_number"]
    size, font = spec.get("size", 9), spec.get("font", "宋体")
    fmt = spec.get("format", "dash")
    if fmt == "plain":
        add_page_field(paragraph, size=size, font=font)
    elif fmt == "dash":
        set_run_font(paragraph.add_run("- "), size=size, font=font)
        add_page_field(paragraph, size=size, font=font)
        set_run_font(paragraph.add_run(" -"), size=size, font=font)
    elif fmt == "cn":
        set_run_font(paragraph.add_run("第 "), size=size, font=font)
        add_page_field(paragraph, size=size, font=font)
        set_run_font(paragraph.add_run(" 页 共 "), size=size, font=font)
        _field(paragraph, " NUMPAGES ", placeholder="1", size=size, font=font)
        set_run_font(paragraph.add_run(" 页"), size=size, font=font)
    elif fmt == "roman":
        add_page_field(paragraph, size=size, font=font)
    else:  # pragma: no cover  已被 profile 校验拦截
        raise profile_mod.ProfileError(f"未知页码格式: {fmt}")


def _header_bottom_rule(paragraph):
    """页眉下横线（0.5mm 单线）。"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")   # 1 磅
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    pBdr.append(bottom)
    pPr.append(pBdr)


def setup_page(section, project=None, header=True, footer=True, page_size=None,
               start_at=None, front=False, cover_blank=True):
    """设置纸张、页边距、页眉与页脚页码。"""
    page = CFG["page"]
    size_key = page_size or page["size"]
    width_cm, height_cm = profile_mod.PAGE_SIZES.get(
        size_key, profile_mod.PAGE_SIZES["A4"])
    section.page_width = Cm(width_cm)
    section.page_height = Cm(height_cm)
    margins = page["margins"]
    section.top_margin = Cm(margins["top"])
    section.bottom_margin = Cm(margins["bottom"])
    section.left_margin = Cm(margins["left"])
    section.right_margin = Cm(margins["right"])
    section.header_distance = Cm(page["header_distance"])
    section.footer_distance = Cm(page["footer_distance"])
    blank_first = bool(page.get("different_first_page")) and cover_blank
    if blank_first:
        section.different_first_page_header_footer = True

    hf = CFG["header_footer"]
    header_text = hf.get("header_text") or project
    if header and header_text:
        # 首页单独设置时，首页页眉与正文页眉保持一致（只让页码不同）
        targets = [section.header]
        if blank_first:
            targets.append(section.first_page_header)
        for part in targets:
            hp = part.paragraphs[0]
            hp.text = ""
            set_para_format(hp, indent=False, align="center", line_spacing=1.0)
            set_run_font(hp.add_run(header_text), size=hf["header_size"],
                         font=hf["header_font"])
            if hf.get("header_rule") == "single":
                _header_bottom_rule(hp)

    # 前置节若规定不编页码，则页脚留空（连域也不写）
    front_fmt = CFG["structure"].get("front_numbering", "none")
    suppress_number = bool(front) and front_fmt == "none"
    if footer and not suppress_number:
        # 页码位置：footer-center / header-right
        position = hf["page_number"].get("position", "footer-center")
        if position == "header-right":
            target = section.header.paragraphs[0] if header_text else \
                section.header.add_paragraph()
            target.text = ""
            set_para_format(target, indent=False, align="right", line_spacing=1.0)
        else:
            target = section.footer.paragraphs[0]
            target.text = ""
            set_para_format(target, indent=False, align="center", line_spacing=1.0)
        _write_page_number(target, None)

    if front:
        fmt = {"roman": "upperRoman", "plain": "decimal"}.get(front_fmt)
        if fmt:
            set_page_numbering(section, fmt=fmt, start=1)
    elif start_at is not None:
        set_page_numbering(section, fmt="decimal", start=start_at)


def add_toc_front_matter(doc):
    """前置节：目录标题 + TOC 域（打开 Word 后「更新域」生成页码）。"""
    structure = CFG["structure"]
    heading = doc.add_paragraph()
    set_para_format(heading, indent=False, align="center", line_spacing=1.0,
                    space_after={"pt": 18})
    size = (_heading_cfg(1) or {}).get("size", 16)
    font = (_heading_cfg(1) or {}).get("font", "黑体")
    set_run_font(heading.add_run(structure.get("toc_title", "目录")), bold=True,
                 size=size, font=font)

    body = doc.add_paragraph()
    set_para_format(body, indent=False, align="left")
    _field(body, ' TOC \\o "1-3" \\h \\z \\u ', placeholder="右键「更新域」生成目录")


def convert(md_path, docx_path, project=None, header=True, footer=True,
            page_size=None, profile=None, toc=None):
    """将 md 转换为 docx。"""
    global CFG
    CFG, notices = profile_mod.load_profile(profile or profile_mod.DEFAULT_PROFILE)
    for notice in notices:
        note(notice)

    with open(md_path, encoding="utf-8") as f:
        lines = f.read().splitlines()

    doc = Document()

    # 文档默认样式：字体字号行距全部按 profile
    normal = doc.styles["Normal"]
    normal.font.name = CFG["body"]["font_en"]
    normal.font.size = Pt(CFG["body"]["size"])
    normal.paragraph_format.line_spacing = current_line_spacing()
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), _body_font())

    want_toc = CFG["structure"].get("toc", False) if toc is None else toc
    first_section = doc.sections[0]
    if want_toc:
        setup_page(first_section, project=project, header=header,
                   footer=footer, page_size=page_size, front=True,
                   cover_blank=False)
        add_toc_front_matter(doc)
        body_section = doc.add_section(WD_SECTION.NEW_PAGE)
        body_section.header.is_linked_to_previous = False
        body_section.footer.is_linked_to_previous = False
        setup_page(body_section, project=project, header=header, footer=footer,
                   page_size=page_size, cover_blank=False,
                   start_at=CFG["structure"].get("body_start_at", 1))
    else:
        setup_page(first_section, project=project, header=header,
                   footer=footer, page_size=page_size)

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
                add_caption(doc, m_cap.group(1), kind="table")
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
            set_para_format(p, indent=False, align="left")
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
    print(f"已生成: {docx_path}  [{CFG.get('name', '?')}]")


def add_body(doc, text):
    """正文段落：首行缩进、行距与段距按 profile。"""
    p = doc.add_paragraph()
    set_para_format(p, indent=True,
                    space_before=CFG["body"].get("space_before"),
                    space_after=CFG["body"].get("space_after"))
    fill_paragraph(p, text)
    return p


def build_parser():
    parser = argparse.ArgumentParser(
        description="Markdown → Word（profile 驱动：三线表/图注/页眉页码/目录）")
    parser.add_argument("md", help="输入 Markdown 路径")
    parser.add_argument("docx", help="输出 docx 路径")
    parser.add_argument("--profile", default=None,
                        help="排版 profile 名或 .toml 路径（默认 ug-thesis-bit）")
    parser.add_argument("--project", default=os.environ.get("DOC_PROJECT_NAME"),
                        help="页眉项目名；缺省读环境变量 DOC_PROJECT_NAME，未提供则不生成页眉")
    parser.add_argument("--toc", dest="toc", action="store_true", default=None,
                        help="生成「目录」分节 + TOC 域（覆盖 profile 的 structure.toc）")
    parser.add_argument("--no-toc", dest="toc", action="store_false",
                        help="即使 profile 要求也不生成目录分节")
    parser.add_argument("--no-header", action="store_true", help="不生成页眉")
    parser.add_argument("--no-footer", action="store_true", help="不生成页脚页码")
    parser.add_argument("--page-size", default=None, choices=sorted(profile_mod.PAGE_SIZES),
                        help="纸张大小（覆盖 profile 的 page.size）")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    _WARNED.clear()
    try:
        convert(args.md, args.docx, project=args.project,
                header=not args.no_header, footer=not args.no_footer,
                page_size=args.page_size, profile=args.profile, toc=args.toc)
    except profile_mod.ProfileError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
