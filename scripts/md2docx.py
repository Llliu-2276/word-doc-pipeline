# -*- coding: utf-8 -*-
"""中文文档 Markdown → Word (.docx) 转换器（profile 驱动）。

用法:
    python scripts/md2docx.py <输入.md> <输出.docx> [--profile tech-doc] [--project "页眉项目名"]
                              [--toc] [--no-header] [--no-footer] [--page-size A4]

排版参数全部来自 profiles/*.toml（默认 profile: ug-thesis-bit，输出与改造前一致）:
- 页面：纸张、四边页边距、装订线、镜像边距、页眉页脚距边界、首页是否单独
- 正文：中西文字体、字号、行距（固定磅/倍数）、首行缩进或悬挂缩进、对齐、段前段后、字距
- 标题：每级独立 font/size/bold/align/段距/段前分页
- 表格：三线表线宽、单元格字体字号、表头加粗与跨页重复、表注字体字号
- 图注/表注：字体、字号、段距
- 页眉页脚：页眉文字与横线、页码位置与格式（plain/dash/cn/roman）、起始页
- 结构：目录分节 + TOC 域、正文页码起始值、章标题前分节（换页/另起奇数页）

生成的 docx 会写入「生成标记」（docProps/custom.xml 里的 wdp_* 属性），
check_docx.py 据此区分「profile 要求而缺失」与「生成时用 --no-header/--no-footer 显式关闭」。

支持的 Markdown 子集:
- # / ## / ### 标题；#### 在 profile 定义了 4 级标题时按 4 级处理，否则退化为加粗段落并告警
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
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.image.image import Image
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
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

# 表格分隔行：| --- | --- |（识别表格的必要条件，第二行不是它就不是表格）
TABLE_SEP_RE = re.compile(r"^\|[\s:\-|]+\|?$")

# 文件开头为 UTF-8 BOM（记事本默认写法）；读入时剥掉，check_md.py 会单独报错
UTF8_BOM = b"\xef\xbb\xbf"

# 当前转换使用的 profile（convert() 内设置；helper 读取它，避免层层传参）
CFG = {}

# 已告警的语法类型，避免同一文件逐行刷屏
_WARNED = set()

# 生成标记：写进 docProps/custom.xml，供 check_docx.py 判断「profile 要求而缺失」
# 还是「生成时用 --no-header / --no-footer / --no-toc 显式关掉了」
GENERATOR_TAG = "word-doc-pipeline"
CUSTOM_PROPS_CT = ("application/vnd.openxmlformats-officedocument."
                   "custom-properties+xml")
CUSTOM_PROPS_NS = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                   "custom-properties")
CUSTOM_PROPS_VT_NS = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                      "docPropsVTypes")
_CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# w:mirrorMargins 之后可能出现的元素（按 CT_Settings 顺序），用于插到合法位置
_MIRROR_SUCCESSORS = (
    "w:defaultTabStop", "w:characterSpacingControl", "w:compat", "w:rsids",
    "m:mathPr", "w:themeFontLang", "w:clrSchemeMapping", "w:shapeDefaults",
    "w:decimalSymbol", "w:listSeparator",
)


def write_generation_tags(document, values):
    """把生成参数写成 docProps/custom.xml 的自定义属性（wdp_<name>）。"""
    props = []
    for index, (name, value) in enumerate(values, start=2):
        props.append(
            f'<property fmtid="{_CUSTOM_FMTID}" pid="{index}" '
            f'name="wdp_{escape(str(name))}">'
            f'<vt:lpwstr>{escape(str(value))}</vt:lpwstr></property>')
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<Properties xmlns="{CUSTOM_PROPS_NS}" '
           f'xmlns:vt="{CUSTOM_PROPS_VT_NS}">' + "".join(props) + "</Properties>")
    package = document.part.package
    # 已存在（例如在模板上重复生成）就原地覆盖，避免重复 part 与重复关系
    for part in package.iter_parts():
        if str(part.partname) == "/docProps/custom.xml":
            part._blob = xml.encode("utf-8")
            return
    part = Part(PackURI("/docProps/custom.xml"), CUSTOM_PROPS_CT,
                xml.encode("utf-8"), package)
    package.relate_to(part, RT.CUSTOM_PROPERTIES)


def set_mirror_margins(document, enabled):
    """奇偶页镜像边距（w:mirrorMargins），插入到 CT_Settings 允许的位置。"""
    settings = document.settings.element
    existing = settings.find(qn("w:mirrorMargins"))
    if not enabled:
        if existing is not None:
            settings.remove(existing)
        return
    if existing is None:
        settings.insert_element_before(OxmlElement("w:mirrorMargins"),
                                       *_MIRROR_SUCCESSORS)


def apply_body_indent(pPr):
    """按 profile 的 first_line_indent.mode 写首行缩进或悬挂缩进（字符单位）。"""
    spec = CFG["body"].get("first_line_indent") or {}
    chars = int(spec.get("chars") or 0)
    mode = spec.get("mode", "first_line")
    if mode == "none" or chars <= 0:
        return
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:hangingChars" if mode == "hanging" else "w:firstLineChars"),
            str(chars * 100))


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


def _grid_line_pitch():
    """grid 行距模式下的固定行高（磅）：版心高 / 每页行数。

    Word 的 w:docGrid 只保证字符格宽度，实测行高仍由段落行距决定
    （22 行 × 28 字、版心 225mm 时 Word 实测行高 579 twips ≈ 28.95 磅）。
    """
    body = CFG["body"]
    grid = body.get("grid") or {}
    lines = int(grid.get("lines") or 0)
    if body["line_spacing"]["mode"] != "grid" or not lines:
        return None
    page = CFG["page"]
    width_cm, height_cm = profile_mod.PAGE_SIZES.get(
        page["size"], profile_mod.PAGE_SIZES["A4"])
    text_height = height_cm - page["margins"]["top"] - page["margins"]["bottom"]
    return text_height / lines * 72 / 2.54


def current_line_spacing():
    """按 profile 的行距模式返回 line_spacing 值；grid 模式返回 None（由文档网格控制）。"""
    spec = CFG["body"]["line_spacing"]
    if spec["mode"] == "fixed":
        return Pt(float(spec["value"]))
    if spec["mode"] == "multiple":
        return float(spec["value"])
    if spec["mode"] == "grid":
        pitch = _grid_line_pitch()
        return Pt(round(pitch, 2)) if pitch else None
    raise profile_mod.ProfileError(f"未知行距模式: {spec['mode']}")


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
    if line_spacing is None:
        spacing = current_line_spacing()
        if spacing is None:  # grid 模式：单倍行距，行高由 w:docGrid 决定
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        else:
            pf.line_spacing = spacing
    else:
        pf.line_spacing = line_spacing
    pPr = p._p.get_or_add_pPr()
    if indent:
        apply_body_indent(pPr)
    spacing = None
    if space_before or space_after:
        spacing = pPr.get_or_add_spacing()
    for spec, before in ((space_before, True), (space_after, False)):
        if not spec:
            continue
        if spec.get("lines") is not None:
            spacing.set(qn("w:beforeLines" if before else "w:afterLines"),
                        str(int(round(float(spec["lines"]) * 100))))
        elif spec.get("pt") is not None:
            spacing.set(qn("w:before" if before else "w:after"),
                        str(int(round(float(spec["pt"]) * 20))))


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
    """表格框线：三线表（上下框线 + 无竖线）或全线表（六条边全画）。"""
    border = CFG["table"]["border"]
    style = CFG["table"].get("style", "three_line")
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    if style == "grid":
        size = str(int(border.get("grid", border["outer"])))
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement("w:" + edge)
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), size)
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
            borders.append(el)
        tblPr.append(borders)
        return
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
    cols = len(rows[0])
    table = doc.add_table(rows=len(data_rows), cols=cols)
    table.autofit = True
    set_table_borders(table)
    for r, cells in enumerate(data_rows):
        for c, text in enumerate(cells[:cols]):  # 列数已在 convert() 里对齐，这里再兜一层
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
    """嵌入图片（宽度不超页面），图注按 profile 置于图片下方或上方。"""
    img_path = path if os.path.isabs(path) else os.path.join(md_dir, path)
    if not os.path.isfile(img_path):
        warn(f"图片不存在，已跳过: {img_path}")
        return
    try:
        img = Image.from_file(img_path)
    except Exception as exc:  # 文件损坏 / 不是图片
        warn(f"图片无法解析，已跳过: {img_path}（{exc}）")
        return
    max_width = Cm(float(CFG["body"]["max_image_width"]))
    width = img.width if img.width <= max_width else max_width
    caption_below = CFG["numbering"]["figure"].get("caption", "below") != "above"
    if alt and not caption_below:  # profile 要求图注在图片上方
        add_caption(doc, alt, kind="figure")
    p = doc.add_paragraph()
    set_para_format(p, indent=False, align="center",
                    line_spacing=float(CFG["body"]["image_line_spacing"]))
    try:
        p.add_run().add_picture(img_path, width=width)
    except Exception as exc:
        warn(f"图片无法嵌入，已跳过: {img_path}（{exc}）")
        p._element.getparent().remove(p._element)
        return
    if alt and caption_below:
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


def set_doc_grid(section):
    """按 profile 写文档网格（每页行数 × 每行字数），仅 grid 行距模式生效。

    w:docGrid 的编码以 Word 实测为准（A4、37/35/28/26mm、22 行 × 28 字）：
    linePitch = 版心高(twips) / 行数；charSpace = (每字格宽 - 字号) 磅值 × 4096 / 20。
    """
    body = CFG["body"]
    grid = body.get("grid") or {}
    lines = int(grid.get("lines") or 0)
    chars = int(grid.get("chars") or 0)
    if body["line_spacing"]["mode"] != "grid" or (not lines and not chars):
        return
    sectPr = section._sectPr
    old = sectPr.find(qn("w:docGrid"))
    if old is not None:
        sectPr.remove(old)
    el = OxmlElement("w:docGrid")
    el.set(qn("w:type"), "linesAndChars")
    height = (section.page_height.twips - section.top_margin.twips
              - section.bottom_margin.twips)
    width = (section.page_width.twips - section.left_margin.twips
             - section.right_margin.twips - section.gutter.twips)
    if lines:
        el.set(qn("w:linePitch"), str(int(height / lines)))
    if chars:
        diff = width / chars - float(body["size"]) * 20   # twips
        el.set(qn("w:charSpace"), str(int(round(diff * 4096 / 20))))
    sectPr.append(el)


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
    elif fmt == "emdash":
        # 公文：一字线页码「— 1 —」，4 号宋体左右各一字线
        set_run_font(paragraph.add_run("— "), size=size, font=font)
        add_page_field(paragraph, size=size, font=font)
        set_run_font(paragraph.add_run(" —"), size=size, font=font)
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


def _right_tab_stop(paragraph, section):
    """在版心右边缘加右对齐制表位（页眉右侧页码与页眉文字共用一行）。"""
    width = (section.page_width - section.left_margin - section.right_margin
             - section.gutter)
    paragraph.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)


def setup_page(section, project=None, header=True, footer=True, page_size=None,
               start_at=None, front=False, cover_blank=True, document=None):
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
    section.gutter = Cm(float(page.get("gutter") or 0.0))
    section.header_distance = Cm(page["header_distance"])
    section.footer_distance = Cm(page["footer_distance"])
    blank_first = bool(page.get("different_first_page")) and cover_blank
    if blank_first:
        section.different_first_page_header_footer = True

    hf = CFG["header_footer"]
    position = hf["page_number"].get("position", "footer-center")
    header_text = hf.get("header_text") or project
    if header and not hf.get("header_enabled", True):
        header = False   # profile 声明不要页眉（公文类）
    header_parts = []
    if header and header_text:
        # 首页单独设置时，首页页眉与正文页眉保持一致（只让页码不同）
        header_parts = [section.header]
        if blank_first:
            header_parts.append(section.first_page_header)
        # 页码在页眉右侧时，页眉文字与页码共用一行：文字靠左、页码靠右制表位，
        # 不能像页脚那样先 text="" 再写域，否则会把项目名抹掉。
        align = "left" if position == "header-right" else "center"
        for part in header_parts:
            hp = part.paragraphs[0]
            hp.text = ""
            set_para_format(hp, indent=False, align=align, line_spacing=1.0)
            set_run_font(hp.add_run(header_text), size=hf["header_size"],
                         font=hf["header_font"])
            if hf.get("header_rule") == "single":
                _header_bottom_rule(hp)

    # 前置节若规定不编页码，则页脚留空（连域也不写）
    front_fmt = CFG["structure"].get("front_numbering", "none")
    suppress_number = bool(front) and front_fmt == "none"
    if footer and not suppress_number:
        # 页码位置：footer-center / header-right / odd-even-outside（公文式外侧）
        if position == "odd-even-outside":
            odd = section.footer.paragraphs[0]
            odd.text = ""
            set_para_format(odd, indent=False, align="right", line_spacing=1.0)
            _write_page_number(odd, None)
            even = section.even_page_footer.paragraphs[0]
            even.text = ""
            set_para_format(even, indent=False, align="left", line_spacing=1.0)
            _write_page_number(even, None)
            if document is not None:
                document.settings.odd_and_even_pages_header_footer = True
        elif position == "header-right":
            # 页码与页眉文字共用一行：文字靠左，制表位把页码推到版心右边缘。
            # 不能像页脚那样先 text="" 再写域，否则会把页眉项目名抹掉。
            targets = list(header_parts) or [section.header]
            if blank_first and not header_parts:
                targets.append(section.first_page_header)
            for target in targets:
                para = target.paragraphs[0]
                if not header_parts:   # --no-header：页眉本来是空的
                    para.text = ""
                    set_para_format(para, indent=False, align="left", line_spacing=1.0)
                _right_tab_stop(para, section)
                para.add_run("\t")
                _write_page_number(para, None)
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
    set_doc_grid(section)


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
    # 目录收录到第几级标题：沿用 numbering.max_depth（编号/目录层级同一个概念）
    depth = max(1, min(9, int(CFG["numbering"].get("max_depth") or 3)))
    _field(body, f' TOC \\o "1-{depth}" \\h \\z \\u ',
           placeholder="右键「更新域」生成目录")


def convert(md_path, docx_path, project=None, header=True, footer=True,
            page_size=None, profile=None, toc=None, template_doc=None,
            anchor=None):
    """将 md 转换为 docx。

    template_doc 不为 None 时走「模板接管」：以既有 docx 为容器，只注入正文，
    不改页面设置、不加页眉页脚、不生成目录、不做章前分节；anchor 是模板里
    {{正文}} 所在的段落，注入内容会被移到它前面并删除该段落（anchor 为
    None 时内容追加到文末）。
    """
    global CFG
    CFG, notices = profile_mod.load_profile(profile or profile_mod.DEFAULT_PROFILE)
    for notice in notices:
        note(notice)

    with open(md_path, "rb") as f:
        raw = f.read()
    if raw.startswith(UTF8_BOM):
        # 记事本默认给 UTF-8 加 BOM，不剥掉会把首行写成「\ufeff# 标题」的普通段落
        note(f"输入 md 带 UTF-8 BOM，已自动忽略（{os.path.basename(md_path)}）")
        raw = raw[len(UTF8_BOM):]
    try:
        lines = raw.decode("utf-8").splitlines()
    except UnicodeDecodeError as exc:
        raise profile_mod.ProfileError(
            f"{md_path} 不是 UTF-8 编码（{exc}），请另存为 UTF-8 后重试")

    template_mode = template_doc is not None
    doc = template_doc if template_mode else Document()
    # 模板接管时先记下正文子元素个数，渲染完把新增元素整体挪到 {{正文}} 锚点处
    # （不能用 id() 比较：lxml 的元素代理对象不稳定，同一节点两次遍历可能不是同一对象）
    before_children = list(doc.element.body) if template_mode else []
    if not template_mode:
        set_mirror_margins(doc, bool(CFG["page"].get("mirror")))

    # 文档默认样式：字体字号行距全部按 profile（模板接管时保留模板样式）
    if not template_mode:
        normal = doc.styles["Normal"]
        normal.font.name = CFG["body"]["font_en"]
        normal.font.size = Pt(CFG["body"]["size"])
        normal.paragraph_format.line_spacing = current_line_spacing()
        normal.paragraph_format.space_before = Pt(0)
        normal.paragraph_format.space_after = Pt(0)
        rPr = normal.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = rPr.makeelement(qn("w:rFonts"), {})
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), _body_font())

    want_toc = False if template_mode else (
        CFG["structure"].get("toc", False) if toc is None else toc)
    first_section = doc.sections[0]
    if template_mode:
        pass  # 页面设置、页眉页脚、目录、分节全部由模板负责
    elif want_toc:
        setup_page(first_section, project=project, header=header,
                   footer=footer, page_size=page_size, front=True,
                   cover_blank=False, document=doc)
        add_toc_front_matter(doc)
        body_section = doc.add_section(WD_SECTION.NEW_PAGE)
        body_section.header.is_linked_to_previous = False
        body_section.footer.is_linked_to_previous = False
        setup_page(body_section, project=project, header=header, footer=footer,
                   page_size=page_size, cover_blank=False, document=doc,
                   start_at=CFG["structure"].get("body_start_at", 1))
    else:
        setup_page(first_section, project=project, header=header,
                   footer=footer, page_size=page_size, document=doc)

    if CFG["page"].get("mirror") and not template_mode:
        note("启用镜像边距（page.mirror）：Word 会在正文前补一张空白页，"
             "使正文从右页（recto）开始，装订成册时属正常版式，不是脚本多输出一页。")

    md_dir = os.path.dirname(os.path.abspath(md_path))
    chapter_break = ("none" if template_mode
                     else CFG["structure"].get("chapter_break", "none"))
    chapter_level = int(CFG["structure"].get("chapter_level") or 1)
    chapters = 0
    pending_table_caption = None   # profile 要求表注在表格下方时，先存着晚一步输出

    def start_chapter_section():
        """章标题前分节（换页 / 另起奇数页），页码沿用前一节、不重编。"""
        new = doc.add_section(WD_SECTION.ODD_PAGE if chapter_break == "odd_page"
                              else WD_SECTION.NEW_PAGE)
        new.header.is_linked_to_previous = False
        new.footer.is_linked_to_previous = False
        # add_section 会复制上一节的节属性，页码起始值必须清掉，否则每章重编
        old_pgnum = new._sectPr.find(qn("w:pgNumType"))
        if old_pgnum is not None:
            new._sectPr.remove(old_pgnum)
        setup_page(new, project=project, header=header, footer=footer,
                   page_size=page_size, cover_blank=False, document=doc)

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
                    and TABLE_SEP_RE.match(lines[j + 1].strip()):
                if CFG["numbering"]["table"].get("caption", "above") == "below":
                    pending_table_caption = m_cap.group(1)   # 表格画完再输出
                else:
                    add_caption(doc, m_cap.group(1), kind="table")
                i = j
                continue

        # 表格：表头行 + |---| 分隔行 + 数据行（缺分隔行就不是表格）
        if s.startswith("|") and i + 1 < len(lines) \
                and TABLE_SEP_RE.match(lines[i + 1].strip()):
            start = i + 1
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                rows.append(cells)
                i += 1
            cols = len(rows[0])
            for offset, cells in enumerate(rows):
                if len(cells) != cols:
                    warn(f"第 {start + offset} 行有 {len(cells)} 列，表头是 {cols} 列，"
                         f"已按表头补齐/截断；单元格内的竖线请写成 &#124;")
                    rows[offset] = (cells + [""] * cols)[:cols]
            add_table(doc, rows)
            if pending_table_caption:
                add_caption(doc, pending_table_caption, kind="table")
                pending_table_caption = None
            # 表注的书写位置固定在上方，profile 的 numbering.table.caption 只决定渲染位置。
            # 写在表格下方的「**表 x-x**」会被当成普通加粗段落，这里提醒一次。
            nxt = i
            while nxt < len(lines) and not lines[nxt].strip():
                nxt += 1
            if nxt < len(lines) and re.match(r"^\*\*表\s*\S+", lines[nxt].strip()) \
                    and "table-cap-after" not in _WARNED:
                warn("发现写在表格下方的表注，已按普通加粗段落输出；"
                     "表注请写在表格上方，渲染位置由 numbering.table.caption 决定")
                _WARNED.add("table-cap-after")
            continue

        # 疑似表格但第二行不是 |---|：脚本不识别，会按正文输出，提醒一次
        if s.startswith("|") and s.count("|") >= 2 and "table-sep" not in _WARNED:
            warn("发现以 | 开头的行但第二行不是 |---| 分隔行，已按正文输出；"
                 "表格必须在表头下一行写 | --- | --- |")
            _WARNED.add("table-sep")

        # 图片：![图注文字](图片相对路径)
        m_img = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", line)
        if m_img:
            add_image(doc, m_img.group(1).strip(), m_img.group(2).strip(), md_dir)
            i += 1
            continue

        # 标题：# ~ ####（#### 只在 profile 定义了 4 级标题时按标题处理）
        m_head = re.match(r"^(#{1,4}) (.+)$", s)
        if m_head and (len(m_head.group(1)) < 4 or _heading_cfg(4) is not None):
            level, text = len(m_head.group(1)), m_head.group(2).strip()
            if level == chapter_level and chapter_break != "none":
                chapters += 1
                if chapters > 1:  # 首个章标题之前不空出一页
                    start_chapter_section()
            if level == 1 and first_h1 is None:
                first_h1 = text
            add_heading(doc, text, level=level)
            i += 1
            continue

        # 四级标题：profile 未定义 4 级标题时退化为加粗段落
        if s.startswith("#### "):
            if "h4" not in _WARNED:
                warn("当前 profile 未定义 4 级标题，#### 已退化为加粗段落"
                     "（需要真标题请在该 profile 中加 [[headings]] level = 4）")
                _WARNED.add("h4")
            p = doc.add_paragraph()
            set_para_format(p, indent=False, align="left")
            fill_paragraph(p, s[5:], bold_base=True)
            i += 1
            continue

        if s.startswith("- "):
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

    if first_h1 and not template_mode:
        doc.core_properties.title = first_h1
    if template_mode:
        _relocate_to_anchor(doc, before_children, anchor)
        write_generation_tags(doc, [
            ("generator", GENERATOR_TAG),
            ("profile", CFG.get("name", "")),
            ("mode", "template"),
        ])
    else:
        header_text = CFG["header_footer"].get("header_text") or project
        write_generation_tags(doc, [
            ("generator", GENERATOR_TAG),
            ("profile", CFG.get("name", "")),
            ("profile_chain", " > ".join(CFG.get("_meta", {}).get("chain", []))),
            ("page_size", page_size or CFG["page"]["size"]),
            ("header", "1" if (header and header_text) else "0"),
            ("footer", "1" if footer else "0"),
            ("toc", "1" if want_toc else "0"),
            ("chapter_break", chapter_break),
        ])
    doc.save(docx_path)
    suffix = "  [模板接管]" if template_mode else ""
    print(f"已生成: {docx_path}  [{CFG.get('name', '?')}]{suffix}")


def _relocate_to_anchor(doc, before_children, anchor):
    """把这次渲染新增的段落/表格移到锚点段落之前，并删除锚点。

    python-docx 只能往文末追加，所以先在文末生成，再按顺序搬过去；
    图片关系挂在 document part 上，整体搬 XML 不会丢。
    """
    body = doc.element.body
    children = list(body)
    # 新增元素都插在 sectPr 之前，所以切点是「快照长度 − 末尾 sectPr」
    cut = len(before_children)
    if before_children and before_children[-1].tag == qn("w:sectPr"):
        cut -= 1
    new_elements = [el for el in children[cut:] if el.tag != qn("w:sectPr")]
    if not new_elements:
        return
    if anchor is None:
        return  # 没有锚点就留在文末（追加模式）
    anchor_el = getattr(anchor, "_p", anchor)
    for el in new_elements:
        anchor_el.addprevious(el)
    anchor_el.getparent().remove(anchor_el)


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
