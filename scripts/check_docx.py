# -*- coding: utf-8 -*-
"""Word 成品格式检查（期望值全部来自同一份 profile）。

用法:
    python scripts/check_docx.py <docx文件或目录>... [--profile tech-doc] [--project "页眉项目名"]

检查项（与 md2docx.py 使用同一份 profiles/*.toml，因此检查标准与生成标准不会漂移）:
- 纸张与四边页边距（page.size / page.margins）
- 页眉文字（header_footer.header_text 或 --project）
- 页码域：PAGE（cn 格式还要 NUMPAGES），以及页码位置所在部件
- 图片：数量、单张宽度 ≤ body.max_image_width、缺失图片时提示
- 表格：数量、三线表 tblBorders、表头下框线
- 正文：Normal 样式东亚字体/字号、首个正文段落行距（固定磅或倍数）
- 标题：一至三级标题的字体、字号、加粗
- 目录：profile 要求 structure.toc 时检查 TOC 域是否存在

profile 中标了 verified = false 的取值只报警告、不报错误（该值尚未对照国标核实）。
退出码：0 通过（可含警告）；1 存在错误。
"""
from __future__ import annotations

import argparse
import os
import sys
import zipfile

import docx
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import profile as profile_mod  # noqa: E402

TOL_CM = 0.15
TOL_PT = 0.5

_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _near(value, expected, tol=TOL_CM):
    return abs(value - expected) <= tol


def _run_east_asia(run):
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return None
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return None
    return rFonts.get(qn("w:eastAsia"))


def _style_east_asia(document, style_name):
    try:
        style = document.styles[style_name]
    except KeyError:
        return None
    rPr = style.element.find(qn("w:rPr"))
    if rPr is None:
        return None
    rFonts = rPr.find(qn("w:rFonts"))
    return rFonts.get(qn("w:eastAsia")) if rFonts is not None else None


def _spacing_text(value):
    """把 line_spacing 值转成可读文本。"""
    if value is None:
        return "未设置"
    if isinstance(value, float):
        return f"{value:g} 倍行距"
    return f"固定 {value.pt:g} 磅"


def _spacing_matches(value, spec):
    """判断段落的 line_spacing 是否符合 profile 的行距定义。"""
    if value is None:
        return False
    if spec["mode"] == "fixed":
        return not isinstance(value, float) and abs(value.pt - float(spec["value"])) <= TOL_PT
    if spec["mode"] == "multiple":
        return isinstance(value, float) and abs(value - float(spec["value"])) <= 0.05
    return True


def _spacing_expected_text(spec):
    if spec["mode"] == "fixed":
        return f"固定 {float(spec['value']):g} 磅"
    return f"{float(spec['value']):g} 倍行距"


def check_docx(path, project=None, profile=None, cfg=None):
    """检查单个 docx，返回 (errors, warnings, stats)。cfg 可传入已加载的 profile。"""
    if cfg is None:
        cfg, _ = profile_mod.load_profile(profile or profile_mod.DEFAULT_PROFILE,
                                          check_fonts=False)

    errors, warnings = [], []
    stats = {"images": 0, "tables": 0, "page": "?", "profile": cfg.get("name", "?")}

    def report(block, message):
        """verified = false 的参数只报警告。"""
        if isinstance(block, dict) and block.get("verified") is False:
            warnings.append(message + "（该值 profile 标记为未核实）")
        else:
            errors.append(message)

    try:
        document = docx.Document(path)
    except Exception as exc:  # 文件损坏/不是 docx
        return [f"无法打开: {exc}"], [], stats

    if not document.sections:
        return ["文档没有 section，无法检查页面设置"], [], stats

    # ---------------- 页面与页边距
    page = cfg["page"]
    expected_w, expected_h = profile_mod.PAGE_SIZES.get(page["size"],
                                                        profile_mod.PAGE_SIZES["A4"])
    for index, section in enumerate(document.sections, start=1):
        width, height = section.page_width.cm, section.page_height.cm
        label = "" if len(document.sections) == 1 else f"第 {index} 节"
        if _near(width, expected_w) and _near(height, expected_h):
            stats["page"] = page["size"]
        else:
            stats["page"] = f"{width:.2f}×{height:.2f}cm"
            report(page, f"{label}页面 {width:.2f}×{height:.2f}cm，"
                         f"profile「{cfg.get('name')}」要求 {page['size']}"
                         f"（{expected_w:.2f}×{expected_h:.2f}cm）")
        margins = page["margins"]
        for name, value, expected in (
            ("上边距", section.top_margin.cm, margins["top"]),
            ("下边距", section.bottom_margin.cm, margins["bottom"]),
            ("左边距", section.left_margin.cm, margins["left"]),
            ("右边距", section.right_margin.cm, margins["right"]),
        ):
            if not _near(value, expected):
                report(page, f"{label}{name} {value:.2f}cm，"
                             f"profile 要求 {expected:.2f}cm")

    section = document.sections[0]
    # 分节文档（目录 + 正文）的页码/页眉在正文节上，前置节可能刻意留空
    multi_section = len(document.sections) > 1
    body_section = document.sections[-1] if multi_section else section
    hf = cfg["header_footer"]

    # ---------------- 页眉
    header_text = "\n".join(p.text for p in body_section.header.paragraphs).strip()
    expected_header = hf.get("header_text") or project
    if expected_header:
        if expected_header not in header_text:
            warnings.append(f"页眉未找到「{expected_header}」（当前: {header_text or '空'}）")

    # ---------------- 页码
    fmt = hf["page_number"].get("format", "dash")
    position = hf["page_number"].get("position", "footer-center")
    part = body_section.header if position == "header-right" else body_section.footer
    part_xml = part._element.xml
    if "PAGE" not in part_xml:
        warnings.append(f"页码域缺失（profile 要求 {fmt} 格式，位置 {position}）")
    elif fmt == "cn" and "NUMPAGES" not in part_xml:
        warnings.append("页码为 cn 格式（第 X 页 共 Y 页），但缺少 NUMPAGES 域")
    part_text = "\n".join(p.text for p in part.paragraphs).strip()
    if fmt == "dash" and part_text and not part_text.startswith("-"):
        warnings.append(f"页码格式应为「- 1 -」，实际为「{part_text}」")

    # ---------------- 图片
    try:
        names = zipfile.ZipFile(path).namelist()
        media = [n for n in names if n.startswith("word/media/") and not n.endswith("/")]
        stats["images"] = len(media)
    except Exception:
        media = []

    inline = len(document.inline_shapes)
    if media and inline and len(media) != inline:
        warnings.append(f"嵌入图片 {len(media)} 张，内联图 {inline} 张（可能有浮动图片，需目视核查）")

    max_img = float(cfg["body"]["max_image_width"])
    for index, shape in enumerate(document.inline_shapes, start=1):
        try:
            width_cm = shape.width.cm
        except Exception:
            continue
        if width_cm > max_img + 0.1:
            report(cfg["body"], f"第 {index} 张图片宽度 {width_cm:.2f}cm > {max_img}cm")

    # ---------------- 表格
    stats["tables"] = len(document.tables)
    if cfg["table"]["style"] == "three_line":
        for index, table in enumerate(document.tables, start=1):
            tbl_pr = table._tbl.tblPr
            if tbl_pr.find(_W_NS + "tblBorders") is None:
                warnings.append(f"第 {index} 个表格缺少 tblBorders，可能不是三线表")
                break
            first_cell = table.rows[0].cells[0]
            tc_pr = first_cell._tc.find(qn("w:tcPr"))
            if tc_pr is None or tc_pr.find(qn("w:tcBorders")) is None:
                warnings.append(f"第 {index} 个表格表头无下框线，三线表可能不完整")
                break

    # ---------------- 正文
    body = cfg["body"]
    east_asia = _style_east_asia(document, "Normal")
    if east_asia and east_asia != body["font_cn"]:
        report(body, f"Normal 样式东亚字体为 {east_asia}，profile 要求 {body['font_cn']}")
    try:
        normal = document.styles["Normal"]
        if normal.font.size and abs(normal.font.size.pt - float(body["size"])) > TOL_PT:
            report(body, f"Normal 字号 {normal.font.size.pt}pt，profile 要求 {body['size']}pt")
    except KeyError:
        pass

    # 跳过目录等前置内容，只抽查正文首段
    started = not (multi_section or cfg["structure"].get("toc"))
    for paragraph in document.paragraphs:
        if not started:
            if paragraph.style.name.startswith("Heading"):
                started = True
            continue
        if paragraph.style.name == "Normal" and paragraph.text.strip():
            spacing = paragraph.paragraph_format.line_spacing
            if not _spacing_matches(spacing, body["line_spacing"]):
                report(body, f"首个正文段落行距 {_spacing_text(spacing)}，"
                             f"profile 要求 {_spacing_expected_text(body['line_spacing'])}")
            break

    # ---------------- 标题
    headings = {item["level"]: item for item in cfg.get("headings", [])}
    seen = {}
    for paragraph in document.paragraphs:
        name = paragraph.style.name
        if not name.startswith("Heading "):
            continue
        level = int(name.split()[-1])
        if level in seen or level not in headings:
            continue
        seen[level] = True
        spec = headings[level]
        run = paragraph.runs[0] if paragraph.runs else None
        if run is None:
            continue
        size = run.font.size.pt if run.font.size else None
        if size is None or abs(size - float(spec["size"])) > TOL_PT:
            report(spec, f"{level} 级标题「{paragraph.text[:12]}」字号 "
                         f"{size or '未设置'}pt，profile 要求 {spec['size']}pt")
        font = _run_east_asia(run)
        if font and font != spec["font"]:
            report(spec, f"{level} 级标题字体为 {font}，profile 要求 {spec['font']}")
        if bool(run.font.bold) != bool(spec["bold"]):
            report(spec, f"{level} 级标题加粗为 {run.font.bold}，"
                         f"profile 要求 {spec['bold']}")

    # ---------------- 目录
    if cfg["structure"].get("toc"):
        body_xml = document.element.body.xml
        if "TOC \\o" not in body_xml and " TOC " not in body_xml:
            warnings.append("profile 要求自动目录，但文档中未发现 TOC 域")

    return errors, warnings, stats


def iter_docx(paths):
    for path in paths:
        if os.path.isdir(path):
            for root, _, files in os.walk(path):
                for name in sorted(files):
                    if name.endswith(".docx") and not name.startswith("~$"):
                        yield os.path.join(root, name)
        else:
            yield path


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="检查 docx 是否符合 profile 规定的排版参数")
    parser.add_argument("paths", nargs="+", help="docx 文件或目录")
    parser.add_argument("--profile", default=None,
                        help="排版 profile 名或 .toml 路径（默认 ug-thesis-bit）")
    parser.add_argument("--project", default=os.environ.get("DOC_PROJECT_NAME"),
                        help="页眉应包含的项目名")
    args = parser.parse_args(argv)

    try:
        cfg, notices = profile_mod.load_profile(args.profile or profile_mod.DEFAULT_PROFILE,
                                                check_fonts=False)
    except profile_mod.ProfileError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    for notice in notices:
        print(f"提示: {notice}", file=sys.stderr)

    files = errors_total = warnings_total = 0
    for path in iter_docx(args.paths):
        files += 1
        errors, warnings, stats = check_docx(path, project=args.project, cfg=cfg)
        errors_total += len(errors)
        warnings_total += len(warnings)
        try:
            rel = os.path.relpath(path)
        except ValueError:  # Windows 下跨盘符路径无法相对化
            rel = os.path.abspath(path)
        status = "FAIL" if errors else "OK"
        print(f"[{status}] {rel}  (profile {stats['profile']} / 页面 {stats['page']} / "
              f"图 {stats['images']} / 表 {stats['tables']})")
        for message in errors:
            print(f"  错误 {message}")
        for message in warnings:
            print(f"  警告 {message}")

    print(f"\n共检查 {files} 个文件：错误 {errors_total}，警告 {warnings_total}")
    return 1 if errors_total else 0


if __name__ == "__main__":
    sys.exit(main())
