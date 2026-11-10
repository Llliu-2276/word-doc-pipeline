# -*- coding: utf-8 -*-
"""Word 成品格式检查（对照 references/02-Word格式规范.md）。

用法:
    python scripts/check_docx.py <docx文件或目录>... [--project "页眉项目名"]

检查项:
- 纸张 A4（21.0×29.7cm）、页边距（上下 2.54cm、左右 3.18cm）
- 页眉（可选：--project 时校验项目名）、页脚 PAGE 页码域
- 图片：数量、单张宽度 ≤14.5cm、是否存在缺失（media 数与内联图数不一致时提示）
- 表格：数量、是否三线表（tblBorders / 表头下框线）
- 正文字体：Normal 样式东亚字体是否为宋体
- 段落行距：正文首段固定 22 磅（抽查）

退出码：0 通过（可含警告）；1 存在错误。
"""
from __future__ import annotations

import argparse
import os
import sys
import zipfile

import docx
from docx.shared import Cm, Pt

A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
LETTER_WIDTH_CM = 21.59
LETTER_HEIGHT_CM = 27.94
TOL_CM = 0.15
EXPECT_TOP_BOTTOM = 2.54
EXPECT_LEFT_RIGHT = 3.18
MAX_IMG_WIDTH_CM = 14.5


def _near(value, expected, tol=TOL_CM):
    return abs(value - expected) <= tol


def check_docx(path, project=None):
    """检查单个 docx，返回 (errors, warnings, stats)。"""
    errors, warnings = [], []
    stats = {"images": 0, "tables": 0, "page": "?"}
    try:
        document = docx.Document(path)
    except Exception as exc:  # 文件损坏/不是 docx
        return [(f"无法打开: {exc}")], [], stats

    if not document.sections:
        return ["文档没有 section，无法检查页面设置"], [], stats
    section = document.sections[0]

    width = section.page_width.cm
    height = section.page_height.cm
    if _near(width, A4_WIDTH_CM) and _near(height, A4_HEIGHT_CM):
        stats["page"] = "A4"
    elif _near(width, LETTER_WIDTH_CM) and _near(height, LETTER_HEIGHT_CM):
        stats["page"] = "Letter"
        errors.append(
            f"页面为 Letter（{width:.2f}×{height:.2f}cm），规范要求 A4（21.0×29.7cm）")
    else:
        stats["page"] = f"{width:.2f}×{height:.2f}cm"
        errors.append(f"页面尺寸 {width:.2f}×{height:.2f}cm 不是 A4")

    for name, value, expected in (
        ("上边距", section.top_margin.cm, EXPECT_TOP_BOTTOM),
        ("下边距", section.bottom_margin.cm, EXPECT_TOP_BOTTOM),
        ("左边距", section.left_margin.cm, EXPECT_LEFT_RIGHT),
        ("右边距", section.right_margin.cm, EXPECT_LEFT_RIGHT),
    ):
        if not _near(value, expected):
            errors.append(f"{name} {value:.2f}cm，规范为 {expected:.2f}cm")

    header_text = "\n".join(p.text for p in section.header.paragraphs).strip()
    if project:
        if project and project not in header_text:
            warnings.append(f"页眉未找到项目名「{project}」（当前: {header_text or '空'}）")
    footer_xml = section.footer._element.xml
    if "PAGE" not in footer_xml:
        warnings.append("页脚未发现 PAGE 页码域（页码可能缺失）")

    try:
        names = zipfile.ZipFile(path).namelist()
        media = [n for n in names if n.startswith("word/media/") and not n.endswith("/")]
        stats["images"] = len(media)
    except Exception:
        media = []

    inline = len(document.inline_shapes)
    if media and inline and len(media) != inline:
        warnings.append(f"嵌入图片 {len(media)} 张，内联图 {inline} 张（可能有浮动图片，需目视核查）")

    for index, shape in enumerate(document.inline_shapes, start=1):
        try:
            width_cm = shape.width.cm
        except Exception:
            continue
        if width_cm > MAX_IMG_WIDTH_CM + 0.1:
            errors.append(f"第 {index} 张图片宽度 {width_cm:.2f}cm > {MAX_IMG_WIDTH_CM}cm")

    stats["tables"] = len(document.tables)
    for index, table in enumerate(document.tables, start=1):
        tbl_pr = table._tbl.tblPr
        if tbl_pr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders") is None:
            warnings.append(f"第 {index} 个表格缺少 tblBorders，可能不是三线表")
            break

    try:
        normal = document.styles["Normal"]
        rpr = normal.element.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr")
        east_asia = None
        if rpr is not None:
            rfonts = rpr.find(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
            if rfonts is not None:
                east_asia = rfonts.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia")
        if east_asia and east_asia != "宋体":
            warnings.append(f"Normal 样式东亚字体为 {east_asia}，规范为宋体")
        if normal.font.size and normal.font.size != Pt(12):
            warnings.append(f"Normal 字号为 {normal.font.size.pt}pt，规范为 12pt")
    except Exception:
        pass

    for paragraph in document.paragraphs:
        if paragraph.style.name == "Normal" and paragraph.text.strip():
            spacing = paragraph.paragraph_format.line_spacing
            if spacing is not None and not isinstance(spacing, float):
                if abs(spacing.pt - 22) > 0.5:
                    warnings.append(f"首个正文段落行距 {spacing.pt}pt，规范为固定 22 磅")
            break

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
    parser = argparse.ArgumentParser(description="检查 docx 是否符合 Word 格式规范")
    parser.add_argument("paths", nargs="+", help="docx 文件或目录")
    parser.add_argument("--project", default=os.environ.get("DOC_PROJECT_NAME"),
                        help="页眉应包含的项目名")
    args = parser.parse_args(argv)

    files = errors_total = warnings_total = 0
    for path in iter_docx(args.paths):
        files += 1
        errors, warnings, stats = check_docx(path, project=args.project)
        errors_total += len(errors)
        warnings_total += len(warnings)
        try:
            rel = os.path.relpath(path)
        except ValueError:  # Windows 下跨盘符路径无法相对化
            rel = os.path.abspath(path)
        status = "FAIL" if errors else "OK"
        print(f"[{status}] {rel}  (页面 {stats['page']} / 图 {stats['images']} / 表 {stats['tables']})")
        for message in errors:
            print(f"  错误 {message}")
        for message in warnings:
            print(f"  警告 {message}")

    print(f"\n共检查 {files} 个文件：错误 {errors_total}，警告 {warnings_total}")
    return 1 if errors_total else 0


if __name__ == "__main__":
    sys.exit(main())
