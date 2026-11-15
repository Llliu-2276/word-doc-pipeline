# -*- coding: utf-8 -*-
"""Word 成品格式检查（期望值全部来自同一份 profile）。

用法:
    python scripts/check_docx.py <docx文件或目录>... [--profile tech-doc] [--project "页眉项目名"]

检查项（与 md2docx.py 使用同一份 profiles/*.toml，因此检查标准与生成标准不会漂移）:
- 纸张与四边页边距、装订线（page.size / page.margins / page.gutter）、镜像边距（page.mirror）
- 页眉文字（header_footer.header_text 或 --project）
- 页码域：PAGE（cn 格式还要 NUMPAGES），以及页码位置所在部件
- 图片：数量、单张宽度 ≤ body.max_image_width、缺失图片时提示
- 表格：数量、三线表 tblBorders、表头下框线
- 正文：Normal 样式东亚字体/字号、首个正文段落行距与缩进
- 标题：各级标题的字体、字号、加粗
- 目录：profile 要求 structure.toc 时检查 TOC 域是否存在

「缺件」默认按错误处理（页眉、页码、三线表边框、嵌入图数、TOC 域）。
md2docx.py 会把生成时的开关写进 docProps/custom.xml（wdp_profile / wdp_header /
wdp_footer / wdp_toc …），本脚本据此区分「profile 要求而缺失」（错误）与
「生成时用 --no-header / --no-footer / --no-toc 显式关闭」（跳过）；
没有生成标记的旧文档只能严格判定，并给一条提示建议重新生成。

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
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import profile as profile_mod  # noqa: E402

TOL_CM = 0.15
TOL_PT = 0.5

_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

CUSTOM_PROPS_PART = "/docProps/custom.xml"


def read_generation_tags(document):
    """读 md2docx.py 写入的生成标记（docProps/custom.xml 的 wdp_* 属性）。

    返回 dict（键去掉 wdp_ 前缀）；旧版脚本生成或手工文档没有该部件，返回 None。
    """
    try:
        parts = list(document.part.package.iter_parts())
    except Exception:
        return None
    for part in parts:
        if str(part.partname) != CUSTOM_PROPS_PART:
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            return None
        tags = {}
        for prop in root:
            name = prop.get("name") or ""
            if name.startswith("wdp_"):
                tags[name[4:]] = "".join(prop.itertext()).strip()
        return tags or None
    return None


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


def _font_candidates(cfg, expected):
    """profile 指定的字体 + fonts.fallback 链：缺字体时脚本会替换成链上第一个可用字体。"""
    fallback = (cfg.get("fonts") or {}).get("fallback") or {}
    names = [expected]
    for name in fallback.get(expected, []):
        if name not in names:
            names.append(name)
    return names


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

    # ---------------- 生成标记（区分「缺失」与「生成时显式关闭」）
    tags = read_generation_tags(document)
    if tags is None:
        warnings.append("未找到生成标记（旧版脚本或手工改过的文档）：按 profile 严格判定，"
                        "可用当前 md2docx.py 重新生成以区分「缺失」与「显式关闭」")
    elif tags.get("mode") == "template":
        # 模板接管：页面、页眉页脚、表格边框、目录全部由模板决定，比不了 profile
        stats["mode"] = "template"
        stats["page"] = "模板"
        return errors, warnings, stats
    else:
        generated_by = tags.get("profile") or "?"
        stats["generated_by"] = generated_by
        if generated_by != cfg.get("name"):
            errors.append(f"文档由 profile「{generated_by}」生成，本次校验用的是"
                          f"「{cfg.get('name')}」——两者期望值不同，先统一 profile 再校验")

    def enabled(name, default=True):
        """生成时是否启用了该项；没有生成标记的旧文档按「启用」处理。"""
        if tags is None:
            return default
        return tags.get(name, "1") != "0"

    # ---------------- 页面与页边距
    page = cfg["page"]
    # md2docx 允许 --page-size 覆盖纸张，覆盖值记在生成标记里；有标记就按标记校验，
    # 否则「B5 生成 + A4 profile」会被误判为页面不符
    tag_size = (tags or {}).get("page_size")
    expected_key = tag_size if tag_size in profile_mod.PAGE_SIZES else page["size"]
    expected_w, expected_h = profile_mod.PAGE_SIZES.get(
        expected_key, profile_mod.PAGE_SIZES["A4"])
    size_note = "" if expected_key == page["size"] else \
        f"生成时用 --page-size {expected_key} 覆盖了 profile 的 {page['size']}"
    for index, section in enumerate(document.sections, start=1):
        width, height = section.page_width.cm, section.page_height.cm
        label = "" if len(document.sections) == 1 else f"第 {index} 节"
        if _near(width, expected_w) and _near(height, expected_h):
            stats["page"] = expected_key
        else:
            stats["page"] = f"{width:.2f}×{height:.2f}cm"
            report(page, f"{label}页面 {width:.2f}×{height:.2f}cm，"
                         f"应为 {expected_key}"
                         f"（{expected_w:.2f}×{expected_h:.2f}cm）"
                         + (f"；{size_note}" if size_note else
                            f"；profile「{cfg.get('name')}」要求 {page['size']}"))
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
        gutter = float(page.get("gutter") or 0.0)
        if abs(section.gutter.cm - gutter) > 0.05:
            report(page, f"{label}装订线 {section.gutter.cm:.2f}cm，"
                         f"profile 要求 {gutter:.2f}cm")

    if page.get("mirror"):
        if "mirrorMargins" not in document.settings.element.xml:
            report(page, "profile 要求奇偶页镜像边距（page.mirror = true），"
                         "文档未设置 w:mirrorMargins")

    # ---------------- 文档网格（grid 行距模式：每页行数 × 每行字数）
    body_cfg = cfg["body"]
    if body_cfg["line_spacing"]["mode"] == "grid":
        grid = body_cfg.get("grid") or {}
        for index, sec in enumerate(document.sections, start=1):
            el = sec._sectPr.find(qn("w:docGrid"))
            if el is None or el.get(qn("w:type")) not in ("lines", "linesAndChars"):
                errors.append(f"第 {index} 节未设置文档网格 w:docGrid"
                              "（profile 行距为 grid 模式）")
                continue
            height = (sec.page_height.twips - sec.top_margin.twips
                      - sec.bottom_margin.twips)
            expected = height / grid["lines"] if grid.get("lines") else None
            pitch = int(el.get(qn("w:linePitch")) or 0)
            if expected and abs(pitch - expected) > 2:
                report(body_cfg, f"第 {index} 节文档网格 linePitch={pitch}，"
                                 f"与每页 {grid['lines']} 行不符（应约 {expected:.0f}）")

    section = document.sections[0]
    # 分节文档（目录 + 正文）的页码/页眉在正文节上，前置节可能刻意留空
    multi_section = len(document.sections) > 1
    body_section = document.sections[-1] if multi_section else section
    hf = cfg["header_footer"]

    # ---------------- 页眉（生成时用了 --no-header 则豁免）
    header_text = "\n".join(p.text for p in body_section.header.paragraphs).strip()
    expected_header = hf.get("header_text") or project
    if expected_header and enabled("header") and hf.get("header_enabled", True):
        if expected_header not in header_text:
            errors.append(f"页眉未找到「{expected_header}」（当前: {header_text or '空'}）；"
                          "若确实不要页眉，请用 --no-header 生成，校验才会识别为「显式关闭」")

    # ---------------- 页码（生成时用了 --no-footer 则豁免）
    fmt = hf["page_number"].get("format", "dash")
    position = hf["page_number"].get("position", "footer-center")
    outside = position == "odd-even-outside"
    part = body_section.header if position == "header-right" else body_section.footer
    page_flag = "header" if position == "header-right" else "footer"
    if enabled(page_flag):
        parts = [part]
        if outside:
            if "evenAndOddHeaders" not in document.settings.element.xml:
                errors.append("页码要求奇偶页外侧（odd-even-outside），"
                              "但文档未启用 w:evenAndOddHeaders 奇偶页页脚")
            parts.append(body_section.even_page_footer)
        for one in parts:
            part_xml = one._element.xml
            if "PAGE" not in part_xml:
                errors.append(f"页码域缺失（profile 要求 {fmt} 格式，位置 {position}）")
                break
        else:
            if fmt == "cn" and "NUMPAGES" not in part._element.xml:
                errors.append("页码为 cn 格式（第 X 页 共 Y 页），但缺少 NUMPAGES 域")
            part_text = "\n".join(p.text for p in part.paragraphs).strip()
            if position == "header-right":
                # 页眉右置页码与页眉项目名在同一段：页码在制表符之后，只看最后一段
                part_text = part_text.split("\n")[-1].rsplit("\t", 1)[-1].strip()
            if fmt == "dash" and part_text and not part_text.startswith("-"):
                errors.append(f"页码格式应为「- 1 -」，实际为「{part_text}」")
            if fmt == "emdash" and part_text and not part_text.startswith("—"):
                errors.append(f"页码格式应为「— 1 —」，实际为「{part_text}」")

    # ---------------- 图片
    try:
        names = zipfile.ZipFile(path).namelist()
        media = [n for n in names if n.startswith("word/media/") and not n.endswith("/")]
        stats["images"] = len(media)
    except Exception:
        media = []

    inline = len(document.inline_shapes)
    if media and inline and len(media) != inline:
        errors.append(f"嵌入图片 {len(media)} 张，内联图 {inline} 张——有图片不在正文内联位置"
                      "（浮动图片或多出的页眉图），需目视核查并改为内联")

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
    style = cfg["table"].get("style", "three_line")
    for index, table in enumerate(document.tables, start=1):
        tbl_pr = table._tbl.tblPr
        borders = tbl_pr.find(_W_NS + "tblBorders") if tbl_pr is not None else None
        if borders is None:
            errors.append(f"第 {index} 个表格缺少 tblBorders，不是"
                          + ("全线表" if style == "grid" else "三线表"))
            continue   # 继续查下一个表格，别让第一处不合格掩盖其余问题
        if style == "grid":
            missing = [e for e in ("top", "left", "bottom", "right", "insideH", "insideV")
                       if borders.find(qn("w:" + e)) is None
                       or borders.find(qn("w:" + e)).get(qn("w:val")) == "nil"]
            if missing:
                errors.append(f"第 {index} 个表格应为全线表，但缺少/关闭了框线: "
                              + "、".join(missing))
                continue
        if not table.rows:
            errors.append(f"第 {index} 个表格没有行，无法检查表头框线")
            continue
        first_cell = table.rows[0].cells[0]
        tc_pr = first_cell._tc.find(qn("w:tcPr"))
        if tc_pr is None or tc_pr.find(qn("w:tcBorders")) is None:
            errors.append(f"第 {index} 个表格表头无下框线，"
                          + ("全线表不完整" if style == "grid" else "三线表不完整"))
            continue

    # ---------------- 正文
    body = cfg["body"]
    east_asia = _style_east_asia(document, "Normal")
    if east_asia and east_asia not in _font_candidates(cfg, body["font_cn"]):
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
            indent_spec = body.get("first_line_indent") or {}
            mode = indent_spec.get("mode", "first_line")
            chars = int(indent_spec.get("chars") or 0)
            if mode != "none" and chars:
                pPr = paragraph._p.pPr
                ind = pPr.find(qn("w:ind")) if pPr is not None else None
                key = "w:hangingChars" if mode == "hanging" else "w:firstLineChars"
                actual = ind.get(qn(key)) if ind is not None else None
                if actual is None or int(float(actual)) != chars * 100:
                    label = "悬挂缩进" if mode == "hanging" else "首行缩进"
                    report(body, f"首个正文段落{label}为 {actual or '未设置'}，"
                                 f"profile 要求 {chars} 字符（{chars * 100}，单位 0.01 字符）")
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
        if font and font not in _font_candidates(cfg, spec["font"]):
            report(spec, f"{level} 级标题字体为 {font}，profile 要求 {spec['font']}")
        if bool(run.font.bold) != bool(spec["bold"]):
            report(spec, f"{level} 级标题加粗为 {run.font.bold}，"
                         f"profile 要求 {spec['bold']}")

    # ---------------- 目录（生成时用了 --no-toc 则豁免，手工插目录属正常）
    if cfg["structure"].get("toc") and enabled("toc"):
        body_xml = document.element.body.xml
        if "TOC \\o" not in body_xml and " TOC " not in body_xml:
            errors.append("profile 要求自动目录，但文档中未发现 TOC 域；"
                          "若目录是手工插入的，请用 --no-toc 生成以便校验识别")

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
        if stats.get("mode") == "template":
            print(f"[{status}] {rel}  (profile {stats['profile']} / "
                  "模板接管：版式由模板决定，未做版式校验)")
            for message in errors:
                print(f"  错误 {message}")
            for message in warnings:
                print(f"  警告 {message}")
            continue
        source = stats.get("generated_by")
        origin = f" / 生成自 {source}" if source else " / 无生成标记"
        print(f"[{status}] {rel}  (profile {stats['profile']} / 页面 {stats['page']} / "
              f"图 {stats['images']} / 表 {stats['tables']}{origin})")
        for message in errors:
            print(f"  错误 {message}")
        for message in warnings:
            print(f"  警告 {message}")

    print(f"\n共检查 {files} 个文件：错误 {errors_total}，警告 {warnings_total}")
    return 1 if errors_total else 0


if __name__ == "__main__":
    sys.exit(main())
