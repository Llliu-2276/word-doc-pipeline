# -*- coding: utf-8 -*-
"""Markdown 写作规范检查（针对 md2docx.py 支持子集）。

用法:
    python scripts/check_md.py <md文件或目录>...

检查项:
- 禁用语法：代码块、####、超链接、嵌套列表、引用、HTML、行内代码、删除线、任务列表、脚注
- 图片：路径是否存在、图注是否形如「图 x-x 名称」、PNG 宽度是否 ≥1600px
- 图表编号：重复编号、按章编号是否从 1 递增
- 表格：上方是否有「**表 x-x 说明**」表注（表注与表格之间允许空行）
- 标题层级是否跳级
- 目录扫描时默认跳过 *绘制说明.md、*核查报告.md 等工作笔记（--all 可包含）

退出码：0 通过（可含警告）；1 存在错误。
"""
from __future__ import annotations

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import profile as profile_mod  # noqa: E402

# (级别, 规则名, 正则, 说明)
RULES = [
    ("ERROR", "代码块", r"^\s*```", "不支持代码块，请改为文字描述或表格"),
    ("ERROR", "四级标题", r"^####\s", "不支持 ####，请用 **4.4.1 标题** 加粗段落"),
    ("ERROR", "超链接", r"(?<!!)\[[^\]]+\]\([^)]+\)", "不支持链接，转换后仅剩文字，请去掉 URL"),
    ("ERROR", "嵌套列表", r"^\s{2,}(?:[-*+]|\d+\.)\s", "不支持嵌套列表，请改为单层列表或三线表"),
    ("ERROR", "HTML 标签", r"</?[a-zA-Z][^>]*>", "不支持 HTML，请用 Markdown 子集"),
    ("ERROR", "行内代码", r"`[^`\n]+`", "不支持行内代码反引号"),
    ("WARN", "引用块", r"^\s*>\s", "不支持 > 引用，转换时会去掉前缀"),
    ("WARN", "删除线", r"~~[^~\n]+~~", "不支持删除线"),
    ("WARN", "任务列表", r"^\s*[-*+]\s+\[[ xX]\]", "不支持任务列表，请改为普通列表"),
    ("WARN", "脚注", r"\[\^[^\]]+\]", "不支持脚注"),
    ("WARN", "斜体", r"(?<![*\w])\*[^*\n]+\*(?![*\w])", "不支持斜体，请去掉 * 号"),
]

IMG_RE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$")
# 按章编号：图 4-1 / 表 4-1；连续编号：图 1 / 表 1
FIG_CAP_CHAPTER = re.compile(r"^图\s*(\d+|[A-Z])\s*[-–—]\s*(\d+)")
TAB_CAP_CHAPTER = re.compile(r"^\*\*表\s*(\d+|[A-Z])\s*[-–—]\s*(\d+)")
FIG_CAP_CONTINUOUS = re.compile(r"^图\s*(\d+)(?!\s*[-–—]\s*\d)(?![0-9])")
TAB_CAP_CONTINUOUS = re.compile(r"^\*\*表\s*(\d+)(?!\s*[-–—]\s*\d)(?![0-9])")
TABLE_SEP_RE = re.compile(r"^\|[\s:\-|]+\|?$")

# 由 main() 按 profile 设置：{"figure": "by_chapter"|"continuous", ...}
NUMBERING = {"figure": "by_chapter", "table": "by_chapter"}


def fig_cap_re():
    return FIG_CAP_CHAPTER if NUMBERING["figure"] == "by_chapter" else FIG_CAP_CONTINUOUS


def tab_cap_re():
    return TAB_CAP_CHAPTER if NUMBERING["table"] == "by_chapter" else TAB_CAP_CONTINUOUS

_IMAGE_WIDTH_CACHE = {}


def image_width(path):
    """返回图片像素宽度；Pillow 不可用或读取失败时返回 None。"""
    if path in _IMAGE_WIDTH_CACHE:
        return _IMAGE_WIDTH_CACHE[path]
    width = None
    try:
        from PIL import Image
        with Image.open(path) as im:
            width = im.size[0]
    except Exception:
        width = None
    _IMAGE_WIDTH_CACHE[path] = width
    return width


def _record_number(kind, chapter, number, idx, counter, last, errors, warnings):
    """记录编号，检查重复与不连续；连续编号时 chapter 为 None。"""
    label = f"{kind} {chapter}-{number}" if chapter is not None else f"{kind} {number}"
    key = (chapter, number)
    if key in counter:
        errors.append((idx, kind + "号", f"{label} 与第 {counter[key]} 行重复"))
    counter[key] = idx
    if last is not None and last[0] == chapter and number != last[1] + 1:
        prev = f"{kind} {chapter}-{last[1]}" if chapter is not None else f"{kind} {last[1]}"
        warnings.append((idx, kind + "号", f"{label} 前一个是 {prev}，编号不连续"))
    return (chapter, number)


def scan_file(md_path):
    """扫描单个 md，返回 (errors, warnings, stats)。"""
    errors, warnings = [], []
    stats = {"figures": 0, "tables": 0, "images": 0}
    base = os.path.dirname(os.path.abspath(md_path))
    with open(md_path, encoding="utf-8") as f:
        lines = f.read().splitlines()

    prev_level = 0
    fig_numbers = {}
    tab_numbers = {}
    last_fig = None
    last_tab = None
    by_chapter_fig = NUMBERING["figure"] == "by_chapter"
    by_chapter_tab = NUMBERING["table"] == "by_chapter"

    for idx, line in enumerate(lines, start=1):
        stripped = line.strip()

        # 图片：单独处理（避免被通用规则误报）
        m_img = IMG_RE.match(stripped)
        if m_img:
            stats["images"] += 1
            alt, src = m_img.group(1).strip(), m_img.group(2).strip()
            if re.match(r"^[a-zA-Z]+://", src):
                errors.append((idx, "图片", f"不支持网络图片地址: {src}"))
            else:
                img_path = src if os.path.isabs(src) else os.path.join(base, src)
                if not os.path.isfile(img_path):
                    errors.append((idx, "图片", f"图片不存在: {src}"))
                else:
                    width = image_width(img_path)
                    if width is not None and width < 1600:
                        warnings.append(
                            (idx, "图片宽度", f"{src} 宽度 {width}px < 1600px，打印可能模糊"))
            m_cap = fig_cap_re().match(alt)
            if not m_cap:
                sample = "图 4-1 名称" if by_chapter_fig else "图 1 名称"
                warnings.append((idx, "图注", f"图注「{alt}」应形如「{sample}」"))
            else:
                stats["figures"] += 1
                if by_chapter_fig:
                    raw = m_cap.group(1)
                    chapter = raw if raw.isalpha() else int(raw)
                    number = int(m_cap.group(2))
                else:
                    chapter, number = None, int(m_cap.group(1))
                last_fig = _record_number("图", chapter, number, idx, fig_numbers,
                                          last_fig, errors, warnings)
            continue

        # 表注：记录编号并核对下一个非空行是表格
        m_tab = tab_cap_re().match(stripped)
        if m_tab:
            if by_chapter_tab:
                raw = m_tab.group(1)
                chapter = raw if raw.isalpha() else int(raw)
                number = int(m_tab.group(2))
            else:
                chapter, number = None, int(m_tab.group(1))
            j = idx
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines) and not lines[j].strip().startswith("|"):
                warnings.append((idx, "表注", "表注下方（跳过空行）不是表格，请紧贴表格上方"))
            last_tab = _record_number("表", chapter, number, idx, tab_numbers,
                                      last_tab, errors, warnings)
            continue

        # 表格起始：记录表数并检查上方（跳过空行）是否有表注
        if stripped.startswith("|") and idx < len(lines) and TABLE_SEP_RE.match(lines[idx].strip()):
            stats["tables"] += 1
            k = idx - 2
            while k >= 0 and not lines[k].strip():
                k -= 1
            if k < 0 or not tab_cap_re().match(lines[k].strip()):
                warnings.append((idx, "表注", "表格上方缺少「**表 x-x 说明**」表注"))
            continue

        # 标题层级
        m_head = re.match(r"^(#{1,6})\s", stripped)
        if m_head:
            level = len(m_head.group(1))
            if level > 3:
                errors.append((idx, "标题", f"{'#' * level} 级标题不支持（最多三级）"))
            elif prev_level and level > prev_level + 1:
                warnings.append((idx, "标题", f"标题从 {prev_level} 级跳到 {level} 级"))
            prev_level = level
            continue

        # 通用禁用语法
        for level, name, pattern, hint in RULES:
            if re.search(pattern, line):
                message = f"{name}: {hint}"
                if level == "ERROR":
                    errors.append((idx, name, message))
                else:
                    warnings.append((idx, name, message))

    if not lines:
        warnings.append((0, "空文件", "文件为空"))
    return errors, warnings, stats


DEFAULT_SKIP_SUFFIX = ("绘制说明.md", "核查报告.md")


def iter_md(paths, scan_all=False):
    for path in paths:
        if os.path.isdir(path):
            for root, _, files in os.walk(path):
                if os.path.basename(root) in {".git", "node_modules", "资料", "附件"}:
                    continue
                for name in sorted(files):
                    if not name.endswith(".md"):
                        continue
                    if not scan_all and any(name.endswith(s) for s in DEFAULT_SKIP_SUFFIX):
                        continue
                    yield os.path.join(root, name)
        else:
            yield path


def main(argv=None):
    parser = argparse.ArgumentParser(description="检查 Markdown 是否符合转换子集")
    parser.add_argument("paths", nargs="+", help="md 文件或目录")
    parser.add_argument("--profile", default=None,
                        help="排版 profile 名或 .toml 路径（决定图/表按章还是连续编号）")
    parser.add_argument("--all", action="store_true",
                        help="目录扫描时也检查 绘制说明/核查报告 等工作笔记")
    args = parser.parse_args(argv)

    try:
        cfg, notices = profile_mod.load_profile(
            args.profile or profile_mod.DEFAULT_PROFILE, check_fonts=False)
    except profile_mod.ProfileError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    for notice in notices:
        print(f"提示: {notice}", file=sys.stderr)
    numbering = cfg["numbering"]
    NUMBERING["figure"] = numbering["figure"]["scheme"]
    NUMBERING["table"] = numbering["table"]["scheme"]

    total_errors = total_warnings = files = 0
    for md_path in iter_md(args.paths, scan_all=args.all):
        files += 1
        errors, warnings, stats = scan_file(md_path)
        total_errors += len(errors)
        total_warnings += len(warnings)
        try:
            rel = os.path.relpath(md_path)
        except ValueError:  # Windows 下跨盘符路径无法相对化
            rel = os.path.abspath(md_path)
        status = "FAIL" if errors else "OK"
        print(f"[{status}] {rel}  (图 {stats['figures']} / 表 {stats['tables']} / 图片 {stats['images']})")
        for idx, name, message in errors:
            print(f"  错误 {rel}:{idx} [{name}] {message}")
        for idx, name, message in warnings:
            print(f"  警告 {rel}:{idx} [{name}] {message}")

    print(f"\n共检查 {files} 个文件：错误 {total_errors}，警告 {total_warnings}")
    return 1 if total_errors else 0


if __name__ == "__main__":
    sys.exit(main())
