# -*- coding: utf-8 -*-
"""模板接管：把变量与 Markdown 正文填进既有 .docx 模板。

适用场景：实验报告、课程设计说明书、NSFC 申请书等「没有统一排版标准、靠
学校/单位模板承载版式」的文档类型。这类文档的页面设置、页眉页脚、封面、
表格框架都在模板里，脚本只负责两件事：

1. 替换 ``{{占位符}}``（正文段落、表格单元格、页眉页脚都支持）；
2. 把 Markdown 正文渲染后插入 ``{{正文}}`` 锚点所在位置。

用法::

    python scripts/fill_template.py 模板.docx 正文.md 输出.docx \
        [--profile tech-doc] [--vars 变量.toml] [--var 键=值 ...] [--list] [--strict]

变量来源（后者覆盖前者）：
- 变量文件 ``[vars]`` 表里的字符串，或文件顶层直接写的字符串键；
- ``--var 键=值``（可重复）；
- 内置变量 ``today``（2026-11-15 这样的日期）与 ``文件名``（正文 md 的文件名）。

模板里没找到 ``{{正文}}`` 锚点时，正文追加到文末并告警。
正文排版仍按 profile（字体/字号/行距/缩进/三线表/图注），模板只提供页面与
页眉页脚；页面设置不做任何改动，也不写页码域、不生成目录。
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import tomllib
from datetime import date

from docx import Document

import md2docx
import profile as profile_mod

# {{ 名字 }}，名字里不允许再出现花括号，允许中文与空格
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

# 正文锚点：模板里写 {{正文}} 或 {{content}}，渲染后该段落被正文替换
ANCHOR_NAMES = ("正文", "content")

# 这些占位符由脚本自己消费，不当成未填项报出来
_RESERVED = set(ANCHOR_NAMES)


def warn(msg):
    print(f"警告: {msg}", file=sys.stderr)


def note(msg):
    print(f"提示: {msg}", file=sys.stderr)


def _iter_table_paragraphs(tables):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph
                yield from _iter_table_paragraphs(cell.tables)


def iter_body_paragraphs(document):
    """正文段落 + 表格（含嵌套）单元格段落，不含页眉页脚。"""
    for paragraph in document.paragraphs:
        yield paragraph
    yield from _iter_table_paragraphs(document.tables)


def iter_paragraphs(document):
    """正文段落 + 表格（含嵌套）单元格段落 + 页眉页脚段落。"""
    yield from iter_body_paragraphs(document)

    for section in document.sections:
        for hf in (section.header, section.footer,
                   section.even_page_header, section.even_page_footer,
                   section.first_page_header, section.first_page_footer):
            try:
                if hf.is_linked_to_previous:
                    continue
            except Exception:  # 极少数损坏的模板取不到定义
                continue
            for paragraph in hf.paragraphs:
                yield paragraph
            yield from _iter_table_paragraphs(hf.tables)


def _run_spans(paragraph):
    """返回 [(run, start, end)]，按 runs 拼接后的字符偏移。"""
    spans, pos = [], 0
    for run in paragraph.runs:
        text = run.text or ""
        spans.append((run, pos, pos + len(text)))
        pos += len(text)
    return spans


def replace_in_paragraph(paragraph, mapping, unresolved):
    """在单个段落内替换占位符，保留占位符首尾所在 run 的格式。

    占位符可能被 Word 拆到多个 run 里（编辑过后很常见），所以按拼接文本
    定位，再回写到跨度的第一个 run 上。
    """
    text = "".join(run.text or "" for run in paragraph.runs)
    if "{{" not in text:
        return
    matches = [m for m in PLACEHOLDER_RE.finditer(text)]
    if not matches:
        return

    for match in reversed(matches):  # 从后往前替换，偏移不会互相影响
        name = match.group(1).strip()
        if name in mapping:
            value = str(mapping[name])
        elif name in _RESERVED:
            continue
        else:
            unresolved.add(name)
            continue

        spans = _run_spans(paragraph)
        start, end = match.start(), match.end()
        first = last = None
        for run, run_start, run_end in spans:
            if first is None and run_end > start:
                first = (run, run_start)
            if run_start < end:
                last = (run, run_start, run_end)
        if first is None or last is None:  # 理论上不会发生
            continue

        first_run, first_start = first
        last_run, last_start, last_end = last
        prefix = (first_run.text or "")[:start - first_start]
        suffix = (last_run.text or "")[end - last_start:]
        if first_run is last_run:
            first_run.text = prefix + value + suffix
        else:
            first_run.text = prefix + value
            last_run.text = suffix
            # 中间被占位符跨过的 run 清空，避免残留片段
            for run, run_start, run_end in spans:
                if run is not first_run and run is not last_run \
                        and run_start >= first_start and run_end <= last_end:
                    run.text = ""


def find_anchor(document):
    """返回第一个含 {{正文}} / {{content}} 的段落（只看正文与表格单元格）。

    页眉页脚不参与：把 {{正文}} 写在页眉里多半是笔误，若当锚点会把整篇正文
    注入页眉。
    """
    for paragraph in iter_body_paragraphs(document):
        for match in PLACEHOLDER_RE.finditer(paragraph.text or ""):
            if match.group(1).strip() in ANCHOR_NAMES:
                return paragraph
    return None


def collect_placeholders(document):
    names = []
    for paragraph in iter_paragraphs(document):
        for match in PLACEHOLDER_RE.finditer(paragraph.text or ""):
            name = match.group(1).strip()
            if name not in names:
                names.append(name)
    return names


def load_vars(path):
    if not path:
        return {}
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError as exc:
        raise SystemExit(f"错误: 读不到变量文件 {path}: {exc}")
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]   # 记事本默认写 BOM，它不是 TOML 内容，直接去掉
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise SystemExit(f"错误: 变量文件不是 UTF-8 编码（{path}）: {exc}"
                         "（请另存为「UTF-8 无 BOM」）")
    try:
        data = tomllib.loads(text)
    except tomllib.TOMLDecodeError as exc:
        raise SystemExit(f"错误: 变量文件不是合法 TOML（{path}）: {exc}"
                         "（中文键名要加英文引号，例如 \"姓名\" = \"张三\"）")
    values = {}
    for key, value in data.items():
        if isinstance(value, str):
            values[key] = value
    nested = data.get("vars")
    if isinstance(nested, dict):
        for key, value in nested.items():
            if isinstance(value, (str, int, float, bool)):
                values[key] = value
    values.pop("vars", None)
    return values


def parse_var_args(pairs):
    values = {}
    for item in pairs or []:
        if "=" not in item:
            raise SystemExit(f"错误: --var 需要「键=值」形式，收到 {item!r}")
        key, _, value = item.partition("=")
        values[key.strip()] = value
    return values


def build_parser():
    parser = argparse.ArgumentParser(
        description="模板接管：变量替换 + Markdown 正文注入既有 .docx 模板")
    parser.add_argument("template", help="模板 .docx 路径")
    parser.add_argument("md", nargs="?", help="正文 Markdown 路径（可省略＝只填变量）")
    parser.add_argument("out", nargs="?", help="输出 .docx 路径（省略则原地覆盖副本）")
    parser.add_argument("--profile", default=None,
                        help="正文排版 profile（默认 ug-thesis-bit）")
    parser.add_argument("--vars", default=None, help="变量文件（TOML）")
    parser.add_argument("--var", action="append", default=[], metavar="键=值",
                        help="命令行变量，可重复，覆盖变量文件")
    parser.add_argument("--list", action="store_true",
                        help="只列出模板里的占位符，不生成文件")
    parser.add_argument("--strict", action="store_true",
                        help="存在未填占位符时以退出码 1 结束")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    if not os.path.exists(args.template):
        print(f"错误: 找不到模板 {args.template}", file=sys.stderr)
        return 2

    document = Document(args.template)

    if args.list:
        names = collect_placeholders(document)
        if not names:
            print("模板里没有 {{占位符}}")
        else:
            print(f"模板占位符 {len(names)} 个：")
            for name in names:
                print(f"  {{{{{name}}}}}")
        return 0

    if not args.out:
        print("错误: 需要输出路径（第三个位置参数）", file=sys.stderr)
        return 2

    anchor = find_anchor(document)

    mapping = {}
    mapping.update(load_vars(args.vars))
    mapping.update(parse_var_args(args.var))
    mapping["today"] = date.today().isoformat()
    if args.md:
        mapping["文件名"] = os.path.splitext(os.path.basename(args.md))[0]

    unresolved = set()
    for paragraph in iter_paragraphs(document):
        replace_in_paragraph(paragraph, mapping, unresolved)

    if unresolved:
        listed = "、".join("{{%s}}" % n for n in sorted(unresolved))
        warn(f"模板里还有 {len(unresolved)} 个占位符没填：{listed}"
             "（用 --var 键=值 或 --vars 文件补上）")

    if args.md:
        if not os.path.exists(args.md):
            print(f"错误: 找不到正文 {args.md}", file=sys.stderr)
            return 2
        if anchor is None:
            warn("模板里没找到 {{正文}} 锚点，正文已追加到文末"
                 "（建议在模板里想要插入正文的位置写一个只含 {{正文}} 的段落）")
        md2docx.convert(args.md, args.out, profile=args.profile,
                        template_doc=document, anchor=anchor)
    else:
        document.save(args.out)
        print(f"已生成: {args.out}  [模板接管·仅填变量]")

    if args.strict and unresolved:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
