# -*- coding: utf-8 -*-
"""批量转换 docs/ 目录下的 Markdown 并做格式校验。

用法:
    python scripts/build_docs.py <docs目录> [--project "页眉项目名"]
                                  [--include "*.md"] [--skip-suffix 绘制说明.md 核查报告.md]
                                  [--dry-run] [--no-check]

行为:
- 递归查找 docs 目录下的 .md，跳过 资料/ 附件/ .git/ node_modules/ 目录
- 跳过文件名以「绘制说明.md」「核查报告.md」结尾的工作笔记（可用 --skip-suffix 覆盖）
- 在 md 同目录生成同名 .docx
- 生成后调用 check_docx 校验，汇总输出
退出码：0 全部成功；1 有转换或校验失败。
"""
from __future__ import annotations

import argparse
import fnmatch
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import check_docx  # noqa: E402
import md2docx  # noqa: E402

SKIP_DIRS = {".git", "node_modules", "资料", "附件"}
DEFAULT_SKIP_SUFFIX = ("绘制说明.md", "核查报告.md")


def find_markdown(docs_dir, include="*.md", skip_suffix=DEFAULT_SKIP_SUFFIX):
    for root, dirs, files in os.walk(docs_dir):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith(".")]
        for name in sorted(files):
            if not fnmatch.fnmatch(name, include):
                continue
            if any(name.endswith(suffix) for suffix in skip_suffix):
                continue
            yield os.path.join(root, name)


def main(argv=None):
    parser = argparse.ArgumentParser(description="批量转换并校验 docs 下的文档")
    parser.add_argument("docs", help="docs 目录（或任意包含 md 的目录）")
    parser.add_argument("--project", default=os.environ.get("DOC_PROJECT_NAME"),
                        help="页眉项目名")
    parser.add_argument("--include", default="*.md", help="匹配的 md 文件模式")
    parser.add_argument("--skip-suffix", nargs="*", default=list(DEFAULT_SKIP_SUFFIX),
                        help="跳过这些后缀的工作笔记")
    parser.add_argument("--dry-run", action="store_true", help="只列出将转换的文件")
    parser.add_argument("--no-check", action="store_true", help="转换后不校验")
    args = parser.parse_args(argv)

    if not os.path.isdir(args.docs):
        print(f"错误: 目录不存在: {args.docs}", file=sys.stderr)
        return 1

    md_files = list(find_markdown(args.docs, args.include, tuple(args.skip_suffix)))
    if not md_files:
        print("未找到匹配的 Markdown 文件")
        return 0
    if args.dry_run:
        for path in md_files:
            print(path)
        print(f"\n共 {len(md_files)} 个文件（dry-run，未转换）")
        return 0

    converted, failed = 0, 0
    check_errors = 0
    for md_path in md_files:
        docx_path = os.path.splitext(md_path)[0] + ".docx"
        try:
            shown = os.path.relpath(md_path)
        except ValueError:  # Windows 下跨盘符路径无法相对化
            shown = os.path.abspath(md_path)
        print(f"转换: {shown}")
        try:
            md2docx._WARNED.clear()
            md2docx.convert(md_path, docx_path, project=args.project)
            converted += 1
        except Exception as exc:
            failed += 1
            print(f"  失败: {exc}", file=sys.stderr)
            continue
        if not args.no_check:
            errors, warnings, stats = check_docx.check_docx(
                docx_path, project=args.project)
            check_errors += len(errors)
            for message in errors:
                print(f"  校验错误: {message}")
            for message in warnings:
                print(f"  校验警告: {message}")

    print(f"\n完成：转换 {converted} 个，失败 {failed} 个，校验错误 {check_errors} 个")
    return 1 if (failed or check_errors) else 0


if __name__ == "__main__":
    sys.exit(main())
