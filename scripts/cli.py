"""pip 安装后的命令行入口（console scripts，`wdp-*`）。

仓库内开发时不需要它：直接 `python scripts/md2docx.py ...` 即可。
安装后每个入口只是把 `sys.argv[1:]` 交给对应模块的 `main()`，
退出码原样返回（argparse 的错误、`ProfileError` 的 rc=2 等行为一致）。
"""

from __future__ import annotations

import importlib
import sys


def _run(module_name, argv=None):
    module = importlib.import_module(module_name)
    return module.main(sys.argv[1:] if argv is None else argv)


def profile():
    return _run("profile")


def md2docx():
    return _run("md2docx")


def check_md():
    return _run("check_md")


def check_docx():
    return _run("check_docx")


def build_docs():
    return _run("build_docs")


def fill_template():
    return _run("fill_template")


def selftest():
    module = importlib.import_module("selftest")
    return module.main()
