"""word-doc-pipeline：把中文 Markdown 文档确定性产出合规 Word（.docx）。

本包只是 `scripts/` 目录的打包外壳：`pip install` 后脚本作为模块安装，
命令行入口见 `word_doc_pipeline.cli`。仓库内直接运行 `python scripts/xxx.py`
的行为完全不变。
"""

from __future__ import annotations

import os
import sys

__version__ = "1.1.0"

# 包内模块之间仍是「裸 import」（import profile / import md2docx），
# 与仓库内 `python scripts/xxx.py` 的加载方式保持一致：
# 把包目录放进 sys.path，两种运行方式共用同一份实现。
_PKG_DIR = os.path.dirname(os.path.abspath(__file__))
if _PKG_DIR not in sys.path:
    sys.path.insert(0, _PKG_DIR)

__all__ = ["__version__"]
