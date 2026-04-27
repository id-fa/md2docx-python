"""md2docx: Markdown を Word (.docx) に変換する CLI ツール (Python 移植版)。"""
from __future__ import annotations

from .config import Config
from .converter import convert_to_docx
from .parser import parse_markdown

__all__ = ["Config", "convert_to_docx", "parse_markdown"]
__version__ = "0.1.0"
