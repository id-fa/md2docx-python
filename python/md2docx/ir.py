"""中間表現 (IR): Markdown をパースした結果を表すブロック/インライン要素。

Rust 版 `src/ir.rs` の Python 移植。docx-rs / python-docx などの出力ライブラリには
依存せず、parser と converter を分離するための DTO 群。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Union


class Alignment(Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    NONE = "none"


@dataclass
class Text:
    value: str


@dataclass
class Code:
    value: str


@dataclass
class Bold:
    children: list["Inline"]


@dataclass
class Italic:
    children: list["Inline"]


@dataclass
class Link:
    text: list["Inline"]
    url: str


@dataclass
class SoftBreak:
    pass


@dataclass
class HardBreak:
    pass


Inline = Union[Text, Code, Bold, Italic, Link, SoftBreak, HardBreak]


def inline_to_plain_text(inline: Inline) -> str:
    """インライン要素からプレーンテキストを抽出する。"""
    if isinstance(inline, Text):
        return inline.value
    if isinstance(inline, Code):
        return f"「{inline.value}」"
    if isinstance(inline, (Bold, Italic)):
        return "".join(inline_to_plain_text(c) for c in inline.children)
    if isinstance(inline, Link):
        return "".join(inline_to_plain_text(c) for c in inline.text)
    return ""


def inlines_to_plain_text(inlines: list[Inline]) -> str:
    return "".join(inline_to_plain_text(c) for c in inlines)


@dataclass
class ListItem:
    content: list[Inline] = field(default_factory=list)
    children: list["Block"] = field(default_factory=list)


@dataclass
class Heading:
    level: int
    content: list[Inline]


@dataclass
class PageBreak:
    pass


@dataclass
class Paragraph:
    content: list[Inline]


@dataclass
class BulletList:
    items: list[ListItem]


@dataclass
class OrderedList:
    items: list[ListItem]
    start: int = 1


@dataclass
class Table:
    headers: list[list[Inline]]
    rows: list[list[list[Inline]]]
    alignments: list[Alignment]


@dataclass
class CodeBlock:
    code: str
    lang: str | None = None


@dataclass
class Image:
    alt: str
    path: str


@dataclass
class BlockQuote:
    children: list["Block"]


@dataclass
class ThematicBreak:
    pass


Block = Union[
    Heading,
    PageBreak,
    Paragraph,
    BulletList,
    OrderedList,
    Table,
    CodeBlock,
    Image,
    BlockQuote,
    ThematicBreak,
]
