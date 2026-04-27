"""parser.py のテスト。Rust 版 src/parser.rs::tests と対応する。"""
from __future__ import annotations

import pytest

from md2docx import ir
from md2docx.parser import parse_markdown


def test_preserves_link_url_in_inline_ir():
    blocks = parse_markdown("[Rust](https://www.rust-lang.org/)")
    assert len(blocks) == 1
    assert isinstance(blocks[0], ir.Paragraph)
    assert isinstance(blocks[0].content[0], ir.Link)
    link = blocks[0].content[0]
    assert link.url == "https://www.rust-lang.org/"
    assert len(link.text) == 1
    assert isinstance(link.text[0], ir.Text)
    assert link.text[0].value == "Rust"


def test_does_not_mix_urls_between_multiple_links():
    blocks = parse_markdown("[A](https://a.example) [B](https://b.example)")
    assert len(blocks) == 1
    para = blocks[0]
    assert isinstance(para, ir.Paragraph)
    links = [c for c in para.content if isinstance(c, ir.Link)]
    assert len(links) == 2
    assert links[0].url == "https://a.example"
    assert links[1].url == "https://b.example"


def test_parses_page_break_directive_as_dedicated_block():
    blocks = parse_markdown("\\pagebreak")
    assert len(blocks) == 1
    assert isinstance(blocks[0], ir.PageBreak)


def test_rejects_page_break_directive_inside_regular_paragraph():
    with pytest.raises(ValueError) as exc_info:
        parse_markdown("before \\pagebreak after")
    assert "段落単位で単独指定した場合のみ利用できます" in str(exc_info.value)


def test_rejects_page_break_in_code_block():
    with pytest.raises(ValueError):
        parse_markdown("```\n\\pagebreak\n```")


def test_table_alignment_extraction():
    md = "| L | C | R |\n| :--- | :---: | ---: |\n| 1 | 2 | 3 |"
    blocks = parse_markdown(md)
    assert len(blocks) == 1
    assert isinstance(blocks[0], ir.Table)
    table = blocks[0]
    assert len(table.headers) == 3
    assert table.alignments[0] == ir.Alignment.LEFT
    assert table.alignments[1] == ir.Alignment.CENTER
    assert table.alignments[2] == ir.Alignment.RIGHT


def test_image_becomes_block_level():
    blocks = parse_markdown("![caption](path/to/img.png)")
    assert any(isinstance(b, ir.Image) for b in blocks)
    img = next(b for b in blocks if isinstance(b, ir.Image))
    assert img.alt == "caption"
    assert img.path == "path/to/img.png"


def test_nested_bullet_list():
    md = "- a\n  - b\n  - c\n- d"
    blocks = parse_markdown(md)
    assert len(blocks) == 1
    bl = blocks[0]
    assert isinstance(bl, ir.BulletList)
    assert len(bl.items) == 2
    inner = bl.items[0].children
    assert len(inner) == 1
    assert isinstance(inner[0], ir.BulletList)
    assert len(inner[0].items) == 2


def test_ordered_list_with_start():
    blocks = parse_markdown("3. third\n4. fourth")
    assert len(blocks) == 1
    ol = blocks[0]
    assert isinstance(ol, ir.OrderedList)
    assert ol.start == 3
    assert len(ol.items) == 2
