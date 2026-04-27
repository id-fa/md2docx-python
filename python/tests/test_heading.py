"""heading.py のテスト。"""
from __future__ import annotations

from md2docx.heading import HeadingManager
from md2docx.ir import Text


def _content(text: str):
    return [Text(value=text)]


def test_simple_h1_h2_h3_numbering():
    mgr = HeadingManager()
    assert mgr.next_heading(1, _content("intro")) == "1"
    assert mgr.next_heading(2, _content("sub")) == "1.1"
    assert mgr.next_heading(3, _content("subsub")) == "1.1.1"
    assert mgr.next_heading(2, _content("next")) == "1.2"
    assert mgr.next_heading(1, _content("two")) == "2"
    assert mgr.next_heading(2, _content("subof2")) == "2.1"


def test_h4_uses_parens():
    mgr = HeadingManager()
    mgr.next_heading(1, _content("one"))
    assert mgr.next_heading(4, _content("first")) == "(1)"
    assert mgr.next_heading(4, _content("second")) == "(2)"


def test_h5_uses_circled_numbers():
    mgr = HeadingManager()
    assert mgr.next_heading(5, _content("a")) == "①"
    assert mgr.next_heading(5, _content("b")) == "②"


def test_existing_number_in_h1_is_respected_and_synced():
    mgr = HeadingManager()
    assert mgr.next_heading(1, _content("8 タイトル")) == "8"
    # 次の H2 は 8.1 になる
    assert mgr.next_heading(2, _content("subtitle")) == "8.1"


def test_strip_number_h1():
    mgr = HeadingManager()
    assert mgr.strip_number(1, "8 タイトル") == "タイトル"
    assert mgr.strip_number(1, "no number") == "no number"


def test_strip_number_h4_full_paren():
    mgr = HeadingManager()
    assert mgr.strip_number(4, "(3) hello") == "hello"


def test_strip_number_h5_circled():
    mgr = HeadingManager()
    assert mgr.strip_number(5, "③ foo") == "foo"


def test_h1_resets_lower_levels():
    mgr = HeadingManager()
    mgr.next_heading(1, _content("a"))
    mgr.next_heading(2, _content("b"))
    mgr.next_heading(3, _content("c"))
    mgr.next_heading(4, _content("d"))
    mgr.next_heading(5, _content("e"))
    # 新しい H1 が来たら下位がリセットされる
    assert mgr.next_heading(1, _content("z")) == "2"
    assert mgr.next_heading(2, _content("y")) == "2.1"


def test_current_h1_number():
    mgr = HeadingManager()
    assert mgr.current_h1_number() == 0
    mgr.next_heading(1, _content("a"))
    assert mgr.current_h1_number() == 1
    mgr.next_heading(1, _content("5 jump"))
    assert mgr.current_h1_number() == 5
