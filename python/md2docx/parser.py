"""Markdown → IR パーサ。

Rust 版 `src/parser.rs` の Python 移植。`pulldown-cmark` の代わりに `markdown-it-py`
の token ストリームを利用する。token は `Token(type, tag, nesting, ...)` 形式の
フラットなリストで、`nesting` が 1=open / -1=close / 0=self-closing を表す。

`\\pagebreak` ディレクティブは段落単独で書かれた場合のみ Block.PageBreak に変換し、
それ以外の位置に出現するとエラー。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

from markdown_it import MarkdownIt
from markdown_it.token import Token

from .ir import (
    Alignment,
    Block,
    Bold,
    BulletList,
    Code,
    CodeBlock,
    HardBreak,
    Heading,
    Image,
    Inline,
    Italic,
    Link,
    ListItem,
    OrderedList,
    PageBreak,
    Paragraph,
    SoftBreak,
    Table,
    Text,
    ThematicBreak,
    inlines_to_plain_text,
)


PAGE_BREAK_DIRECTIVE = r"\pagebreak"


@dataclass
class _ListContext:
    ordered: bool
    start: int = 1
    items: list[ListItem] = field(default_factory=list)
    current_item_inlines: list[Inline] = field(default_factory=list)
    current_item_children: list[Block] = field(default_factory=list)


@dataclass
class _TableState:
    headers: list[list[Inline]] = field(default_factory=list)
    rows: list[list[list[Inline]]] = field(default_factory=list)
    current_row: list[list[Inline]] = field(default_factory=list)
    current_cell: list[Inline] = field(default_factory=list)
    in_header: bool = False
    alignments: list[Alignment] = field(default_factory=list)


def _attr_str(token: Token, name: str) -> str:
    """`Token.attrGet` は str/int/float/None を返すため、必ず文字列として取り出す。"""
    value = token.attrGet(name)
    if value is None:
        return ""
    return str(value)


class _Parser:
    def __init__(self) -> None:
        self.blocks: list[Block] = []
        # ネスト可能なインライン (em / strong / link / paragraph 内側) のスタック
        self.inline_stack: list[list[Inline]] = []
        self.list_stack: list[_ListContext] = []
        self.table_state: _TableState | None = None
        self.in_block_quote = False
        self.block_quote_blocks: list[Block] = []
        self.current_link_url: str | None = None
        self._pending_heading_level: int | None = None

    def parse(self, source: str) -> list[Block]:
        md = MarkdownIt("commonmark").enable(["table", "strikethrough"])
        tokens = md.parse(source)
        for token in tokens:
            self._process_block_token(token)
        _validate_page_break_usage(self.blocks)
        return self.blocks

    # ---- block-level token dispatch ----

    def _process_block_token(self, token: Token) -> None:
        t = token.type
        if t == "heading_open":
            self.inline_stack.append([])
            self._pending_heading_level = int(token.tag[1:])
        elif t == "heading_close":
            content = self.inline_stack.pop() if self.inline_stack else []
            level = self._pending_heading_level or 1
            self._pending_heading_level = None
            self._add_block(Heading(level=level, content=content))
        elif t == "paragraph_open":
            self.inline_stack.append([])
        elif t == "paragraph_close":
            content = self.inline_stack.pop() if self.inline_stack else []
            if _is_page_break_paragraph(content):
                self._add_block(PageBreak())
            elif content:
                self._add_block(Paragraph(content=content))
        elif t == "inline":
            self._process_inline_children(token.children or [])
        elif t == "bullet_list_open":
            self.list_stack.append(_ListContext(ordered=False))
        elif t == "bullet_list_close":
            ctx = self.list_stack.pop()
            self._add_block(BulletList(items=ctx.items))
        elif t == "ordered_list_open":
            start_attr = token.attrGet("start")
            try:
                start = int(start_attr) if start_attr is not None else 1
            except (TypeError, ValueError):
                start = 1
            self.list_stack.append(_ListContext(ordered=True, start=start))
        elif t == "ordered_list_close":
            ctx = self.list_stack.pop()
            self._add_block(OrderedList(items=ctx.items, start=ctx.start))
        elif t == "list_item_open":
            ctx = self.list_stack[-1]
            ctx.current_item_inlines = []
            ctx.current_item_children = []
        elif t == "list_item_close":
            ctx = self.list_stack[-1]
            ctx.items.append(
                ListItem(
                    content=ctx.current_item_inlines,
                    children=ctx.current_item_children,
                )
            )
            ctx.current_item_inlines = []
            ctx.current_item_children = []
        elif t == "table_open":
            self.table_state = _TableState()
        elif t == "table_close":
            if self.table_state is not None:
                state = self.table_state
                self.table_state = None
                self._add_block(
                    Table(
                        headers=state.headers,
                        rows=state.rows,
                        alignments=state.alignments,
                    )
                )
        elif t == "thead_open":
            if self.table_state is not None:
                self.table_state.in_header = True
        elif t == "thead_close":
            if self.table_state is not None:
                self.table_state.in_header = False
        elif t in ("tbody_open", "tbody_close"):
            pass
        elif t == "tr_open":
            if self.table_state is not None:
                self.table_state.current_row = []
        elif t == "tr_close":
            if self.table_state is not None:
                state = self.table_state
                if state.in_header:
                    state.headers = state.current_row
                else:
                    state.rows.append(state.current_row)
                state.current_row = []
        elif t in ("th_open", "td_open"):
            if self.table_state is not None:
                self.table_state.current_cell = []
                if self.table_state.in_header:
                    self.table_state.alignments.append(_extract_alignment(token))
        elif t in ("th_close", "td_close"):
            if self.table_state is not None:
                self.table_state.current_row.append(self.table_state.current_cell)
                self.table_state.current_cell = []
        elif t == "code_block":
            self._add_block(CodeBlock(code=token.content, lang=None))
        elif t == "fence":
            lang = token.info.strip() if token.info else None
            self._add_block(CodeBlock(code=token.content, lang=lang or None))
        elif t == "hr":
            self._add_block(ThematicBreak())
        elif t == "blockquote_open":
            self.in_block_quote = True
            self.block_quote_blocks = []
        elif t == "blockquote_close":
            children = self.block_quote_blocks
            self.block_quote_blocks = []
            self.in_block_quote = False
            from .ir import BlockQuote

            self._add_block(BlockQuote(children=children))

    # ---- inline token dispatch ----

    def _process_inline_children(self, children: Iterable[Token]) -> None:
        for child in children:
            t = child.type
            if t == "text":
                self._push_inline(Text(value=child.content))
            elif t == "code_inline":
                self._push_inline(Code(value=child.content))
            elif t == "softbreak":
                self._push_inline(SoftBreak())
            elif t == "hardbreak":
                self._push_inline(HardBreak())
            elif t == "em_open":
                self.inline_stack.append([])
            elif t == "em_close":
                inner = self.inline_stack.pop() if self.inline_stack else []
                self._push_inline(Italic(children=inner))
            elif t == "strong_open":
                self.inline_stack.append([])
            elif t == "strong_close":
                inner = self.inline_stack.pop() if self.inline_stack else []
                self._push_inline(Bold(children=inner))
            elif t == "link_open":
                self.inline_stack.append([])
                self.current_link_url = _attr_str(child, "href")
            elif t == "link_close":
                text = self.inline_stack.pop() if self.inline_stack else []
                url = self.current_link_url or ""
                self.current_link_url = None
                if url:
                    self._push_inline(Link(text=text, url=url))
                else:
                    for inl in text:
                        self._push_inline(inl)
            elif t == "image":
                # 画像は段落内に出現してもブロックレベルとして取り出す (Rust 版踏襲)
                alt_inlines = list(_inline_children_to_inlines(child.children or []))
                if alt_inlines:
                    alt = inlines_to_plain_text(alt_inlines)
                else:
                    alt = child.content or ""
                src = _attr_str(child, "src")
                self._add_block(Image(alt=alt, path=src))
            # strikethrough (s_open/s_close) や html_inline などは無視

    # ---- helpers ----

    def _push_inline(self, inline: Inline) -> None:
        if self.inline_stack:
            self.inline_stack[-1].append(inline)
        elif self.table_state is not None:
            self.table_state.current_cell.append(inline)
        elif self.list_stack:
            self.list_stack[-1].current_item_inlines.append(inline)
        # それ以外 (タイトリスト外) では捨てる: 文書直下のインラインは
        # markdown-it-py 上は paragraph_open に必ず包まれるため通常到達しない

    def _add_block(self, block: Block) -> None:
        if self.in_block_quote:
            self.block_quote_blocks.append(block)
            return
        if self.list_stack:
            ctx = self.list_stack[-1]
            if isinstance(block, Paragraph):
                # リストアイテム内の最初の段落 → リスト項目テキスト
                # 2 つ目以降の段落 → 子ブロック
                if not ctx.current_item_inlines:
                    ctx.current_item_inlines = list(block.content)
                else:
                    ctx.current_item_children.append(block)
            else:
                ctx.current_item_children.append(block)
            return
        self.blocks.append(block)


def _inline_children_to_inlines(tokens: Iterable[Token]) -> list[Inline]:
    """画像 alt 用に、簡易的に inline トークンを Inline へ変換する。"""
    result: list[Inline] = []
    for tk in tokens:
        if tk.type == "text":
            result.append(Text(value=tk.content))
        elif tk.type == "code_inline":
            result.append(Code(value=tk.content))
    return result


def _extract_alignment(token: Token) -> Alignment:
    """th_open / td_open の style 属性から alignment を取り出す。"""
    style = _attr_str(token, "style")
    compact = style.replace(" ", "")
    if "text-align:right" in compact:
        return Alignment.RIGHT
    if "text-align:center" in compact:
        return Alignment.CENTER
    if "text-align:left" in compact:
        return Alignment.LEFT
    return Alignment.NONE


def _is_page_break_paragraph(content: list[Inline]) -> bool:
    if len(content) != 1:
        return False
    only = content[0]
    if not isinstance(only, Text):
        return False
    return only.value.strip() == PAGE_BREAK_DIRECTIVE


# ---- pagebreak 検証 ----

class PageBreakError(ValueError):
    pass


def _validate_page_break_usage(blocks: Iterable[Block]) -> None:
    for block in blocks:
        _validate_block(block)


def _validate_block(block: Block) -> None:
    from .ir import BlockQuote

    if isinstance(block, (Heading, Paragraph)):
        _validate_inlines(block.content)
    elif isinstance(block, (BulletList, OrderedList)):
        for item in block.items:
            _validate_inlines(item.content)
            for child in item.children:
                _validate_block(child)
    elif isinstance(block, Table):
        for cell in block.headers:
            _validate_inlines(cell)
        for row in block.rows:
            for cell in row:
                _validate_inlines(cell)
    elif isinstance(block, CodeBlock):
        _ensure_no_directive(block.code)
    elif isinstance(block, Image):
        _ensure_no_directive(block.alt)
    elif isinstance(block, BlockQuote):
        for child in block.children:
            _validate_block(child)
    # PageBreak / ThematicBreak はチェック対象外


def _validate_inlines(inlines: list[Inline]) -> None:
    for inline in inlines:
        _validate_inline(inline)


def _validate_inline(inline: Inline) -> None:
    if isinstance(inline, (Text, Code)):
        _ensure_no_directive(inline.value)
    elif isinstance(inline, (Bold, Italic)):
        _validate_inlines(inline.children)
    elif isinstance(inline, Link):
        _validate_inlines(inline.text)
        _ensure_no_directive(inline.url)


def _ensure_no_directive(text: str) -> None:
    if PAGE_BREAK_DIRECTIVE in text:
        raise PageBreakError(
            f"`{PAGE_BREAK_DIRECTIVE}`は段落単位で単独指定した場合のみ利用できます"
        )


# ---- public API ----

def parse_markdown(source: str) -> list[Block]:
    """Markdown 文字列を IR ブロックリストに変換する。"""
    return _Parser().parse(source)
