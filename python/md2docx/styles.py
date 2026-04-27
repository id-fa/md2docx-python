"""Word ドキュメントのスタイル / 見出し採番 / 箇条書き採番を python-docx 上に構築。

Rust 版 `src/styles.rs` の Python 移植。`python-docx` のハイレベル API では
abstractNum / num の登録ができないため、`OxmlElement` で styles.xml と
numbering.xml を直接組み立てる。

NOTE: Word のフォントサイズは half-point (1pt = 2 半ポイント) 単位、
段落間隔やインデントは twip (1pt = 20 twip) 単位を使う。
"""
from __future__ import annotations

from typing import Iterable

from docx.document import Document as DocumentObj
from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.parts.numbering import NumberingPart
from lxml import etree  # type: ignore[import-not-found]

from .config import Config


def pt_to_half_point(pt: float) -> int:
    """pt → half-point (Word 内部のフォントサイズ単位)。"""
    return int(pt * 2.0)


def pt_to_twip(pt: float) -> int:
    """pt → twip (1/20 pt)。段落間隔などに使う。"""
    return int(pt * 20.0)


# 本文-見出しスタイルの styleId
BODY_TEXT_STYLE_ID = "13"

# 見出し採番の numId / abstractNumId
HEADING_NUM_ID = 2
HEADING_ABSTRACT_NUM_ID = 8

# 箇条書き採番の numId / abstractNumId / 段落 styleId
BULLET_NUM_ID = 3
BULLET_ABSTRACT_NUM_ID = 9
BULLET_STYLE_ID = "BulletList"

# 見出し1/2 の前後段落間隔 (pt)
HEADING1_BEFORE_PT = 24.0
HEADING1_AFTER_PT = 12.0
HEADING2_BEFORE_PT = 18.0
HEADING2_AFTER_PT = 8.0


def _w(name: str, attrs: dict[str, str] | None = None) -> etree._Element:
    """w:foo な要素を生成するショートカット。"""
    el = OxmlElement(f"w:{name}")
    if attrs:
        for k, v in attrs.items():
            el.set(qn(f"w:{k}"), v)
    return el


def _make_run_fonts(
    ascii_font: str,
    east_asia: str,
    *,
    hi_ansi: str | None = None,
    cs: str | None = None,
) -> etree._Element:
    """w:rFonts 要素を生成する。"""
    rFonts = _w("rFonts", {
        "ascii": ascii_font,
        "hAnsi": hi_ansi or ascii_font,
        "eastAsia": east_asia,
        "cs": cs or ascii_font,
    })
    return rFonts


def _make_run_fonts_east_asia_only(east_asia: str) -> etree._Element:
    return _w("rFonts", {"eastAsia": east_asia})


# --------- numbering part 取り回し ---------

_NUMBERING_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)


def _ensure_numbering_part(document: DocumentObj) -> NumberingPart:
    """numbering.xml を含むパートを取得。なければ新規作成して関連付ける。"""
    main_part = document.part
    for rel in main_part.rels.values():
        if rel.reltype == RELATIONSHIP_TYPE.NUMBERING:
            return rel.target_part
    package = main_part.package
    partname = PackURI("/word/numbering.xml")
    numbering_part = NumberingPart.load(
        partname, CONTENT_TYPE.WML_NUMBERING, _NUMBERING_TEMPLATE.encode("utf-8"), package
    )
    main_part.relate_to(numbering_part, RELATIONSHIP_TYPE.NUMBERING)
    return numbering_part


# --------- style 構築ヘルパ ---------

def _get_or_add_pPr(style_element: etree._Element) -> etree._Element:
    pPr = style_element.find(qn("w:pPr"))
    if pPr is None:
        pPr = _w("pPr")
        # styles.xml では pPr は rPr より前
        style_element.insert(0, pPr)
    return pPr


def _get_or_add_rPr(style_element: etree._Element) -> etree._Element:
    rPr = style_element.find(qn("w:rPr"))
    if rPr is None:
        rPr = _w("rPr")
        style_element.append(rPr)
    return rPr


def _make_style_element(
    style_id: str,
    style_name: str,
    *,
    based_on: str | None = None,
    next_style: str | None = None,
) -> etree._Element:
    """空の <w:style w:type="paragraph"> を作る。"""
    style = _w("style", {"type": "paragraph", "styleId": style_id})
    style.append(_w("name", {"val": style_name}))
    if based_on is not None:
        style.append(_w("basedOn", {"val": based_on}))
    if next_style is not None:
        style.append(_w("next", {"val": next_style}))
    return style


def _set_indent(
    pPr: etree._Element,
    *,
    left: int | None = None,
    right: int | None = None,
    first_line: int | None = None,
    hanging: int | None = None,
    left_chars: int | None = None,
) -> None:
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = _w("ind")
        pPr.append(ind)
    if left is not None:
        ind.set(qn("w:left"), str(left))
    if right is not None:
        ind.set(qn("w:right"), str(right))
    if first_line is not None:
        ind.set(qn("w:firstLine"), str(first_line))
    if hanging is not None:
        ind.set(qn("w:hanging"), str(hanging))
    if left_chars is not None:
        ind.set(qn("w:leftChars"), str(left_chars))


def _set_spacing(pPr: etree._Element, *, before: int, after: int) -> None:
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = _w("spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), str(before))
    spacing.set(qn("w:after"), str(after))


def _set_outline_level(pPr: etree._Element, level: int) -> None:
    outline = pPr.find(qn("w:outlineLvl"))
    if outline is None:
        outline = _w("outlineLvl")
        pPr.append(outline)
    outline.set(qn("w:val"), str(level))


def _set_numbering_property(
    pPr: etree._Element, *, num_id: int | None = None, ilvl: int | None = None
) -> None:
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        numPr = _w("numPr")
        pPr.append(numPr)
    if ilvl is not None:
        existing = numPr.find(qn("w:ilvl"))
        if existing is None:
            existing = _w("ilvl")
            numPr.append(existing)
        existing.set(qn("w:val"), str(ilvl))
    if num_id is not None:
        existing = numPr.find(qn("w:numId"))
        if existing is None:
            existing = _w("numId")
            numPr.append(existing)
        existing.set(qn("w:val"), str(num_id))


def _set_alignment(pPr: etree._Element, value: str) -> None:
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = _w("jc")
        pPr.append(jc)
    jc.set(qn("w:val"), value)


def _set_bold(rPr: etree._Element) -> None:
    if rPr.find(qn("w:b")) is None:
        rPr.append(_w("b"))


def _set_size(rPr: etree._Element, half_pt: int) -> None:
    sz = rPr.find(qn("w:sz"))
    if sz is None:
        sz = _w("sz")
        rPr.append(sz)
    sz.set(qn("w:val"), str(half_pt))


def _set_font(rPr: etree._Element, fonts: etree._Element) -> None:
    existing = rPr.find(qn("w:rFonts"))
    if existing is not None:
        rPr.remove(existing)
    rPr.insert(0, fonts)


# --------- abstractNum / num 構築 ---------

def _build_heading_abstract_num(config: Config) -> etree._Element:
    """見出し採番 (abstractNumId=8) の定義を組み立てる。"""
    abstract = _w("abstractNum", {"abstractNumId": str(HEADING_ABSTRACT_NUM_ID)})
    abstract.append(_w("multiLevelType", {"val": "multilevel"}))

    indent = config.indent

    levels: list[tuple[int, str, str, str | None, int, int]] = [
        # (ilvl, num_fmt, lvl_text, pStyle, left, hanging)
        (0, "decimal", "%1.", "1", indent.heading1_left, indent.heading1_hanging),
        (1, "decimal", "%1.%2.", "2", indent.heading2_left, indent.heading2_hanging),
        (2, "decimal", "%1.%2.%3", "3", indent.heading3_left, indent.heading3_hanging),
        # 全角括弧
        (3, "decimal", "（%4）", "4", indent.heading4_left, indent.heading4_hanging),
        (4, "decimalEnclosedCircle", "%5", "5", indent.heading5_left, indent.heading5_hanging),
        (5, "decimalEnclosedCircle", "%6", None, indent.heading6_left, indent.heading6_hanging),
        (6, "decimal", "%7.", None, 2940, 420),
        (7, "aiueoFullWidth", "(%8)", None, 3360, 420),
        (8, "decimalEnclosedCircle", "%9", None, 3780, 420),
    ]

    for ilvl, num_fmt, lvl_text, p_style, left, hanging in levels:
        lvl = _w("lvl", {"ilvl": str(ilvl)})
        lvl.append(_w("start", {"val": "1"}))
        lvl.append(_w("numFmt", {"val": num_fmt}))
        if p_style is not None:
            lvl.append(_w("pStyle", {"val": p_style}))
        lvl.append(_w("lvlText", {"val": lvl_text}))
        lvl.append(_w("lvlJc", {"val": "left"}))
        pPr = _w("pPr")
        ind = _w("ind", {"left": str(left), "hanging": str(hanging)})
        pPr.append(ind)
        lvl.append(pPr)
        abstract.append(lvl)

    return abstract


def _build_bullet_abstract_num(config: Config) -> etree._Element:
    """箇条書き採番 (abstractNumId=9) の定義を組み立てる。"""
    abstract = _w("abstractNum", {"abstractNumId": str(BULLET_ABSTRACT_NUM_ID)})
    bullet_chars = [
        config.bullet.level0,
        config.bullet.level1,
        config.bullet.level2,
    ]
    for i, ch in enumerate(bullet_chars):
        left = (i + 1) * 360
        hanging = 360
        lvl = _w("lvl", {"ilvl": str(i)})
        lvl.append(_w("start", {"val": "1"}))
        lvl.append(_w("numFmt", {"val": "bullet"}))
        lvl.append(_w("lvlText", {"val": ch}))
        lvl.append(_w("lvlJc", {"val": "left"}))
        pPr = _w("pPr")
        ind = _w("ind", {"left": str(left), "hanging": str(hanging)})
        pPr.append(ind)
        lvl.append(pPr)
        abstract.append(lvl)
    return abstract


def _build_num_binding(num_id: int, abstract_num_id: int) -> etree._Element:
    num = _w("num", {"numId": str(num_id)})
    num.append(_w("abstractNumId", {"val": str(abstract_num_id)}))
    return num


# --------- スタイル登録 ---------

def _register_doc_defaults(document: DocumentObj, config: Config) -> None:
    """docDefaults: 既定フォントとサイズを styles.xml に書き込む。"""
    styles_el = document.styles.element
    docDefaults = styles_el.find(qn("w:docDefaults"))
    if docDefaults is None:
        docDefaults = _w("docDefaults")
        styles_el.insert(0, docDefaults)

    # rPrDefault > rPr > rFonts / sz
    rPrDefault = docDefaults.find(qn("w:rPrDefault"))
    if rPrDefault is None:
        rPrDefault = _w("rPrDefault")
        docDefaults.append(rPrDefault)
    rPr = rPrDefault.find(qn("w:rPr"))
    if rPr is None:
        rPr = _w("rPr")
        rPrDefault.append(rPr)

    fonts = _make_run_fonts(config.fonts.body_en, config.fonts.body_ja)
    _set_font(rPr, fonts)
    _set_size(rPr, pt_to_half_point(config.sizes.body))


def _add_style(document: DocumentObj, style_element: etree._Element) -> None:
    """既存の同名 style があれば置き換え、なければ追加する。"""
    styles_el = document.styles.element
    style_id = style_element.get(qn("w:styleId"))
    if style_id is not None:
        for existing in styles_el.findall(qn("w:style")):
            if existing.get(qn("w:styleId")) == style_id:
                styles_el.remove(existing)
                break
    styles_el.append(style_element)


def _build_normal_style(config: Config) -> etree._Element:
    style = _make_style_element("Normal", "Normal")
    rPr = _get_or_add_rPr(style)
    _set_font(rPr, _make_run_fonts(config.fonts.body_en, config.fonts.body_ja))
    _set_size(rPr, pt_to_half_point(config.sizes.body))
    pPr = _get_or_add_pPr(style)
    _set_alignment(pPr, "both")
    return style


def _build_heading1_style(config: Config) -> etree._Element:
    style = _make_style_element("1", "heading 1", based_on="Normal", next_style="Normal")
    rPr = _get_or_add_rPr(style)
    _set_font(rPr, _make_run_fonts(config.fonts.heading_en, config.fonts.heading_ja))
    _set_size(rPr, pt_to_half_point(config.sizes.heading1))
    _set_bold(rPr)
    pPr = _get_or_add_pPr(style)
    _set_spacing(
        pPr,
        before=pt_to_twip(HEADING1_BEFORE_PT),
        after=pt_to_twip(HEADING1_AFTER_PT),
    )
    _set_outline_level(pPr, 0)
    _set_numbering_property(pPr, num_id=HEADING_NUM_ID)
    return style


def _build_heading2_style(config: Config) -> etree._Element:
    # フォントは heading1 から basedOn で継承するため指定しない
    style = _make_style_element("2", "heading 2", based_on="1", next_style="Normal")
    rPr = _get_or_add_rPr(style)
    _set_size(rPr, pt_to_half_point(config.sizes.heading2))
    pPr = _get_or_add_pPr(style)
    _set_spacing(
        pPr,
        before=pt_to_twip(HEADING2_BEFORE_PT),
        after=pt_to_twip(HEADING2_AFTER_PT),
    )
    _set_outline_level(pPr, 1)
    _set_numbering_property(pPr, ilvl=1)
    return style


def _build_heading3_style(config: Config) -> etree._Element:
    style = _make_style_element("3", "heading 3", based_on="Normal", next_style="Normal")
    rPr = _get_or_add_rPr(style)
    _set_font(rPr, _make_run_fonts(config.fonts.heading_en, config.fonts.heading_ja))
    _set_size(rPr, pt_to_half_point(config.sizes.heading3))
    _set_bold(rPr)
    pPr = _get_or_add_pPr(style)
    _set_outline_level(pPr, 2)
    _set_numbering_property(pPr, num_id=HEADING_NUM_ID, ilvl=2)
    return style


def _build_heading4_style(config: Config) -> etree._Element:
    style = _make_style_element("4", "heading 4", based_on="Normal", next_style="Normal")
    rPr = _get_or_add_rPr(style)
    _set_font(rPr, _make_run_fonts_east_asia_only(config.fonts.heading_ja))
    _set_size(rPr, pt_to_half_point(config.sizes.heading4))
    _set_bold(rPr)
    pPr = _get_or_add_pPr(style)
    _set_indent(
        pPr,
        left=config.indent.heading4_left,
        hanging=config.indent.heading4_hanging,
    )
    _set_outline_level(pPr, 3)
    _set_numbering_property(pPr, num_id=HEADING_NUM_ID, ilvl=3)
    return style


def _build_heading5_style(config: Config) -> etree._Element:
    style = _make_style_element("5", "heading 5", based_on="Normal", next_style="Normal")
    rPr = _get_or_add_rPr(style)
    _set_font(rPr, _make_run_fonts_east_asia_only(config.fonts.heading_ja))
    _set_size(rPr, pt_to_half_point(config.sizes.heading5))
    _set_bold(rPr)
    pPr = _get_or_add_pPr(style)
    _set_outline_level(pPr, 4)
    _set_numbering_property(pPr, num_id=HEADING_NUM_ID, ilvl=4)
    return style


def _build_body_text_style(config: Config) -> etree._Element:
    style = _make_style_element(BODY_TEXT_STYLE_ID, "本文ｰ見出し", based_on="Normal")
    pPr = _get_or_add_pPr(style)
    _set_indent(
        pPr,
        left=config.indent.body_left,
        first_line=config.indent.body_first_line,
        right=config.indent.body_right,
        left_chars=config.indent.body_left_chars,
    )
    return style


def _build_bullet_style() -> etree._Element:
    style = _make_style_element(BULLET_STYLE_ID, "Bullet List", based_on="Normal")
    return style


# --------- メインエントリ ---------

def setup_document_styles(document: DocumentObj, config: Config) -> None:
    """sample.docx 準拠のスタイル・採番定義を Word ドキュメントに適用する。"""
    _register_doc_defaults(document, config)

    for builder in (
        _build_normal_style(config),
        _build_body_text_style(config),
        _build_heading1_style(config),
        _build_heading2_style(config),
        _build_heading3_style(config),
        _build_heading4_style(config),
        _build_heading5_style(config),
        _build_bullet_style(),
    ):
        _add_style(document, builder)

    # numbering.xml を構築
    numbering_part = _ensure_numbering_part(document)
    numbering_el = numbering_part.element
    _append_numberings(
        numbering_el,
        abstract_nums=[
            _build_heading_abstract_num(config),
            _build_bullet_abstract_num(config),
        ],
        nums=[
            _build_num_binding(HEADING_NUM_ID, HEADING_ABSTRACT_NUM_ID),
            _build_num_binding(BULLET_NUM_ID, BULLET_ABSTRACT_NUM_ID),
        ],
    )


def _append_numberings(
    numbering_el: etree._Element,
    *,
    abstract_nums: Iterable[etree._Element],
    nums: Iterable[etree._Element],
) -> None:
    """w:numbering 直下に abstractNum (numId より前) と num を追加する。"""
    for abstract in abstract_nums:
        _replace_or_append(numbering_el, abstract, "abstractNumId")
    for num in nums:
        _replace_or_append(numbering_el, num, "numId")
    # OOXML スキーマ的には abstractNum を先、num を後にする必要があるため整列
    _reorder_numbering(numbering_el)


def _replace_or_append(parent: etree._Element, child: etree._Element, key_attr: str) -> None:
    """同一 key 属性を持つ既存要素があれば置き換え、なければ追加する。"""
    tag = etree.QName(child).localname
    target = child.get(qn(f"w:{key_attr}"))
    for existing in parent.findall(qn(f"w:{tag}")):
        if existing.get(qn(f"w:{key_attr}")) == target:
            parent.replace(existing, child)
            return
    parent.append(child)


def _reorder_numbering(numbering_el: etree._Element) -> None:
    """abstractNum を先、num を後に並び替える。"""
    abstract_nums = list(numbering_el.findall(qn("w:abstractNum")))
    nums = list(numbering_el.findall(qn("w:num")))
    others = [
        ch for ch in list(numbering_el)
        if etree.QName(ch).localname not in ("abstractNum", "num")
    ]
    for ch in list(numbering_el):
        numbering_el.remove(ch)
    for ch in others:
        numbering_el.append(ch)
    for ch in abstract_nums:
        numbering_el.append(ch)
    for ch in nums:
        numbering_el.append(ch)


__all__ = [
    "BODY_TEXT_STYLE_ID",
    "BULLET_NUM_ID",
    "BULLET_STYLE_ID",
    "HEADING_NUM_ID",
    "pt_to_half_point",
    "pt_to_twip",
    "setup_document_styles",
]
