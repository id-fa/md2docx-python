"""設定ファイル (TOML) のパースとデフォルト値定義。

Rust 版 `src/config.rs` の Python 移植。フォント / フォントサイズ / ページ余白 /
インデント / 箇条書き行頭文字 / 図表番号フォーマットを保持する。

NOTE: ここでのデフォルト値は cli.py の help テキストおよび上位ディレクトリの
README.md と必ず同期させること (3 箇所)。
"""
from __future__ import annotations

import sys
from dataclasses import dataclass, field, fields, is_dataclass
from pathlib import Path
from typing import Any

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover - older Python fallback
    import tomli as tomllib


@dataclass
class FontConfig:
    body_ja: str = "游明朝"
    body_en: str = "Century"
    heading_ja: str = "游ゴシック"
    heading_en: str = "Century"


@dataclass
class SizeConfig:
    body: float = 10.5
    table_body: float = 9.5
    table_header: float = 9.5
    heading1: float = 14.0
    heading2: float = 12.0
    heading3: float = 11.0
    heading4: float = 11.0
    heading5: float = 10.5


@dataclass
class PageConfig:
    width: int = 11906
    height: int = 16838
    margin_top: int = 1985
    margin_right: int = 1701
    margin_bottom: int = 1701
    margin_left: int = 1701
    margin_header: int = 851
    margin_footer: int = 992
    margin_gutter: int = 0


@dataclass
class IndentConfig:
    body_left: int = 210
    body_first_line: int = 210
    body_right: int = 210
    body_left_chars: int = 100
    heading1_left: int = 420
    heading1_hanging: int = 420
    heading2_left: int = 612
    heading2_hanging: int = 612
    heading3_left: int = 783
    heading3_hanging: int = 783
    heading4_left: int = 709
    heading4_hanging: int = 709
    heading5_left: int = 709
    heading5_hanging: int = 709
    heading6_left: int = 709
    heading6_hanging: int = 709


@dataclass
class BulletConfig:
    level0: str = "●"
    level1: str = "■"
    level2: str = "▲"


@dataclass
class NumberingConfig:
    figure_format: str = "sequential"
    table_format: str = "sequential"


@dataclass
class Config:
    fonts: FontConfig = field(default_factory=FontConfig)
    sizes: SizeConfig = field(default_factory=SizeConfig)
    page: PageConfig = field(default_factory=PageConfig)
    indent: IndentConfig = field(default_factory=IndentConfig)
    bullet: BulletConfig = field(default_factory=BulletConfig)
    numbering: NumberingConfig = field(default_factory=NumberingConfig)

    @classmethod
    def load(cls, path: Path) -> "Config":
        data = tomllib.loads(Path(path).read_text(encoding="utf-8"))
        return cls.from_dict(data)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Config":
        config = cls()
        for f in fields(cls):
            section = data.get(f.name)
            if not isinstance(section, dict):
                continue
            target = getattr(config, f.name)
            _apply_section(target, section)
        return config


def _apply_section(target: Any, section: dict[str, Any]) -> None:
    if not is_dataclass(target):
        return
    valid_keys = {f.name for f in fields(target)}
    for key, value in section.items():
        if key in valid_keys:
            setattr(target, key, value)
