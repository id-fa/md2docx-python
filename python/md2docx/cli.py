"""Markdown → Word 変換 CLI。

Rust 版 `src/main.rs` の Python 移植。argparse で `pymdd <input> [-o OUT] [-c CONFIG]`
を受け取り、入力 Markdown を読み、出力 .docx を書き出す。実行ファイル名は Rust 版の
`mdd` と区別するために `pymdd` としている。

NOTE: --help 内のデフォルト値は config.py / 上位 README.md と必ず同期させること。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

from .config import Config
from .converter import convert_to_docx
from .parser import parse_markdown


_HELP_EPILOG = """\
使用例:
  pymdd document.md                          入力と同名の .docx を生成
  pymdd document.md -o report.docx           出力先を指定
  pymdd document.md -c pymdd.toml            設定ファイルを指定
  pymdd document.md -o out.docx -c my.toml   両方指定

設定ファイル (TOML):
  省略時はデフォルト値が使われます。全項目省略可能。

  [fonts]
  body_ja    = "游明朝"        # 本文の日本語フォント
  body_en    = "Century"       # 本文の英語フォント
  heading_ja = "游ゴシック"    # 見出しの日本語フォント
  heading_en = "Century"       # 見出しの英語フォント

  [sizes]                       # 単位: pt
  body         = 10.5           # 本文
  table_body   = 9.5            # 表本文
  table_header = 9.5            # 表ヘッダー
  heading1     = 14.0           # 見出し1
  heading2     = 12.0           # 見出し2
  heading3     = 11.0           # 見出し3
  heading4     = 11.0           # 見出し4
  heading5     = 10.5           # 見出し5

  [page]                        # 単位: twip
  width         = 11906         # ページ幅 (既定: A4 縦)
  height        = 16838         # ページ高さ
  margin_top    = 1985          # 上余白
  margin_right  = 1701          # 右余白
  margin_bottom = 1701          # 下余白
  margin_left   = 1701          # 左余白
  margin_header = 851           # ヘッダー余白
  margin_footer = 992           # フッター余白
  margin_gutter = 0             # とじしろ

  [indent]                      # 単位: twip (1 twip = 1/20 pt, 210 twip ≒ 全角1文字)
  body_left        = 210        # 本文の左インデント
  body_first_line  = 210        # 本文の字下げ
  body_right       = 210        # 本文の右インデント
  body_left_chars  = 100        # 本文の左インデント (文字数×100)
  heading1_left    = 420        # 見出し1の左インデント
  heading1_hanging = 420        # 見出し1のぶら下げインデント
  heading2_left    = 612        # 見出し2の左インデント
  heading2_hanging = 612        # 見出し2のぶら下げインデント
  heading3_left    = 783        # 見出し3の左インデント
  heading3_hanging = 783        # 見出し3のぶら下げインデント
  heading4_left    = 709        # 見出し4の左インデント
  heading4_hanging = 709        # 見出し4のぶら下げインデント
  heading5_left    = 709        # 見出し5の左インデント
  heading5_hanging = 709        # 見出し5のぶら下げインデント
  heading6_left    = 709        # 見出し6の左インデント
  heading6_hanging = 709        # 見出し6のぶら下げインデント

  [bullet]
  level0 = "●"                  # 箇条書きレベル0
  level1 = "■"                  # 箇条書きレベル1
  level2 = "▲"                  # 箇条書きレベル2

  [numbering]
  figure_format = "sequential"  # 図番号の形式 (sequential / chapter)
  table_format  = "sequential"  # 表番号の形式 (sequential / chapter)

対応する Markdown 要素:
  見出し (H1-H5, 自動採番)    段落                  箇条書き (ネスト対応)
  番号付きリスト (ネスト対応) 表 (自動表番号付与)   コードブロック
  画像 (自動図番号付与)       改ページ (`\\pagebreak`)   水平線
  インライン: テキスト / コード / 太字 / 斜体 / リンク
"""


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pymdd",
        description="Markdown ファイルを Word (.docx) に変換する CLI ツール",
        epilog=_HELP_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="変換する Markdown ファイルのパス")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="出力ファイルパス [省略時: <入力ファイル名>.docx]",
        metavar="FILE",
    )
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=None,
        help="設定ファイルパス (TOML) [省略時: デフォルト設定]",
        metavar="FILE",
    )
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=_get_version(),
    )
    return parser


def _get_version() -> str:
    try:
        from importlib.metadata import version

        return f"pymdd {version('md2docx')}"
    except Exception:  # noqa: BLE001
        from . import __version__

        return f"pymdd {__version__}"


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    # 設定ファイル
    if args.config is not None:
        try:
            config = Config.load(args.config)
        except OSError as e:
            print(
                f"設定ファイルの読み込みに失敗: {args.config} ({e})",
                file=sys.stderr,
            )
            return 1
        except Exception as e:  # TOML パースエラーなど
            print(
                f"設定ファイルの解釈に失敗: {args.config} ({e})",
                file=sys.stderr,
            )
            return 1
    else:
        config = Config()

    input_path: Path = args.input
    try:
        markdown = input_path.read_text(encoding="utf-8")
    except OSError as e:
        print(f"入力ファイルの読み込みに失敗: {input_path} ({e})", file=sys.stderr)
        return 1

    output_path: Path = args.output or input_path.with_suffix(".docx")
    base_path = input_path.parent if input_path.parent != Path() else Path(".")

    try:
        blocks = parse_markdown(markdown)
    except Exception as e:  # noqa: BLE001
        print(f"Markdownの解釈に失敗: {input_path} ({e})", file=sys.stderr)
        return 1

    try:
        document = convert_to_docx(blocks, config, base_path)
    except Exception as e:  # noqa: BLE001
        print(f"docx 変換に失敗: ({e})", file=sys.stderr)
        return 1

    try:
        document.save(str(output_path))
    except OSError as e:
        print(f"出力ファイルの作成に失敗: {output_path} ({e})", file=sys.stderr)
        return 1

    print(f"変換完了: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
