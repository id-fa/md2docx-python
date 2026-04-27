# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

`md2docx` (バイナリ名 `mdd`) は Markdown を日本語向け Word (.docx) に変換する Rust 製 CLI。ディレクトリ名に `python` が含まれるが本体は Rust プロジェクト (リポジトリ直下) である点に注意。pandoc とは違い、日本語フォント既定・見出しと図表の自動採番・TOML 設定一本という方針に特化している。

`python/` 配下に同じ仕様を Python に移植したパッケージ (`md2docx`) があり、`pip install -e python/` で **`pymdd`** コマンドが入る (Rust 版バイナリ `mdd` と区別するため別名)。実装は Rust 版とミラー構造 (`ir.py` / `parser.py` / `heading.py` / `styles.py` / `converter.py`)。依存ライブラリは `markdown-it-py` / `python-docx` / `Pillow` / `tomllib`。

## コマンド

```sh
cargo build --release            # リリースビルド (target/release/mdd)
cargo install --path .           # $HOME/.cargo/bin/mdd にインストール
cargo run -- examples/test.md    # 開発時の動作確認
cargo test                       # 全テスト実行
cargo test --lib parser::tests   # 特定モジュールのテストのみ
cargo test <test_name>           # 名前マッチでテスト1件
cargo clippy --all-targets       # lint
cargo fmt                        # フォーマット
```

`rust-toolchain` は固定していない。`Cargo.toml` の `edition = "2024"` のため Rust 1.85+ 推奨 (README には 1.70 と書かれているが edition2024 機能を使う)。

## アーキテクチャ

Markdown → IR → docx の 3 段パイプライン。

```
main.rs         CLI 引数 + ファイル I/O
  └─ parser.rs    pulldown-cmark Event → ir::Block/Inline
       └─ converter.rs IR → docx-rs Docx
            ├─ styles.rs    文書全体のスタイル/Numbering 定義
            └─ heading.rs   見出し採番カウンタ管理
```

### IR (`src/ir.rs`)

`Block` (Heading / Paragraph / BulletList / OrderedList / Table / CodeBlock / Image / BlockQuote / PageBreak / ThematicBreak) と `Inline` (Text / Code / Bold / Italic / Link / SoftBreak / HardBreak) のシンプルな代数データ型。docx-rs の API に直接依存しない中間表現にすることで、parser と converter を分離している。

### Parser (`src/parser.rs`)

`EventConverter` がステートマシンとして pulldown-cmark のイベントストリームを処理する:
- `inline_stack`: 入れ子インライン要素の構築用スタック
- `list_stack`: ネストリストの構築 (`current_item_inlines` と `current_item_children` で「最初の段落 = リスト項目テキスト、それ以降 = 子ブロック」というルール)
- `table_state`: 表のヘッダー/行/セルを段階的に組み立てる
- `\pagebreak` ディレクティブ: 段落単独の場合のみ `Block::PageBreak` に変換し、それ以外の位置に出現するとエラー (`validate_page_break_usage`)

### Converter (`src/converter.rs`)

`ConvertContext` が変換中の状態 (見出しカウンタ、章番号、図表番号) を保持する。重要な不変条件:
- **章番号付き採番**: H1 が出現した時点で `chapter_number` を更新し、`figure_in_chapter` / `table_in_chapter` をリセット。H2 以下では章は変わらない (= リセットしない)。
- **図表番号の生成方式**: `numbering.figure_format` が `"sequential"` のときは Word の `SEQ` フィールドを埋め込み (Word が再採番する)、`"chapter"` のときは `"X.Y"` をプレーンテキストで書く。
- **画像幅**: `fit_image_to_body_width` でページ幅 - 左右余白に収まるよう EMU 単位で縮小。横長画像でも本文幅を超えない。
- **英日間スペース削除**: `process_text` で ASCII↔日本語境界の半角スペースを 1 文字単位で削除。
- **太字/斜体は無視**: Markdown の `**bold**` `*italic*` は IR には残るが converter では子要素を平坦に展開しているだけで Word の bold/italic 属性は付与しない (見出し・表ヘッダーのスタイル側で bold する設計)。

### Styles (`src/styles.rs`)

docx-rs に文書全体のスタイル定義 (Normal, Heading 1-5, BulletList, 本文ｰ見出し) と Numbering 定義 (見出し採番 `HEADING_NUM_ID=2`, 箇条書き `BULLET_NUM_ID=3`) を一括登録する。

- **採番のレベル割当**: H1=Level 0 (`%1.`), H2=Level 1 (`%1.%2.`), H3=Level 2, H4=Level 3 (全角括弧 `（%4）`), H5=Level 4 (`decimalEnclosedCircle` で丸数字)。Level 5-8 は sample.docx 互換のため定義しているが現状未使用。
- **見出し2 のフォント継承**: `heading2_style` は `based_on("1")` でフォントを継承するため `fonts()` を呼ばない。これを忘れると見出し1のフォント指定が効かなくなる。
- **本文スタイル ID は `"13"`**: sample.docx 由来のマジックナンバー。`BODY_TEXT_STYLE_ID` 定数経由で参照する。
- **`docx-rs` の制約への対処**: テーマファイル (theme1.xml) を生成できないため `RunFonts` に実フォント名を直接書く。`rightChars` `firstLineChars` も出力できないため、絶対値 (twip) と left_chars の組み合わせで近似する。

### Heading 採番 (`src/heading.rs`)

`HeadingManager` は H1-H5 のカウンタを保持し、`next_heading()` でレベルに応じて番号を進める。重要なのは **ユーザーが見出しテキストに既に番号を書いていた場合の同期**:
- `detect_existing_number()` がレベルごとのフォーマット (`8`, `8.1`, `8.1.1`, `(1)`, `①`) を検出する
- 既存番号があれば `sync_counters()` でカウンタ自体を上書きし、以降の自動採番がそこから続くようにする
- `strip_number()` で表示テキストから番号部分を除去する (二重表示防止)

H4 は半角括弧入力 `(1)` でも検出するが出力は全角括弧 `（1）` (numbering.xml の LevelText が `\u{FF08}%4\u{FF09}`)。

## 設定変更時の同期

`config.rs` のデフォルト値・`main.rs` の `after_long_help` 内の TOML サンプル・`README.md` の設定例の **3 箇所** が同じ値になるよう同期する必要がある。`main.rs:14` にも同じ注意書きがある。新しい設定項目を増やすときは:

1. `Config` 構造体にフィールド追加 + `serde(default = "...")` + `default_xxx()` 関数 + `Default` impl
2. `main.rs` の `after_long_help` 文字列にコメント付きで追加
3. `README.md` のサンプル TOML と説明を更新
4. `md2docx.toml` (リポジトリ同梱のデフォルト設定例) も更新

## テスト方針

ユニットテストは各モジュール末尾の `#[cfg(test)] mod tests` に同居している (`parser.rs`, `converter.rs`, `styles.rs`)。converter のテストは `convert_to_docx` を呼び、`docx.document.build()` で生成された XML を文字列マッチで検証する手法を取っている (例: `<w:br w:type="page" />` の存在確認)。

## Python 移植版 (`python/`)

Rust 版とミラー構造で `python/md2docx/` 配下に同名モジュール (`ir.py` / `parser.py` / `heading.py` / `styles.py` / `converter.py` / `cli.py`) を持つ。`pip install -e python/` で **`pymdd`** コマンドが入る (Rust 版バイナリ `mdd` と区別するため別名にしている)。テストは `python/tests/` に pytest 形式で 36 ケース。

依存ライブラリの対応:
- `pulldown-cmark` → `markdown-it-py` (token ストリーム; `nesting` 属性で 1=open / -1=close / 0=self-closing)
- `docx-rs` → `python-docx` + 必要箇所は `OxmlElement` で XML 直書き
- `image` クレート → `Pillow`
- `toml` → 標準 `tomllib` (Python 3.11+) / `tomli` フォールバック

### Rust 版との意図的な動作差

- **太字/斜体の扱い**: Rust 版は `**bold**` `*italic*` をプレーンテキストに平坦化して Word の `w:b` `w:i` を出さない (converter.rs に「計画に従いWordスタイルとしてのbold/italicは使わない」コメントあり)。**Python 版は `w:b` / `w:i` を出力する**。`_add_inline` の `Bold` / `Italic` 分岐で `bold=True` / `italic=True` を子に伝播し、`_add_run` / `_add_hyperlink` が `_set_run_bold` / `_set_run_italic` を呼ぶ。ネスト (`***bothbi***`) は両フラグが同じ run に乗る。
- README は元々「太字対応」と書いていたので、Python 版の方が文書化された仕様と一致している。今後 Rust 側もこれに合わせる場合は `converter.rs` の `Inline::Bold/Italic` 分岐を同様に直す。

### `python-docx` の制約と回避策

- **numbering.xml の生成**: python-docx は abstractNum / num の登録 API を持たないため、`styles.py::_ensure_numbering_part` で `NumberingPart.load` を直接呼んで空の numbering.xml を作り、`OxmlElement` で要素を組み立てて挿入している。
- **見出しスタイル/本文スタイル**: 同 ID で既存スタイルが存在する場合に置き換えるため `_add_style` が事前に削除してから追加する。
- **ハイパーリンク**: 高レベル API がないので `OxmlElement("w:hyperlink")` を構築し、外部 URL は `document.part.relate_to(..., is_external=True)` で rel id を取って `r:id`、内部リンクは `w:anchor` 属性をセットする。
- **SEQ フィールド**: `_add_seq_field` が `(begin, instr, separate, text, end)` の 5 つの run を直接 paragraph._p に append する。

### 設定値の同期点 (Python 版)

Rust 版と同じく config 値は **3 箇所同期** が必要だが、Python 側は別ファイル:
1. `python/md2docx/config.py` の dataclass デフォルト
2. `python/md2docx/cli.py` の `_HELP_EPILOG` (--help エピログ)
3. 親ディレクトリの `README.md` のサンプル TOML

Rust 側の値を変えたら Python 側 (config.py + cli.py) も合わせる。逆も同様。
