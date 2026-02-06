# md2word

Markdown ファイルを Word (.docx) ファイルに変換する CLI ツール。

## 必要環境

- Rust 1.70 以上

## ビルド

```sh
cargo build --release
```

ビルド後のバイナリは `target/release/md2word` に生成されます。

## 使い方

```sh
md2word <入力ファイル> [オプション]
```

### 引数

| 引数             | 説明                             |
| ---------------- | -------------------------------- |
| `<入力ファイル>` | 変換する Markdown ファイルのパス |

### オプション

| オプション            | 説明                                                                |
| --------------------- | ------------------------------------------------------------------- |
| `-o, --output <パス>` | 出力ファイルパス（省略時は入力ファイル名の拡張子を `.docx` に変更） |
| `-c, --config <パス>` | 設定ファイルパス（省略時はデフォルト設定を使用）                    |
| `-h, --help`          | ヘルプを表示                                                        |

### 実行例

```sh
# 基本的な変換（output.docx が生成される）
md2word document.md

# 出力先を指定
md2word document.md -o output.docx

# 設定ファイルを指定
md2word document.md -o output.docx -c md2word.toml
```

#### cargo run で直接実行

```sh
# サンプルファイルを変換
cargo run --release -- examples/test.md

# 出力先を指定して変換
cargo run --release -- examples/test.md -o examples/test.docx

# 設定ファイルも指定して変換
cargo run --release -- examples/test.md -o examples/test.docx -c md2word.toml
```

## 設定ファイル

TOML 形式で、フォントやサイズをカスタマイズできます。全項目省略可能で、省略時はデフォルト値が使われます。

```toml
[fonts]
body_ja = "游明朝"        # 本文の日本語フォント
body_en = "Century"       # 本文の英語フォント
heading_ja = "游ゴシック"  # 見出しの日本語フォント
heading_en = "Century"     # 見出しの英語フォント

[sizes]
body = 10.5      # 本文のフォントサイズ (pt)
heading1 = 14.0  # 見出し1のフォントサイズ (pt)
heading2 = 12.0  # 見出し2のフォントサイズ (pt)
heading3 = 11.0  # 見出し3のフォントサイズ (pt)
heading4 = 11.0  # 見出し4のフォントサイズ (pt)

[indent]
body_left = 210         # 本文の左インデント (twip)
body_first_line = 210   # 本文の字下げ (twip)
body_right = 210        # 本文の右インデント (twip)
body_left_chars = 100   # 本文の左インデント (文字数×100)
heading4_left = 709     # 見出し4の左インデント (twip)
heading4_hanging = 709  # 見出し4のぶら下げインデント (twip)

[bullet]
level0 = "●"   # 箇条書きレベル0の行頭文字
level1 = "■"   # 箇条書きレベル1の行頭文字
level2 = "▲"   # 箇条書きレベル2の行頭文字
```

**インデント単位について:**
- `twip`: Word 内部単位 (1 twip = 1/20 pt)。210 twip ≒ 全角1文字幅。
- `body_left_chars`: Word 独自の文字数単位（100 = 1文字）。

## 対応する Markdown 要素

- 見出し（H1〜H4、自動採番付き）
- 段落
- 箇条書き（ネスト対応）
- 番号付きリスト（ネスト対応）
- 表（自動で表番号を付与）
- コードブロック
- 画像（自動で図番号を付与）
- 水平線
- インライン要素：テキスト、コード、太字、斜体、リンク
