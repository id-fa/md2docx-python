# pymdd (Python 移植版)

Markdown ファイルを Word (.docx) に変換する CLI ツール。Rust 版 (`../`) の Python 移植。
Rust 版バイナリ `mdd` と同居しても紛らわしくないよう、コマンド名は `pymdd` にしている。

## インストール

Python 3.10 以上が必要。

```sh
pip install -e .
# あるいはビルド済みのものを
pip install .
```

`pymdd` コマンドが PATH に登録される。

開発用には:

```sh
pip install -e .[dev]   # 編集モード
pytest                  # テスト実行
```

## 使い方

```sh
pymdd document.md                          # document.docx を生成
pymdd document.md -o report.docx           # 出力先を指定
pymdd document.md -c pymdd.toml            # 設定ファイルを指定
pymdd document.md -o out.docx -c my.toml   # 両方指定
pymdd --help                               # 詳細ヘルプ
```

`python -m md2docx <file>` でも実行可能。

## 設定ファイル / 対応 Markdown 要素

Rust 版と同じ仕様。詳細は親ディレクトリの [`README.md`](../README.md) を参照。

## Rust 版との違い

- 採番、フォント、インデント、表番号・図番号など振る舞いはすべて一致するように移植している。
- コマンド名のみ `mdd` (Rust) / `pymdd` (Python) で区別する。
- 内部表現 (IR) と変換ロジックの構造は同一だが、依存ライブラリが異なる:
    - `pulldown-cmark` → `markdown-it-py`
    - `docx-rs` → `python-docx` (+ 必要箇所は `OxmlElement` で XML 直書き)
    - `image` クレート → `Pillow`
- `python-docx` の制約上、numbering.xml は `lxml` で直接構築している。
