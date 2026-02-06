mod config;
mod converter;
mod heading;
mod ir;
mod parser;
mod styles;

use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use clap::Parser;

use crate::config::Config;

#[derive(Parser)]
#[command(name = "md2word", about = "Markdownファイルを Word (.docx) に変換")]
struct Cli {
    /// 入力マークダウンファイル
    input: PathBuf,

    /// 出力ファイルパス（省略時は入力ファイル名の拡張子を .docx に変更）
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// 設定ファイルパス（省略時はデフォルト設定を使用）
    #[arg(short, long)]
    config: Option<PathBuf>,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // 設定ファイルの読み込み
    let config = match &cli.config {
        Some(path) => Config::load(path)
            .with_context(|| format!("設定ファイルの読み込みに失敗: {}", path.display()))?,
        None => Config::default(),
    };

    // 入力ファイルの読み込み
    let input_path = &cli.input;
    let markdown = std::fs::read_to_string(input_path)
        .with_context(|| format!("入力ファイルの読み込みに失敗: {}", input_path.display()))?;

    // 出力パスの決定
    let output_path = cli.output.unwrap_or_else(|| {
        let mut p = input_path.clone();
        p.set_extension("docx");
        p
    });

    // ベースパス（画像の相対パス解決用）
    let base_path = input_path.parent().unwrap_or_else(|| Path::new("."));

    // Markdown → IR
    let blocks = parser::parse_markdown(&markdown);

    // IR → docx
    let docx = converter::convert_to_docx(&blocks, &config, base_path)?;

    // ファイル書き出し
    let file = std::fs::File::create(&output_path)
        .with_context(|| format!("出力ファイルの作成に失敗: {}", output_path.display()))?;

    docx.build().pack(file)?;

    println!("変換完了: {}", output_path.display());
    Ok(())
}
