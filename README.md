# NotebookLM Loader (Powered by MarkItDown)

[ [English](README_EN.md) | **日本語** | [中文](README_CN.md) ]


Microsoft Officeファイル（Word, Excel, PowerPoint）をはじめ、**多種多様なファイル形式**をNotebookLMでの利用に最適化された形式に一括変換するPythonツールです。
**Microsoft公式の変換エンジン `MarkItDown`** を採用し、高い変換精度を実現しています。

## 主な機能

### 1. Smart Chunking (自動分割結合)
- フォルダ内の大量のファイルを、NotebookLMが読みやすいサイズ（約35MB）ごとに自動で結合・分割
- `--merge` オプションで `converted_files_merged` フォルダに最適化されたファイルを生成

### 2. All-in-One Loader
対応するファイル形式を自動検知して取り込み：

| カテゴリ | 対応形式 |
|----------|----------|
| **Office (新形式)** | .docx, .xlsx, .pptx |
| **Office (旧形式)** | .doc, .xls, .ppt |
| **MarkItDown対応** | .rtf, .epub, .msg, .eml |
| **PDF** | .pdf (そのままコピー) |
| **Visio** | .vsdx, .vsd → PDF変換 |
| **画像** | .png, .jpg, .gif, .bmp, .tiff, .webp → PDF変換 |
| **テキスト** | .txt, .md, .py, .js, .html, .css, .json, .yaml, .xml, .csv, .log, .ini, .toml 等 |
| **圧縮ファイル** | .zip, .7z, .rar, .tar.gz, .tgz, .lzh |

### 3. Auto-Switch to PDF (自動PDF化)
- 画像やグラフが多いファイル（High Density）を検知すると、自動的にPDFに変換
- LibreOfficeを使用

### 4. 堅牢なファイル処理
- **MIMEタイプ判定**: 拡張子ではなくファイルの中身で判定
- **エンコーディング自動検出**: Shift-JIS、EUC-JP等も自動対応
- **巨大ファイルスキップ**: 100MB超のファイルは自動スキップ
- **シンボリックリンク無視**: 循環参照を防止
- **パスワード保護検出**: 暗号化ファイルを検出してレポート

### 5. スキップ対象
以下のファイルは処理対象外（自動スキップ）：
- OneNote (.one)
- Access (.accdb, .mdb)
- 動画/音声 (.mp4, .mp3, .wav 等)
- CAD (.dwg, .dxf)
- 実行ファイル (.exe, .dll)

## 必要要件

- Python 3.10以上
- **LibreOffice** (自動PDF化機能を利用する場合に必須)
    - バージョン 7.0 以上推奨
    - Mac: `/Applications/LibreOffice.app` にインストール
    - Linux/Windows: `soffice` コマンドにパスが通っていること

## インストール

```bash
git clone https://github.com/miitarou/NotebookLM-Loader.git
cd NotebookLM-Loader
pip install -r requirements.txt
```

### 追加依存（オプション機能用）

| ライブラリ | 用途 |
|------------|------|
| `py7zr` | 7z形式の展開 |
| `rarfile` | RAR形式の展開（unrarコマンドも必要） |
| `lhafile` | LZH形式の展開 |
| `python-magic-bin` | MIMEタイプ判定 |
| `Pillow` | 画像→PDF変換 |

## 使い方

### 基本
```bash
python office_to_notebooklm.py /path/to/folder
```

### ZIPファイルを直接指定
```bash
python office_to_notebooklm.py /path/to/archive.zip
```

### 推奨オプション（スマート結合モード）
```bash
python office_to_notebooklm.py /path/to/folder --merge
```

### オプション一覧
| オプション | 説明 |
|------------|------|
| `--merge` | スマート結合モード（推奨） |
| `--skip-ppt` | PowerPointをスキップ |

## 出力

### ディレクトリ構造
```
target_folder/
├── converted_files/           # 個別変換ファイル
│   ├── document.md
│   ├── spreadsheet.md
│   └── presentation.pdf
└── converted_files_merged/    # --merge時のみ生成
    ├── Merged_Files_Vol01.md  # 結合ファイル
    ├── Merged_Files_Vol02.md
    └── image.pdf              # PDF類
```

### 実行後レポート
- **Visual Density Report**: 各ファイルの処理結果（Markdown/PDF）
- **Password Protected Files**: パスワード保護で処理できなかったファイル一覧

## ライセンス

MIT License
