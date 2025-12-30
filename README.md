# NotebookLM Loader (Powered by MarkItDown)

[ [English](README_EN.md) | **日本語** | [中文](README_CN.md) ]


Microsoft Officeファイル（Word, Excel, PowerPoint）をはじめ、多種多様なファイル形式を、NotebookLMでの利用に最適な構造化された形式に一括変換するPythonツールです。
**Microsoft公式の変換エンジン `MarkItDown`** を採用し、高い変換精度を実現しています。

## 主な機能

### 1. Smart Chunking (自動分割結合)
- フォルダ内の大量のファイルを、NotebookLMが読みやすいサイズ（約500万文字 ≒ 約5MB）ごとに自動で結合・分割
- `--merge` オプションで `converted_files_merged` フォルダに最適化されたファイルを生成

### 2. All-in-One Loader
対応するファイル形式を自動検知して取り込み：

| カテゴリ | 対応形式 | 処理方法 |
|----------|----------|----------|
| **Office (新形式)** | .docx, .xlsx, .xls, .pptx | 視覚密度判定 → Markdown or PDF変換 |
| **Office (旧形式)** | .doc, .ppt | LibreOffice経由でPDF変換 |
| **MarkItDown対応** | .rtf, .epub, .msg, .eml | MarkItDownでMarkdown化 |
| **PDF** | .pdf | そのままコピー（変換なし） |
| **Visio** | .vsdx, .vsd | LibreOffice経由でPDF変換 |
| **画像** | .png, .jpg, .gif, .bmp, .tiff, .webp | PillowでPDF変換 |
| **テキスト** | .txt, .md, .py, .csv, .json, .yaml 等 | そのままMarkdownとして取り込み |
| **圧縮ファイル** | .zip, .7z, .rar, .tar.gz, .lzh | 展開して中身を再帰処理 |

> 💡 **拡張子にとらわれない処理**
>
> 本ツールは `file` コマンド（MIMEタイプ判定）を使用し、**拡張子ではなくファイルの中身**で形式を判定します。
> 拡張子がなくても、偽装されていても、正しく処理できます。
> 上記の表にないファイル形式でも、テキスト形式であれば自動的に取り込みます。

### 3. Auto-Switch to PDF (自動PDF化)
- 画像やグラフが多いファイル（High Density）を検知すると、自動的にPDFに変換
- `config.yaml` で閾値調整可能 (`visual_density_threshold`)
- LibreOfficeを使用

### 4. 堅牢なファイル処理
- **MIMEタイプ判定**: 拡張子ではなくファイルの中身で判定（上述）
- **エンコーディング自動検出**: Shift-JIS、EUC-JP等も自動対応
- **巨大ファイル対応**: 100MB超のテキストファイルも行単位で安全に分割して処理
- **不可視文字除去**: ゼロ幅スペース(U+200B)、BOM、方向制御文字など目に見えないがNotebookLMでエラーを引き起こす8種類の不可視文字を自動検出・除去
- **シンボリックリンク無視**: 循環参照を防止
- **パスワード保護検出**: 暗号化ファイルを検出してレポート

### 5. スキップ対象
以下のファイルは処理対象外（自動スキップ）：

| カテゴリ | 拡張子 | 理由 |
|----------|--------|------|
| OneNote | .one, .onetoc2 | 非公開バイナリ形式 |
| Access | .accdb, .mdb | データベース形式 |
| 動画 | .mp4, .avi, .mov, .mkv等 | NotebookLMでは利用不可 |
| 音声 | .mp3, .wav, .aac等 | NotebookLMでは利用不可 |
| CAD | .dwg, .dxf | 専用ビューア必要 |
| 実行ファイル | .exe, .dll, .so | バイナリ |

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
