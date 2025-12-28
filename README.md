# NotebookLM Loader (Powered by MarkItDown)

[ [English](README_EN.md) | **日本語** | [中文](README_CN.md) ]


Microsoft Officeファイル（Word, Excel, PowerPoint）を、NotebookLMでの利用に最適化されたMarkdown形式に一括変換するPythonツールです。
**Microsoft公式の変換エンジン `MarkItDown`** を採用し、高い変換精度を実現しています。また、独自の「視覚要素検知レポート」により、NotebookLMに登録するべきファイル形式（Markdown vs PDF）の判断を支援します。

## 主な機能

1.  **Smart Chunking (自動分割結合)**:
    *   フォルダ内の大量のファイルを、NotebookLMが読みやすいサイズ（約50万文字）ごとに自動で結合・分割して **`Merged_Files_VolXX.md`** にまとめます。
    *   これらの結合ファイルと、後述する自動変換PDFは **`converted_files_merged` フォルダ** に出力されます。
    *   ユーザーはこのフォルダの中身をNotebookLMにドラッグ＆ドロップするだけで完了します。
2.  **All-in-One Loader**: Officeファイル、PDF、ソースコード、テキストなど、フォルダ内のあらゆる可読データを自動検知して取り込みます。
3.  **Auto-Switch to PDF (自動PDF化)**: 画像やグラフが多いファイル（High Density）を検知すると、**自動的にPDFに変換**して出力します（LibreOfficeを使用）。これにより、NotebookLMへ登録するために手動でPDF化する作業が不要になります。
4.  **高精度Markdown変換**: Microsoft MarkItDownを使用し、リストや表などの構造を正確にテキスト化します。

## 必要要件
## 必要要件
- Python 3.10以上
- **LibreOffice** (自動PDF化機能を利用する場合に必須)
    - **バージョン 7.0 以上** 推奨（Headlessモードが安定しているため）
    - Mac: `/Applications/LibreOffice.app` にインストールされていること
    - Linux/Windows: `soffice` コマンドにパスが通っていること
- 必要なライブラリ: `markitdown`, `python-docx`, `openpyxl`, `python-pptx`, `pandas`
  - `pip install -r requirements.txt` でインストール可能

## インストール
1. リポジトリをダウンロード
2. 依存ライブラリをインストール:
   ```bash
   pip install -r requirements.txt
   ```

## 使い方 (Usage)

変換したいOfficeファイルが入った**フォルダ**、または **ZIPファイル** を指定して実行します。

### フォルダを指定する場合
```bash
python office_to_notebooklm.py /Users/yourname/Documents/MyProject
```

### ZIPファイルを直接指定する場合
ZIPファイルを自動で一時フォルダに展開し、中身を変換・解析します。
```bash
python office_to_notebooklm.py /Users/yourname/Documents/archive.zip
```

### オプション
- `--merge`: `converted_files_merged` フォルダを作成し、スマート結合されたファイルを生成します。
- `--skip-ppt`: PowerPoint (.pptx) の変換をスキップします（PPTはPDF利用を推奨するため）。

```bash
python office_to_notebooklm.py /target/dir --merge --skip-ppt
```

## 注意事項 - 変換の仕様について
- **MarkItDown Engine:** Microsoft公式の強力なパーサーを使用するため、表やリストなどの構造認識精度が高いです。変換されたテキストは、NotebookLMが文脈を理解するのに最適です。
- **視覚要素レポート (Visual Density):**
    - 実行後に表示されるレポートです。各ファイルが「テキスト（Markdown）」として処理されたか、「画像主体（PDF）」として処理されたかを確認できます。
    - 「High Visual Density」と判定されたファイルは、**自動でPDFに変換され出力されています**。ユーザーによる追加の手順は不要です。
