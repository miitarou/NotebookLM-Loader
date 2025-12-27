# Office to NotebookLM Converter (Powered by MarkItDown)

Microsoft Officeファイル（Word, Excel, PowerPoint）を、NotebookLMでの利用に最適化されたMarkdown形式に一括変換するPythonツールです。
**Microsoft公式の変換エンジン `MarkItDown`** を採用し、高い変換精度を実現しています。また、独自の「視覚要素検知レポート」により、NotebookLMに登録するべきファイル形式（Markdown vs PDF）の判断を支援します。

## 主な機能
1.  **高精度Markdown変換**: Microsoft MarkItDownを使用し、Officeファイルの構造を正確にテキスト化。
2.  **視覚要素レポート**: 画像やグラフの多さを自動検知し、「このファイルはPDFでアップロードすべき」とアドバイス。
3.  **PDF/Markdownの使い分け支援**: テキスト中心なら本ツール、図解中心ならPDF、という最適な運用を提案。

## 必要要件
- Python 3.10以上
- 必要なライブラリ: `markitdown`, `python-docx`, `openpyxl`, `python-pptx`, `pandas`
  - `pip install -r requirements.txt` でインストール可能

## インストール
1. リポジトリをダウンロード
2. 依存ライブラリをインストール:
   ```bash
   pip install -r requirements.txt
   ```

## 使い方 (Usage)

変換したいOfficeファイルが入ったフォルダを指定して実行します。

```bash
python office_to_notebooklm.py /Users/yourname/Documents/MyProject
```

### オプション
- `--combine`: 全変換ファイルを `All_Files_Combined.txt` という1つのファイルに結合します。
- `--skip-ppt`: PowerPoint (.pptx) の変換をスキップします（PPTはPDF利用を推奨するため）。

```bash
python office_to_notebooklm.py /target/dir --combine --skip-ppt
```

## 注意事項 - 変換の仕様について
- **MarkItDown Engine:** Microsoft公式の強力なパーサーを使用するため、表やリストなどの構造認識精度が高いです。変換されたテキストは、NotebookLMが文脈を理解するのに最適です。
- **視覚要素レポート (Visual Density):** 
    - 実行後に表示されるレポートで「High Visual Density」と警告されたファイルは、Markdown変換では情報が欠落する可能性が高いです。
    - **推奨:** これらのファイルはMarkdownを使わず、**PDF形式**でNotebookLMに直接アップロードしてください。
