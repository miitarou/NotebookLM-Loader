# 📚 NotebookLMに大量のOfficeファイルを一括投入できるツールを作った

![NotebookLM Loader](https://img.shields.io/badge/NotebookLM-Loader-blue?style=for-the-badge&logo=google&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.10+-green?style=for-the-badge&logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

## 🚀 TL;DR

- ❌ NotebookLMに**数百ファイル**を投入したいけど手作業が面倒
- ❌ ファイルサイズ制限（50MB）に引っかかる
- ❌ CSV/ログファイルが読み込みエラーになる

**→ ✅ 全部解決するPythonツールを作りました！**

🔗 https://github.com/miitarou/NotebookLM-Loader

---

## 😫 こんな経験ありませんか？

> 💭 「社内のドキュメントをNotebookLMに入れてRAGで検索したい」
> 
> 💭 「でもファイルが多すぎて手作業は無理...」

私も同じ課題を抱えていました。

- 📁 Word/Excel/PowerPointが**数百ファイル**
- 📊 CSVログが**100MB超**
- 🤯 手動でMarkdown変換→アップロードは現実的じゃない

---

## 🛠️ 作ったもの: NotebookLM Loader

### ✨ 主な機能

| 機能 | 説明 |
|------|------|
| 🧩 **Smart Chunking** | 大量ファイルを約500万文字ごとに自動結合・分割 |
| 📦 **All-in-One** | Office, PDF, CSV, 画像, 圧縮ファイルに対応 |
| 🖼️ **Auto-PDF** | 図が多いPPTは自動でPDFに変換 |
| 🇯🇵 **日本語対応** | Shift-JIS, EUC-JPも自動検出 |
| 🧹 **不可視文字除去** | ゼロ幅スペース等、エラーの原因となる見えない文字を自動除去 |

### 📂 対応ファイル形式

```
📄 Office: .docx, .xlsx, .pptx, .doc, .xls, .ppt
📕 PDF: そのままコピー
📝 テキスト: .txt, .md, .py, .csv, .json, .yaml など
🖼️ 画像: .png, .jpg → PDF変換
📦 圧縮: .zip, .7z, .rar, .lzh
```

---

## 📥 使い方

### インストール

```bash
git clone https://github.com/miitarou/NotebookLM-Loader.git
cd NotebookLM-Loader
pip install -r requirements.txt
```

### 実行

```bash
# 🔹 基本
python office_to_notebooklm.py /path/to/folder

# ⭐ 推奨（スマート結合モード）
python office_to_notebooklm.py /path/to/folder --merge
```

### 出力イメージ

```
📁 target_folder/
├── 📂 converted_files/           # 個別変換ファイル
│   ├── document.md
│   └── spreadsheet.md
└── 📂 converted_files_merged/    # 🎯 NotebookLM用
    ├── Merged_Files_Vol01.md
    ├── Merged_Files_Vol02.md
    └── presentation.pdf
```

**`converted_files_merged/` の中身をNotebookLMにドラッグ&ドロップするだけ！** 🎉

---

## 🔧 技術的なポイント

### 1️⃣ Smart Chunking（行単位分割）

NotebookLMのファイルサイズ制限に対応するため、約500万文字（約5MB）ごとにファイルを結合・分割。

> ⚠️ **発見した制限**: NotebookLMには非公開のファイルサイズ制限があり、約5-10MBを超えるとアップロードエラーになることがあります。本ツールは安全マージンを取って500万文字で分割しています。

**💡 工夫した点**: 単純な文字数分割だと日本語が途中で切れる問題があったので、**行単位**で分割するようにしました。

```python
# 改行位置を探して分割
split_pos = remaining.rfind('\n', 0, available_space)
```

さらに、CSVデータは**カンマ位置**、TSVは**タブ位置**で分割することで、レコードの途中で切れることを防いでいます。

```yaml
# config.yamlで調整可能
processing:
  max_chars_per_volume: 5000000  # 500万文字 ≒ 約5MB
```

### 2️⃣ MarkItDown採用

Microsoft公式の変換エンジン [MarkItDown](https://github.com/microsoft/markitdown) を使用。

- ✅ 表の構造を保持
- ✅ 見出し階層を維持
- ✅ 高い変換精度

### 3️⃣ 視覚密度判定

PowerPointなど画像が多いファイルは、Markdownにすると情報が失われます。

そこで「テキスト量/画像数」の比率を計算し、画像が多いファイルは**自動でPDFに変換**します。

### 4️⃣ 不可視文字の自動除去

ここが実は**最も沼にハマった部分**です。

NotebookLMにファイルをアップロードすると、一部のファイルだけエラーになる現象に遭遇。原因を追跡したところ、**ゼロ幅スペース (U+200B)** という**目に見えない文字**が犯人でした。

```
ASUSTeK COMPUTER INC.​  ← この末尾に見えない文字がある！
```

これはExcelやWebからコピペしたデータに混入することがあります。

**対策として、以下8種類の不可視文字を自動検出・除去しています：**

| 文字 | Unicode | 説明 |
|------|---------|------|
| Zero Width Space | U+200B | ゼロ幅スペース |
| Zero Width Non-Joiner | U+200C | ゼロ幅非結合子 |
| Zero Width Joiner | U+200D | ゼロ幅結合子 |
| Left-to-Right Mark | U+200E | 左右方向制御 |
| Right-to-Left Mark | U+200F | 右左方向制御 |
| Word Joiner | U+2060 | 単語結合子 |
| BOM | U+FEFF | バイトオーダーマーク |
| NULL | U+0000 | ヌル文字 |

**こういう細かいところまでケアしています** 💪

---

## 💼 想定ユースケース

| 👤 ユーザー | 📋 用途 |
|-------------|---------|
| 🏢 **情シス** | 社内マニュアル・規程をNotebookLMで横断検索 |
| 💼 **コンサル** | クライアント資料を一括でRAG化 |
| 👨‍💻 **開発者** | 技術ドキュメントやログをAIに読ませる |

---

## ⚠️ 注意点

- 🔧 **LibreOffice**が必要（自動PDF変換に使用）
- 📏 100MB超のバイナリファイルは処理スキップ
- 🔐 パスワード保護ファイルは検出してレポート

---

## 📝 まとめ

NotebookLMに大量ファイルを投入する作業を自動化するツール「**NotebookLM Loader**」を公開しました。

特に🇯🇵日本語環境での利用を意識して作っています。

ぜひ使ってみてフィードバックいただけると嬉しいです！

🔗 **GitHub**: https://github.com/miitarou/NotebookLM-Loader

⭐ Starいただけると励みになります！

---

## 🏷️ タグ

`#NotebookLM` `#Python` `#RAG` `#生成AI` `#業務効率化` `#MarkItDown`
