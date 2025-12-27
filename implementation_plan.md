# NotebookLM用 変換ツール アップグレード計画 (with MarkItDown)

## 概要
Microsoft公式の `markitdown` ライブラリをエンジンとして採用し、変換精度の向上と対応フォーマットの拡充を行います。一方で、我々が独自に実装した「視覚要素検知（PDF利用推奨レポート）」は、NotebookLM運用の肝となるため、これを維持したハイブリッドな構成にします。

## 変更点

### 1. `requirements.txt` の更新
- 追加: `markitdown`
- 削除: `python-docx`, `python-pptx`, `pandas` (MarkItDownが内部で処理する場合、これらは不要になる可能性がありますが、**視覚要素検知（レポート機能）**のために `python-docx` や `openpyxl` は引き続き解析用として残します。MarkItDownはあくまで「変換」に使います)

### 2. `office_to_notebooklm.py` のリファクタリング

#### 現状のロジック
- ファイルを開く -> 自前の `for` ループでテキスト抽出 -> Markdownリストを手作り

#### 新しいロジック (ハイブリッド)
1.  **解析フェーズ (自作ロジック)**:
    - 従来通り `python-docx/openpyxl/pptx` でファイルを開く。
    - 画像数・文字数をカウントし、`Visual Density` を計算。
    - 「PDF推奨」の判定を行う。
2.  **変換フェーズ (MarkItDown)**:
    - PDF推奨でないファイル（テキスト中心）について、`MarkItDown` APIを呼んで Markdown を生成させる。
    - `md = markitdown.convert(file_path)`

### メリット
- **公式の安心感**: Microsoftのパーサーを使うため、複雑なWordの表組みや、Excelの隠れたデータなどの処理漏れが減ります。
- **機能の分担**: 「変換（Execution）」はAI/ライブラリに任せ、「判断・選別（Direction）」は自作ロジックが担う、という構造になります。

## 手順
1. `requirements.txt` に `markitdown` を追加。
2. スクリプト内の `convert_xxx` 関数の中身を、`MarkItDown` 呼び出しに置き換える（ただし、カウントロジックは残す）。
3. 動作確認。
