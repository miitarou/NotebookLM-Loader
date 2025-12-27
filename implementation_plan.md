# NotebookLM用 変換ツール アップグレード計画 (with MarkItDown & Zip/Folder Support)

## 概要
Microsoft公式の `markitdown` ライブラリをエンジンとして採用し、変換精度の向上と対応フォーマットの拡充を行います。
さらに、ZIP/フォルダ構造のメタデータを保持し、NotebookLMが文脈（どのプロジェクトの、どの階層の資料か）を理解できるようにします。

## 変更点

### 1. `office_to_notebooklm.py` の機能強化

#### フォルダ構造の保持 (Context Preservation)
- **現状**: 全て `converted_files` 直下にフラットに出力されるため、同名ファイルが上書きされたり、元々どこのフォルダにあったかが分からなくなる。
- **NotebookLMへのヒント**: AIコンテキストにおいて「ファイルパス」は重要なメタ情報です。
- **対策**:
    1. **ファイル名**: `元の親フォルダ名_ファイル名.md` のようにプレフィックスを付けてユニークにする（あるいは相対パスをアンダースコアで繋ぐ）。
    2. **Markdownヘッダ**: ファイルの冒頭に以下のようなメタデータを埋め込む。
       ```markdown
       # File Metadata
       - Original Path: Project/2024/Meeting_Logs/June.docx
       - Source Archive: archive.zip (ZIPの場合)
       ```

### 2. ハイブリッド構成（維持）
- **解析フェーズ (Architect)**: 画像数カウント等は自作ロジック。
- **変換フェーズ (Builder)**: テキスト化はMarkItDown。

### 3. その他ファイルのパススルー & コンテキスト化 (New)
- **目的**: 変換対象外のファイル（PDF, ソースコード, テキストメモ等）も、`converted_files` に集約し、NotebookLMに一括アップロードできるようにする。
- **戦略**:
    - **PDFなどのバイナリ**: 単純コピーだが、ファイル名にパスを含める (`Folder_Sub_file.pdf`) ことでコンテキストを保持。
    - **テキストファイル (.txt, .md, .py, etc)**: 内容を読み込み、Markdownヘッダ（Original Path等）を付与して `.md` として保存する。これによりNotebookLMが文脈を理解しやすくなる。

## 手順
1. 再帰処理時の `relative_path` を取得するようにロジック修正。
2. 出力ファイル名生成ロジックを `path_to_filename` のような関数に変更。
3. Markdown書き込み時にヘッダ情報を付与。
