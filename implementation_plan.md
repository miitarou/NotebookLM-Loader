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

### 4. Smart Chunking (統合・分割機能) (Refined)
- **目的**: NotebookLMの1ファイルあたりの容量制限（目安）や、ファイル数過多を回避する。
- **仕様**:
    - **出力先**: `converted_files_merged/` (オプション `--merge` 指定時のみ生成)
    - **動作**: 変換されたMarkdownテキストを結合していく。
    - **分割ルール (Strict)**: 
        - ヘッダ文字列（`# Filename...`）も含めた厳密な文字数計算を行う。
        - 巨大ファイル分割時も、分割後の各パーツがヘッダ込みで確実に20万文字以下になるよう制御する。
    - **ログ**: テキストとして読めずにスキップしたファイルは `[Skipped Binary] filename` とログ出力する。
    - **廃止**: 旧 `--combine` オプションは廃止（`--merge` に統合）。

### 5. Universal Text Loader (全可読ファイル対応) (New)
- **目的**: 未知の拡張子（.config, .log, ソースコード等）も取りこぼさない。
- **仕様**:
    - **拡張子リスト拡充**: `.ps1`, `.bat`, `.sh`, `.ini`, `.yaml`, `.sql` 等を追加。
    - **バイナリ判定フォールバック**: 拡張子がリストにない場合、`utf-8` で読み込みを試行。読めたらテキストとして扱う（ヌルバイト検知などでバイナリ除外）。

## 手順
1. 再帰処理時の `relative_path` を取得するようにロジック修正。
2. 出力ファイル名生成ロジックを `path_to_filename` のような関数に変更。
3. Markdown書き込み時にヘッダ情報を付与。
