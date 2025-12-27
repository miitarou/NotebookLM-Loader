# NotebookLM用 Office変換ツール開発タスク

このタスクリストは、NotebookLMへのファイル登録を効率化するための変換ツール開発の進捗を管理します。

## 計画・準備 (Planning)
- [x] **要件定義**: Word/Excel/PPTの変換仕様と視覚情報の扱い(Markdown vs PDF)の決定
- [x] **技術選定**: `python-docx`, `openpyxl`, `python-pptx` (解析用), `markitdown` (変換用) の採用
- [x] **環境構築**: `requirements.txt` の作成と環境セットアップ

## 実装フェーズ (Implementation)
- [x] **Word (.docx) 変換機能**の実装 (MarkItDown Engine)
- [x] **Excel (.xlsx) 変換機能**の実装 (CSV埋め込み or MarkItDown Table)
- [x] **PowerPoint (.pptx) 変換機能**の実装 (MarkItDown Engine)
- [x] **結合機能**の実装 (複数ファイルを1つのテキストにまとめる機能)
- [x] **PPTスキップ機能**の実装 (PDF利用推奨のため)
- [x] **視覚要素検知・レポート機能**の実装 (画像・グラフ数をカウントし、テキストとの比率(密度)でPDF利用を推奨するサマリー表示)

## 検証・仕上げ
- [x] サンプルファイルを用いた変換テスト (Zip/Folder)
- [x] ユーザーへの使用方法ガイド (README) 更新
