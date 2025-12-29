# notebooklm_loader/cli.py
"""コマンドラインインターフェースモジュール"""

import argparse


def setup_args():
    """
    コマンドライン引数をパースする
    
    Returns:
        パース済みの引数オブジェクト
    """
    parser = argparse.ArgumentParser(
        description='Office files to Markdown converter for NotebookLM',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s /path/to/folder                    # 基本変換
  %(prog)s /path/to/folder --merge            # スマート結合モード
  %(prog)s /path/to/folder --merge --verbose  # 詳細ログ出力
  %(prog)s /path/to/folder --dry-run          # 実行計画のみ表示
  %(prog)s /path/to/folder --config config.yaml  # 設定ファイル使用
  %(prog)s /path/to/folder --incremental      # 差分処理モード
        """
    )
    # 基本引数
    parser.add_argument('target_dir', help='Target directory containing Office files or ZIP file')
    
    # 出力オプション
    parser.add_argument('--merge', action='store_true', 
                        help='Also create merged output in converted_files_merged directory')
    parser.add_argument('--output-dir', '-o', type=str, default=None,
                        help='Custom output directory (default: converted_files in target)')
    
    # 処理オプション
    parser.add_argument('--skip-ppt', action='store_true', 
                        help='Skip PowerPoint files')
    parser.add_argument('--incremental', action='store_true',
                        help='Process only new/modified files (default behavior)')
    parser.add_argument('--full-rebuild', action='store_true',
                        help='Force reprocess all files, ignore cache')
    
    # ログ・表示オプション
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Enable verbose logging (DEBUG level)')
    parser.add_argument('-q', '--quiet', action='store_true',
                        help='Suppress console output and progress bar')
    parser.add_argument('--dry-run', action='store_true',
                        help='Show what would be processed without actually converting')
    
    # 設定オプション
    parser.add_argument('--config', '-c', type=str, default=None,
                        help='Path to config.yaml file')
    
    return parser.parse_args()

