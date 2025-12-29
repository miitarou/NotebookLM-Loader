# notebooklm_loader/logger.py
"""ログ機構モジュール"""

import logging
import datetime
from pathlib import Path


def setup_logging(output_dir: Path, verbose: bool = False) -> logging.Logger:
    """
    ログ機構をセットアップする
    
    Args:
        output_dir: 出力ディレクトリ（ログファイル出力先）
        verbose: 詳細ログ出力フラグ
        
    Returns:
        設定済みのロガー
    """
    log_level = logging.DEBUG if verbose else logging.INFO
    
    # ログディレクトリ作成
    log_dir = output_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # ログファイル名（タイムスタンプ付き）
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"processing_{timestamp}.log"
    
    # ロガー設定
    logger = logging.getLogger("notebooklm_loader")
    logger.setLevel(log_level)
    
    # 既存ハンドラをクリア
    logger.handlers.clear()
    
    # ファイルハンドラ
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_format = logging.Formatter('%(asctime)s | %(levelname)-8s | %(message)s')
    file_handler.setFormatter(file_format)
    logger.addHandler(file_handler)
    
    # コンソールハンドラ
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_format = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)
    
    logger.info(f"Log file: {log_file}")
    return logger


def get_logger() -> logging.Logger:
    """アプリケーションロガーを取得"""
    return logging.getLogger("notebooklm_loader")
