# notebooklm_loader/utils.py
"""ユーティリティモジュール"""

import os
import re
from pathlib import Path


def sanitize_filename(name: str) -> str:
    """
    ファイル名として使える文字だけに置換
    
    Args:
        name: 元のファイル名
        
    Returns:
        サニタイズされたファイル名
    """
    return re.sub(r'[\\/*?:"<>|]', "", name)


def get_output_filename(root_path: Path, file_path: Path, extension: str = ".md") -> str:
    """
    元のフォルダ構造を反映したファイル名を生成する
    
    Args:
        root_path: ルートパス
        file_path: ファイルパス
        extension: 出力拡張子
        
    Returns:
        フラット化されたファイル名（例: A_B_file.md）
    """
    try:
        rel_path = file_path.relative_to(root_path)
        flat_name = str(rel_path.with_suffix('')).replace(os.sep, '_')
        return sanitize_filename(flat_name) + extension
    except ValueError:
        return file_path.stem + extension
