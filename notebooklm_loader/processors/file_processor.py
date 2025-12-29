# notebooklm_loader/processors/file_processor.py
"""ファイル判定モジュール"""

import chardet
from typing import Tuple, Optional

try:
    import magic
    HAS_MAGIC = True
    # MIMEタイプ判定用インスタンス（再利用）
    _mime_detector = magic.Magic(mime=True)
except ImportError:
    HAS_MAGIC = False
    _mime_detector = None


def is_text_file(file_path) -> Tuple[bool, Optional[str]]:
    """
    chardetを使ってテキストファイルかどうか判定する
    
    Args:
        file_path: 対象ファイルのパス
        
    Returns:
        (is_text, encoding): テキストファイルかどうかとエンコーディングのタプル
    """
    try:
        with open(file_path, 'rb') as f:
            raw = f.read(8000)  # 先頭8KB程度読んで判定
        
        if not raw:
            return True, 'utf-8'  # 空ファイルはテキスト扱い
        
        result = chardet.detect(raw)
        encoding = result.get('encoding')
        confidence = result.get('confidence', 0)
        
        # 信頼度が低い場合や検出できない場合はバイナリ扱い
        if not encoding or confidence < 0.5:
            return False, None
        
        # 実際に読めるか確認
        try:
            raw.decode(encoding)
            return True, encoding
        except (UnicodeDecodeError, LookupError):
            return False, None
            
    except Exception:
        return False, None


def get_mime_type(file_path) -> Optional[str]:
    """
    ファイルのMIMEタイプを取得する
    
    Args:
        file_path: 対象ファイルのパス
        
    Returns:
        MIMEタイプ文字列、またはNone
    """
    if not _mime_detector:
        return None
    try:
        return _mime_detector.from_file(str(file_path))
    except Exception:
        return None


def is_likely_text_by_mime(file_path) -> Optional[bool]:
    """
    MIMEタイプからテキストファイルかどうか判定
    
    Args:
        file_path: 対象ファイルのパス
        
    Returns:
        True/False/None（判定不可）
    """
    mime = get_mime_type(file_path)
    if mime is None:
        return None
    
    text_mimes = [
        'text/', 'application/json', 'application/xml',
        'application/javascript', 'application/x-sh',
    ]
    return any(mime.startswith(t) for t in text_mimes)
