# notebooklm_loader/extractors/archive_extractor.py
"""その他の圧縮形式展開モジュール"""

import os
import sys
import tarfile
from pathlib import Path

# オプショナルライブラリ
try:
    import py7zr
    HAS_7Z = True
except ImportError:
    HAS_7Z = False

try:
    import rarfile
    HAS_RAR = True
except ImportError:
    HAS_RAR = False

try:
    import lhafile
    HAS_LZH = True
except ImportError:
    HAS_LZH = False


def extract_7z(archive_path, extract_to) -> str:
    """
    7zファイルを展開
    
    Args:
        archive_path: 圧縮ファイルのパス
        extract_to: 展開先ディレクトリ
        
    Returns:
        処理結果（"OK", "PASSWORD_PROTECTED", "LIBRARY_MISSING", "ERROR"）
    """
    if not HAS_7Z:
        print(f"    [Warning] py7zr not installed, skipping: {archive_path.name}")
        return "LIBRARY_MISSING"
    try:
        with py7zr.SevenZipFile(archive_path, mode='r') as z:
            if z.needs_password():
                return "PASSWORD_PROTECTED"
            z.extractall(path=extract_to)
        return "OK"
    except Exception as e:
        if "password" in str(e).lower():
            return "PASSWORD_PROTECTED"
        print(f"    [7z Extract Error] {e}")
        return "ERROR"


def extract_rar(archive_path, extract_to) -> str:
    """
    RARファイルを展開
    
    Args:
        archive_path: 圧縮ファイルのパス
        extract_to: 展開先ディレクトリ
        
    Returns:
        処理結果（"OK", "PASSWORD_PROTECTED", "LIBRARY_MISSING", "MULTI_VOLUME", "ERROR"）
    """
    if not HAS_RAR:
        print(f"    [Warning] rarfile not installed, skipping: {archive_path.name}")
        return "LIBRARY_MISSING"
    try:
        with rarfile.RarFile(archive_path) as rf:
            if rf.needs_password():
                return "PASSWORD_PROTECTED"
            rf.extractall(path=extract_to)
        return "OK"
    except rarfile.NeedFirstVolume:
        print(f"    [Warning] Multi-volume RAR, skipping: {archive_path.name}")
        return "MULTI_VOLUME"
    except Exception as e:
        if "password" in str(e).lower():
            return "PASSWORD_PROTECTED"
        print(f"    [RAR Extract Error] {e}")
        return "ERROR"


def extract_tar(archive_path, extract_to) -> str:
    """
    tar/tar.gz/tgzファイルを展開
    
    Args:
        archive_path: 圧縮ファイルのパス
        extract_to: 展開先ディレクトリ
        
    Returns:
        処理結果（"OK", "ERROR"）
    """
    try:
        with tarfile.open(archive_path, 'r:*') as tf:
            # ディレクトリトラバーサル対策
            for member in tf.getmembers():
                member_path = os.path.join(extract_to, member.name)
                if not os.path.abspath(member_path).startswith(os.path.abspath(extract_to)):
                    continue
                # Python 3.12+ではfilter引数が必要
                if sys.version_info >= (3, 12):
                    tf.extract(member, extract_to, filter='data')
                else:
                    tf.extract(member, extract_to)
        return "OK"
    except Exception as e:
        print(f"    [TAR Extract Error] {e}")
        return "ERROR"


def extract_lzh(archive_path, extract_to) -> str:
    """
    LZHファイルを展開
    
    Args:
        archive_path: 圧縮ファイルのパス
        extract_to: 展開先ディレクトリ
        
    Returns:
        処理結果（"OK", "LIBRARY_MISSING", "ERROR"）
    """
    if not HAS_LZH:
        print(f"    [Warning] lhafile not installed, skipping: {archive_path.name}")
        return "LIBRARY_MISSING"
    try:
        with lhafile.LhaFile(str(archive_path)) as lf:
            for info in lf.infolist():
                target_path = Path(extract_to) / info.filename
                # ディレクトリトラバーサル対策
                if not os.path.abspath(target_path).startswith(os.path.abspath(extract_to)):
                    continue
                target_path.parent.mkdir(parents=True, exist_ok=True)
                with open(target_path, 'wb') as f:
                    f.write(lf.read(info.filename))
        return "OK"
    except Exception as e:
        print(f"    [LZH Extract Error] {e}")
        return "ERROR"
