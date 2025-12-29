# notebooklm_loader/extractors/zip_extractor.py
"""ZIP展開モジュール"""

import zipfile
import shutil
from pathlib import Path
import os


def extract_zip_with_encoding(zip_path, extract_to) -> str:
    """
    ZIPファイルを解凍する際、Windows等で作成されたShift-JISのファイル名を
    正しく復元して展開する。
    
    Args:
        zip_path: ZIPファイルのパス
        extract_to: 展開先ディレクトリ
        
    Returns:
        処理結果（"OK", "PASSWORD_PROTECTED", "ERROR"）
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            # パスワード保護チェック
            for file_info in z.infolist():
                if file_info.flag_bits & 0x1:  # 暗号化フラグ
                    return "PASSWORD_PROTECTED"
            
            for file_info in z.infolist():
                filename = file_info.filename
                
                # UTF-8フラグが立っていない場合、エンコーディングの補正を試みる
                if file_info.flag_bits & 0x800 == 0:
                    try:
                        # Windows (Japanese) ZIP is often CP932 encoded but marked as CP437
                        filename = filename.encode('cp437').decode('cp932')
                    except (UnicodeDecodeError, UnicodeEncodeError):
                        pass  # エンコーディング変換失敗は元のファイル名を使用
                
                # ターゲットパスの生成
                target_path = Path(extract_to) / filename
                
                # ディレクトリトラバーサル対策
                if not os.path.abspath(target_path).startswith(os.path.abspath(extract_to)):
                    continue
                    
                if file_info.is_dir():
                    target_path.mkdir(parents=True, exist_ok=True)
                else:
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    with z.open(file_info) as source, open(target_path, "wb") as target:
                        shutil.copyfileobj(source, target)
        return "OK"
    except RuntimeError as e:
        if "password" in str(e).lower() or "encrypted" in str(e).lower():
            return "PASSWORD_PROTECTED"
        raise
    except Exception as e:
        print(f"    [ZIP Extract Error] {e}")
        return "ERROR"
