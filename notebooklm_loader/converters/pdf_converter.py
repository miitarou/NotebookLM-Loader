# notebooklm_loader/converters/pdf_converter.py
"""PDF変換モジュール"""

import os
import subprocess
import time
import logging
from pathlib import Path
from typing import Optional


def convert_to_pdf_via_libreoffice(input_path: Path, output_dir_path: Path, max_retries: int = 3) -> Optional[Path]:
    """
    LibreOffice (soffice) を使用してPDF変換を行う
    
    Args:
        input_path: 入力ファイルのパス
        output_dir_path: 出力ディレクトリ
        max_retries: 最大リトライ回数（デフォルト: 3）
        
    Returns:
        生成されたPDFファイルのパス、失敗時はNone
    """
    logger = logging.getLogger("notebooklm_loader")
    
    # sofficeパスを検索
    soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not os.path.exists(soffice_path):
        soffice_path = "soffice"  # Try PATH

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir_path),
        str(input_path)
    ]
    
    last_error = None
    for attempt in range(max_retries):
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # 出力ファイル確認
            original_stem = input_path.stem
            generated_pdf = output_dir_path / (original_stem + ".pdf")
            if generated_pdf.exists():
                return generated_pdf
            
            # ファイルが見つからない場合もリトライ
            raise FileNotFoundError(f"Generated PDF not found: {generated_pdf}")
            
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt  # 指数バックオフ: 1, 2, 4秒
                logger.debug(f"Retry {attempt + 1}/{max_retries} PDF conversion for {input_path.name} after {wait_time}s: {e}")
                time.sleep(wait_time)
            else:
                logger.warning(f"    [PDF Convert Error] {input_path.name} after {max_retries} attempts: {e}")
    
    return None

