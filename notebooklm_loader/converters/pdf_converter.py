# notebooklm_loader/converters/pdf_converter.py
"""PDF変換モジュール"""

import os
import subprocess
from pathlib import Path
from typing import Optional


def convert_to_pdf_via_libreoffice(input_path: Path, output_dir_path: Path) -> Optional[Path]:
    """
    LibreOffice (soffice) を使用してPDF変換を行う
    
    Args:
        input_path: 入力ファイルのパス
        output_dir_path: 出力ディレクトリ
        
    Returns:
        生成されたPDFファイルのパス、失敗時はNone
    """
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
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        # 出力ファイル確認
        original_stem = input_path.stem
        generated_pdf = output_dir_path / (original_stem + ".pdf")
        if generated_pdf.exists():
            return generated_pdf
        return None
    except Exception as e:
        print(f"    [PDF Convert Error] {e}")
        return None
