# notebooklm_loader/converters/image_converter.py
"""画像変換モジュール"""

from pathlib import Path
from typing import Optional

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def convert_image_to_pdf(input_path: Path, output_dir_path: Path) -> Optional[Path]:
    """
    画像ファイルをPDFに変換する
    
    Args:
        input_path: 入力画像ファイルのパス
        output_dir_path: 出力ディレクトリ
        
    Returns:
        生成されたPDFファイルのパス、失敗時はNone
    """
    if not HAS_PIL:
        print(f"    [Warning] Pillow not installed, skipping image: {input_path.name}")
        return None
    
    try:
        img = Image.open(input_path)
        # RGBAの場合はRGBに変換（PDF保存のため）
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGB')
        
        output_pdf = output_dir_path / (input_path.stem + ".pdf")
        img.save(output_pdf, "PDF", resolution=100.0)
        return output_pdf
    except Exception as e:
        print(f"    [Image to PDF Error] {e}")
        return None
