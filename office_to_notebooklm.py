import os
import argparse
from pathlib import Path
import docx
import openpyxl
import pandas as pd
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import datetime
from markitdown import MarkItDown
import zipfile
import shutil
import tempfile
import re

# 定数など
OUTPUT_DIR_NAME = "converted_files"
COMBINED_FILENAME = "All_Files_Combined.txt"

# 判定閾値（目安）
TEXT_PER_VISUAL_THRESHOLD = 300 

def setup_args():
    parser = argparse.ArgumentParser(description='Office files to Markdown converter for NotebookLM')
    parser.add_argument('target_dir', help='Target directory containing Office files or ZIP file')
    parser.add_argument('--combine', action='store_true', help='Combine all converted files into one single text file')
    parser.add_argument('--skip-ppt', action='store_true', help='Skip PowerPoint files (recommend using PDF for visual-heavy PPTs)')
    return parser.parse_args()

def analyze_docx(file_path):
    try:
        doc = docx.Document(file_path)
        visual_count = 0
        char_count = 0
        for para in doc.paragraphs:
            text = para.text.strip()
            char_count += len(text)
            for run in para.runs:
                if run.element.xpath('.//a:blip'):
                     visual_count += 1
        return visual_count, char_count
    except:
        return 0, 0

def analyze_xlsx(file_path):
    try:
        visual_count = 0
        char_count = 0
        try:
            wb_obj = openpyxl.load_workbook(file_path, data_only=True)
            for sheet_name in wb_obj.sheetnames:
                sheet = wb_obj[sheet_name]
                if hasattr(sheet, '_charts') and sheet._charts:
                    visual_count += len(sheet._charts)
        except:
            pass 
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sheet_name in wb.sheetnames:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                if not df.empty:
                    csv_text = df.to_csv(index=False)
                    char_count += len(csv_text)
            except:
                pass
        return visual_count, char_count
    except:
        return 0, 0

def analyze_pptx(file_path):
    try:
        prs = Presentation(file_path)
        visual_count = 0
        char_count = 0
        for slide in prs.slides:
            if slide.shapes.title:
                char_count += len(slide.shapes.title.text.strip())
            for shape in slide.shapes:
                if shape == slide.shapes.title:
                    continue
                if shape.has_text_frame:
                    char_count += len(shape.text_frame.text.strip())
                is_visual = False
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE: is_visual = True
                elif shape.shape_type == MSO_SHAPE_TYPE.GROUP: is_visual = True
                elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                     if not shape.has_text_frame or not shape.text_frame.text.strip(): is_visual = True
                if is_visual: visual_count += 1
        return visual_count, char_count
    except:
        return 0, 0

def convert_with_markitdown(file_path):
    try:
        md = MarkItDown()
        result = md.convert(str(file_path)) 
        if result and result.text_content:
            return result.text_content
        return ""
    except Exception as e:
        print(f"    Error converting {file_path.name}: {e}")
        return None

def sanitize_filename(name):
    """ファイル名として使える文字だけに置換"""
    return re.sub(r'[\\/*?:"<>|]', "", name)

def get_output_filename(root_path, file_path):
    """
    元のフォルダ構造を反映したファイル名を生成する。
    例: root/A/B/file.docx -> A_B_file.md
    """
    try:
        rel_path = file_path.relative_to(root_path)
        # フォルダ区切りをアンダースコア等に置換してフラット化
        flat_name = str(rel_path.with_suffix('')).replace(os.sep, '_')
        return sanitize_filename(flat_name) + ".md"
    except ValueError:
        # root_path外にある場合（稀だが一応）
        return file_path.stem + ".md"

def process_directory(current_path, root_path, output_dir, args, report_items, converted_files_content, processed_zips=None):
    """
    ディレクトリを再帰的に処理する関数
    root_path: 処理の基点となるパス（相対パス計算用）
    """
    if processed_zips is None:
        processed_zips = set()

    # current_pathがファイル(ZIP)の場合
    if current_path.is_file() and current_path.suffix.lower() == '.zip':
        if current_path in processed_zips: return
        processed_zips.add(current_path)
        
        print(f"Extracting ZIP: {current_path.name} ...")
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                with zipfile.ZipFile(current_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                # ZIPの中身を処理するとき、root_pathはその一時フォルダにする（ZIP内構造を維持するため）
                process_directory(Path(temp_dir), Path(temp_dir), output_dir, args, report_items, converted_files_content, processed_zips)
        except Exception as e:
            print(f"Error processing zip {current_path}: {e}")
        return

    # ディレクトリ処理
    for root, dirs, files in os.walk(current_path):
        if OUTPUT_DIR_NAME in root: continue
            
        for file in files:
            file_path = Path(root) / file
            ext = file_path.suffix.lower()
            
            # ZIPファイルの再帰処理
            if ext == '.zip':
                if file_path not in processed_zips:
                    processed_zips.add(file_path)
                    print(f"Extracting nested ZIP: {file} ...")
                    try:
                        with tempfile.TemporaryDirectory() as temp_dir:
                            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                                zip_ref.extractall(temp_dir)
                            # Nested ZIPもその内部構造を維持する
                            process_directory(Path(temp_dir), Path(temp_dir), output_dir, args, report_items, converted_files_content, processed_zips)
                    except Exception as e:
                        print(f"Error processing nested zip {file}: {e}")
                continue

            vis_count = 0
            char_count = 0
            
            # 1. 解析
            if ext == '.docx':
                print(f"Processing: {file}")
                vis_count, char_count = analyze_docx(file_path)
            elif ext == '.xlsx':
                print(f"Processing: {file}")
                vis_count, char_count = analyze_xlsx(file_path)
            elif ext == '.pptx':
                if args.skip_ppt:
                    print(f"Skipping PPT: {file}")
                    continue
                print(f"Processing: {file}")
                vis_count, char_count = analyze_pptx(file_path)
            else:
                continue 

            # 2. レポートデータ
            if vis_count > 0:
                ratio = char_count / vis_count if vis_count > 0 else 9999
                is_dense_visual = ratio < TEXT_PER_VISUAL_THRESHOLD
                if is_dense_visual or vis_count >= 5:
                    report_items.append((file, vis_count, char_count, ratio))

            # 3. 変換
            markdown_content = convert_with_markitdown(file_path)
            
            if markdown_content:
                # 構造を維持したファイル名を生成
                output_filename = get_output_filename(root_path, file_path)
                output_path = output_dir / output_filename
                
                # 相対パス（コンテキスト用）
                try:
                    rel_path_str = str(file_path.relative_to(root_path))
                except:
                    rel_path_str = file.name

                # メタデータヘッダーの作成
                metadata_header = f"""# File Info
- Original Filename: {file}
- Relative Path: {rel_path_str}
- Context: {rel_path_str.replace(os.sep, ' > ')}

---
"""
                final_content = metadata_header + markdown_content

                try:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(final_content)
                    
                    if args.combine:
                        converted_files_content.append(final_content)
                        converted_files_content.append("\n\n---\n\n")
                except Exception as e:
                    print(f"Failed to write {output_path}: {e}")

def main():
    args = setup_args()
    target_path = Path(args.target_dir)
    
    if not target_path.exists():
        print(f"Error: Path '{target_path}' not found.")
        return

    # Output directory
    if target_path.is_dir():
        output_dir = target_path / OUTPUT_DIR_NAME
        root_processing_path = target_path # フォルダ指定ならそのフォルダがroot
    else:
        # File (Zip) case
        output_dir = target_path.parent / OUTPUT_DIR_NAME
        root_processing_path = target_path.parent # 単体ファイルなら親がroot (Zipの場合は内部処理で上書きされる)
        
    output_dir.mkdir(exist_ok=True)
    
    print(f"Target: {target_path}")
    print(f"Output: {output_dir}")
    print("-" * 50)
    
    converted_files_content = []
    report_items = [] 
    
    process_directory(target_path, root_processing_path, output_dir, args, report_items, converted_files_content)

    if args.combine and converted_files_content:
        combined_path = output_dir / COMBINED_FILENAME
        try:
            with open(combined_path, 'w', encoding='utf-8') as f:
                f.write(f"# Combined Output - {datetime.datetime.now()}\n\n")
                f.write("".join(converted_files_content))
        except Exception as e:
            print(f"Error writing combined file: {e}")

    print("\n" + "="*60)
    print(" COMPLETED")
    print("="*60)
    
    if report_items:
        print("\n[!] VISUAL DENSITY REPORT (PDF Recommendation)")
        print(f" {'Filename':<30} | {'Visuals':<7} | {'Density'}")
        print("-" * 60)
        for (fname, v, c, r) in report_items:
             rating = "High Density" if r < TEXT_PER_VISUAL_THRESHOLD else "Many Images"
             print(f" {fname:<30} | {v:>7} | {int(r):>5} ({rating})")
        print("-" * 60)
        print(" * 'High Density' indicates text is sparse relative to visuals. Use PDF for these.")

if __name__ == "__main__":
    main()
