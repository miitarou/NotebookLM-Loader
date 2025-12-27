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

# 定数など
OUTPUT_DIR_NAME = "converted_files"
COMBINED_FILENAME = "All_Files_Combined.txt"

# 判定閾値（目安）
# 画像1枚あたりの文字数がこれ以下なら「図解中心」とみなす
TEXT_PER_VISUAL_THRESHOLD = 300 

def setup_args():
    parser = argparse.ArgumentParser(description='Office files to Markdown converter for NotebookLM')
    parser.add_argument('target_dir', help='Target directory containing Office files')
    parser.add_argument('--combine', action='store_true', help='Combine all converted files into one single text file')
    parser.add_argument('--skip-ppt', action='store_true', help='Skip PowerPoint files (recommend using PDF for visual-heavy PPTs)')
    return parser.parse_args()

def get_image_alt_text(shape):
    """画像の代替テキストを取得する（可能な場合）"""
    try:
        return shape.name or "Image"
    except:
        return "Image"

def analyze_docx(file_path):
    """Wordファイルの視覚要素密度を解析する。戻り値: (visual_count, char_count)"""
    try:
        doc = docx.Document(file_path)
        visual_count = 0
        char_count = 0
        
        for para in doc.paragraphs:
            text = para.text.strip()
            char_count += len(text)
            
            # 画像（InlineShape）の簡易検出
            for run in para.runs:
                if run.element.xpath('.//a:blip'):
                     visual_count += 1
        
        return visual_count, char_count
    except Exception as e:
        print(f"Error analyzing docx {file_path}: {e}")
        return 0, 0

def analyze_xlsx(file_path):
    """Excelファイルの視覚要素密度を解析する。戻り値: (visual_count, char_count)"""
    try:
        visual_count = 0
        char_count = 0
        
        # チャート検出用
        try:
            wb_obj = openpyxl.load_workbook(file_path, data_only=True)
            for sheet_name in wb_obj.sheetnames:
                sheet = wb_obj[sheet_name]
                if hasattr(sheet, '_charts') and sheet._charts:
                    visual_count += len(sheet._charts)
        except:
            pass 

        # データ読み込み用 (文字数カウント)
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
    except Exception as e:
        print(f"Error analyzing xlsx {file_path}: {e}")
        return 0, 0

def analyze_pptx(file_path):
    """PPTファイルの視覚要素密度を解析する。戻り値: (visual_count, char_count)"""
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
                
                # 画像・図形のカウント
                is_visual = False
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    is_visual = True
                elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    is_visual = True
                elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                     if not shape.has_text_frame or not shape.text_frame.text.strip():
                         is_visual = True
                
                if is_visual:
                    visual_count += 1
        
        return visual_count, char_count
    except Exception as e:
        print(f"Error analyzing pptx {file_path}: {e}")
        return 0, 0

def convert_with_markitdown(file_path):
    """MarkItDownを使ってファイルをMarkdownに変換する"""
    try:
        md = MarkItDown()
        result = md.convert(str(file_path)) # MarkItDown takes file path as string or LocalFileSource
        if result and result.text_content:
            return result.text_content
        return ""
    except Exception as e:
        print(f"Error converting with MarkItDown {file_path}: {e}")
        return None

def main():
    args = setup_args()
    target_path = Path(args.target_dir)
    
    if not target_path.exists():
        print(f"Error: Directory '{target_path}' not found.")
        return

    output_dir = target_path / OUTPUT_DIR_NAME
    output_dir.mkdir(exist_ok=True)
    
    print(f"Start converting files in: {target_path}")
    print(f"Using Engine: Microsoft MarkItDown")
    print(f"Output directory: {output_dir}")
    
    converted_files_content = []
    # (filename, visual_count, char_count, ratio)
    report_items = [] 

    for root, dirs, files in os.walk(target_path):
        if OUTPUT_DIR_NAME in root:
            continue
            
        for file in files:
            file_path = Path(root) / file
            ext = file_path.suffix.lower()
            
            markdown_content = None
            vis_count = 0
            char_count = 0
            
            # 1. まず解析 (Analyze)
            if ext == '.docx':
                print(f"Analyzing: {file} ...")
                vis_count, char_count = analyze_docx(file_path)
            elif ext == '.xlsx':
                print(f"Analyzing: {file} ...")
                vis_count, char_count = analyze_xlsx(file_path)
            elif ext == '.pptx':
                if args.skip_ppt:
                    print(f"Skipping PPT (as requested): {file}")
                    continue
                print(f"Analyzing: {file} ...")
                vis_count, char_count = analyze_pptx(file_path)
            else:
                continue # 対象外の拡張子

            # 2. レポートデータの蓄積
            if vis_count > 0:
                ratio = char_count / vis_count if vis_count > 0 else 9999
                is_dense_visual = ratio < TEXT_PER_VISUAL_THRESHOLD
                
                if is_dense_visual or vis_count >= 5:
                    report_items.append((file, vis_count, char_count, ratio))

            # 3. 変換 (Convert) - MarkItDownを利用
            # PPTスキップが指定されていて、かつPPTの場合は既にcontinueされている
            # ここでは「解析結果にかかわらず、とりあえずMD変換は行う」（ユーザーが選べるように）
            print(f"  Converting with MarkItDown...")
            markdown_content = convert_with_markitdown(file_path)
            
            if markdown_content:
                output_filename = f"{file_path.stem}.md"
                output_path = output_dir / output_filename
                try:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        # ヘッダーを付ける
                        f.write(f"# File: {file}\n\n")
                        f.write(markdown_content)
                    
                    if args.combine:
                        converted_files_content.append(f"# File: {file}\n\n")
                        converted_files_content.append(markdown_content)
                        converted_files_content.append("\n\n---\n\n")
                except Exception as e:
                    print(f"Failed to write disk {output_path}: {e}")

    if args.combine and converted_files_content:
        combined_path = output_dir / COMBINED_FILENAME
        try:
            with open(combined_path, 'w', encoding='utf-8') as f:
                f.write(f"# All Files Combined - {datetime.datetime.now()}\n\n")
                f.write("".join(converted_files_content))
        except Exception as e:
            print(f"Error creating combined file: {e}")

    print("\n" + "="*60)
    print(" CONVERSION COMPLETED (Powered by MarkItDown)")
    print("="*60)
    
    if report_items:
        print("\n[!] VISUAL CONTENT REPORT")
        print("以下のファイルは「テキストに対して画像/図解の比率が高い」か「画像が大量」です。")
        print("Markdown変換では文脈が損なわれる可能性があるため、")
        print("『PDF形式』でNotebookLMにアップロードすることを強く推奨します。")
        print("-" * 75)
        print(f" {'Filename':<30} | {'Visuals':<7} | {'Chars':<7} | {'Chars/Vis (Density)'}")
        print("-" * 75)
        for (fname, v, c, r) in report_items:
             rating = "High Visual Density" if r < TEXT_PER_VISUAL_THRESHOLD else "Many Images"
             print(f" {fname:<30} | {v:>7} | {c:>7} | {int(r):>5} ({rating})")
        print("-" * 75)
        print(f" (Threshold: < {TEXT_PER_VISUAL_THRESHOLD} chars per visual is considered 'Visual Heavy')\n")

if __name__ == "__main__":
    main()
