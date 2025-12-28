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
import subprocess

# ---------------------------------------------------------
# PDF Conversion Utilities
# ---------------------------------------------------------

def convert_to_pdf_via_libreoffice(input_path, output_dir_path):
    """
    LibreOffice (soffice) を使用してPDF変換を行う。
    Mac: /Applications/LibreOffice.app/Contents/MacOS/soffice
    """
    soffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not os.path.exists(soffice_path):
        soffice_path = "soffice" # Try PATH

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir_path),
        str(input_path)
    ]
    try:
        # Run command
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        # Output filename check (LibreOffice creates filename.pdf)
        original_stem = input_path.stem
        
        generated_pdf = output_dir_path / (original_stem + ".pdf")
        if generated_pdf.exists():
            return generated_pdf
        return None
    except Exception as e:
        print(f"    [PDF Convert Error] {e}")
        return None

# 定数など
OUTPUT_DIR_NAME = "converted_files"
COMBINED_FILENAME = "All_Files_Combined.txt"

# 判定閾値（目安）
TEXT_PER_VISUAL_THRESHOLD = 300 

TEXT_EXTENSIONS = {
    '.txt', '.md', '.py', '.js', '.jsx', '.ts', '.tsx', '.html', '.css', '.json', 
    '.yaml', '.yml', '.org', '.sh', '.bat', '.zsh', '.rb', '.java', '.c', '.cpp', 
    '.h', '.go', '.rs', '.php', '.pl', '.swift', '.kt', '.sql', '.xml', '.csv'
}

def setup_args():
    parser = argparse.ArgumentParser(description='Office files to Markdown converter for NotebookLM')
    parser.add_argument('target_dir', help='Target directory containing Office files or ZIP file')
    parser.add_argument('--merge', action='store_true', help='Also create merged output in converted_files_merged directory')
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

def get_output_filename(root_path, file_path, extension=".md"):
    """
    元のフォルダ構造を反映したファイル名を生成する。
    例: root/A/B/file.docx -> A_B_file.md
    """
    try:
        rel_path = file_path.relative_to(root_path)
        # フォルダ区切りをアンダースコア等に置換してフラット化
        flat_name = str(rel_path.with_suffix('')).replace(os.sep, '_')
        return sanitize_filename(flat_name) + extension
    except ValueError:
        # root_path外にある場合（稀だが一応）
        return file_path.stem + extension


def extract_zip_with_encoding(zip_path, extract_to):
    """
    ZIPファイルを解凍する際、Windows等で作成されたShift-JISのファイル名を
    正しく復元して展開する。
    """
    with zipfile.ZipFile(zip_path, 'r') as z:
        for file_info in z.infolist():
            filename = file_info.filename
            
            # UTF-8フラグが立っていない場合、エンコーディングの補正を試みる
            if file_info.flag_bits & 0x800 == 0:
                try:
                    # Windows (Japanese) ZIP is often CP932 encoded but marked as CP437
                    filename = filename.encode('cp437').decode('cp932')
                except:
                    pass
            
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


# ---------------------------------------------------------
# Feature: Smart Chunking & Merged Output
# ---------------------------------------------------------

MAX_CHARS_PER_VOLUME = 12000000  # Approx 40MB limit (Safe for UTF-8 Japanese)

class MergedOutputManager:
    def __init__(self, output_dir):
        self.output_dir = output_dir
        self.output_dir.mkdir(exist_ok=True)
        self.current_vol = 1
        self.current_content = []
        self.current_char_count = 0
        self.file_index = [] # List of filenames included in current vol

    def add_content(self, filename, content):
        """
        コンテンツを追加する。
        もしコンテンツ単体で制限を超える場合は、分割して追加する（Recursive Split）。
        追加によって制限を超える場合は、現在のVolを書き出して次へ行く。
        """
        content_len = len(content)

        # Case 1: Huge single file -> Recursive Split
        # (ヘッダ込みで計算済みだが、念のためここでもチェックする論理は変えないが、handle_huge_file内で厳密計算する)
        if content_len > MAX_CHARS_PER_VOLUME:
            self._handle_huge_file(filename, content)
            return

        # Case 2: Buffer overflow -> Flush and define new volume
        if self.current_char_count + content_len > MAX_CHARS_PER_VOLUME:
            self._flush_volume()

        # Normal Add
        self.current_content.append(content)
        self.current_char_count += content_len
        self.file_index.append(filename)

    def _handle_huge_file(self, filename, content):
        """巨大ファイルを分割して登録する"""
        # content は既にヘッダーがついている状態だが、
        # ここでは再分割するため、ヘッダーを除去して本文だけ取り出したいところだが、
        # 引数の content は 'final_content' であり、メタデータを含んでいる。
        # 厳密にはメタデータごと分割するのは変なので、
        # 本文だけ抽出するよりは、渡される前の raw content を引数に取るべきだが、
        # 設計上、Markdown変換後の最終テキストを受け取っているので、
        # ここでは「巨大なテキスト塊」として扱い、強制分割する方針とする。
        
        remaining = content
        part_num = 1
        
        while remaining:
            # 次のパート用ヘッダー（概算サイズ）
            part_header = f"\n\n# {filename} (Part {part_num})\n\n"
            header_len = len(part_header)
            
            # 残りの容量
            available_space = MAX_CHARS_PER_VOLUME - header_len
            
            # 現在のバッファに空きがなければフラッシュ
            if self.current_char_count + header_len > MAX_CHARS_PER_VOLUME: # ヘッダすら入らないならフラッシュ
                 self._flush_volume()
            
            # 再計算（フラッシュ後）
            available_space = MAX_CHARS_PER_VOLUME - self.current_char_count - header_len
            
            if len(remaining) > available_space:
                # カット
                c_chunk = remaining[:available_space]
                remaining = remaining[available_space:]
            else:
                c_chunk = remaining
                remaining = ""
            
            full_chunk = part_header + c_chunk
            
            # ここまできたら必ず入るはず
            self.current_content.append(full_chunk)
            self.file_index.append(f"{filename} (Part {part_num})")
            self.current_char_count += len(full_chunk)
            
            # 満杯になったらフラッシュ
            if self.current_char_count >= MAX_CHARS_PER_VOLUME:
                self._flush_volume()
            
            part_num += 1

    def _flush_volume(self):
        """現在のバッファをファイルに書き出す"""
        if not self.current_content:
            return

        if self.current_content: # Double check
            vol_filename = f"Merged_Files_Vol{self.current_vol:02d}.md"
            output_path = self.output_dir / vol_filename
            
            # 目次生成
            index_text = "# Table of Contents\n" + "\n".join([f"- {name}" for name in self.file_index]) + "\n\n---\n\n"
            
            full_text = index_text + "\n".join(self.current_content)
            
            try:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(full_text)
                print(f"[Merged Created] {vol_filename} ({len(full_text)} chars)")
            except Exception as e:
                print(f"Error writing volume {vol_filename}: {e}")

            # Reset
            self.current_vol += 1
            self.current_content = []
            self.current_char_count = 0
            self.file_index = []

    def finalize(self):
        """最後に残っているバッファを書き出す"""
        self._flush_volume()


def is_text_file(file_path):
    """拡張子に関わらず、UTF-8テキストとして読めるか判定する"""
    try:
        # 先頭4KB程度読んで判定
        with open(file_path, 'r', encoding='utf-8') as f:
             f.read(4000)
        return True
    except UnicodeDecodeError:
        return False
    except Exception:
        return False


def process_directory(current_path, root_path, output_dir, args, report_items, merger, processed_zips=None):
    """
    ディレクトリを再帰的に処理する関数
    merger: MergedOutputManager instance (None if merge disabled)
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
                extract_zip_with_encoding(current_path, temp_dir)
                process_directory(Path(temp_dir), Path(temp_dir), output_dir, args, report_items, merger, processed_zips)
        except Exception as e:
            print(f"Error processing zip {current_path}: {e}")
        return

    # ディレクトリ処理
    for root, dirs, files in os.walk(current_path):
        if OUTPUT_DIR_NAME in root or "converted_files_merged" in root: continue
            
        for file in files:
            file_path = Path(root) / file
            # Rename hidden files or specific system files if needed, but 'is_text_file' helps.
            if file.startswith('.'): continue
            
            ext = file_path.suffix.lower()
            
            # ZIPファイルの再帰処理
            if ext == '.zip':
                if file_path not in processed_zips:
                    processed_zips.add(file_path)
                    print(f"Extracting nested ZIP: {file} ...")
                    try:
                        with tempfile.TemporaryDirectory() as temp_dir:
                            extract_zip_with_encoding(file_path, temp_dir)
                            process_directory(Path(temp_dir), Path(temp_dir), output_dir, args, report_items, merger, processed_zips)
                    except Exception as e:
                        print(f"Error processing nested zip {file}: {e}")
                continue

            vis_count = 0
            char_count = 0
            markdown_content = ""
            
            # --- File Type Handling ---

            # Determine High Density first (Shared Logic)
            is_high_density = False
            
            # 1. Office Files
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
                # Not an office file we analyze for density
                pass

            # Check Density if it was analyzed
            if ext in ['.docx', '.xlsx', '.pptx']:
                ratio = char_count / vis_count if vis_count > 0 else 9999
                is_dense_visual = ratio < TEXT_PER_VISUAL_THRESHOLD
                if is_dense_visual or vis_count >= 5:
                    is_high_density = True
                    print(f"  [Auto-Switch] High density detected (Visuals: {vis_count}). Converting to PDF...")
                    
                    # Target Filename (Flat)
                    target_pdf_name = get_output_filename(root_path, file_path, extension=".pdf")
                    final_pdf_path = output_dir / target_pdf_name
                    
                    # Try PDF Conversion
                    pdf_result = convert_to_pdf_via_libreoffice(file_path, output_dir) # Temp output to output_dir
                    
                    if pdf_result:
                        # Rename/Move to correct flat filename
                        # Since we outputted to output_dir, the file is likely there as 'original_name.pdf'
                        # We need to rename it to flat name 'Folder_Sub_original.pdf'
                        try:
                            if pdf_result != final_pdf_path:
                                if final_pdf_path.exists(): final_pdf_path.unlink()
                                pdf_result.rename(final_pdf_path)
                            
                            report_items.append((file, vis_count, char_count, ratio, f"Converted to PDF"))
                            print(f"    -> Success: {target_pdf_name}")
                            
                            # Copy to Merged folder (if enabled)
                            if merger:
                                 try:
                                    shutil.copy2(final_pdf_path, merger.output_dir / target_pdf_name)
                                 except Exception as e:
                                    print(f"Error copying converted PDF to merged dir: {e}")

                        except Exception as e:
                            print(f"    Error renaming PDF: {e}")
                            # Fallback if rename fails? Unlikely.
                    else:
                        # Fallback to Copy Original
                        print("    [Fallback] PDF conversion failed. Copying original file.")
                        report_items.append((file, vis_count, char_count, ratio, "Kept Original (PDF Fail)"))
                        
                        orig_out_name = get_output_filename(root_path, file_path, extension=ext)
                        
                        try:
                            shutil.copy2(file_path, output_dir / orig_out_name)
                        except Exception: pass
                        
                        if merger:
                             try:
                                shutil.copy2(file_path, merger.output_dir / orig_out_name)
                             except Exception: pass
                    
                    continue # Skip Markdown conversion

                else:
                    # Low Density -> Proceed completion
                    # We need to run conversion now since we only analyzed so far
                    markdown_content = convert_with_markitdown(file_path)

            # 2. PDF Files (Direct Copy -> Merged Dir)
            if ext == '.pdf':
                print(f"Copying PDF: {file}")
                output_filename = get_output_filename(root_path, file_path, extension=".pdf")
                
                # Copy to Single folder
                try:
                    shutil.copy2(file_path, output_dir / output_filename)
                except Exception:
                    pass

                # Copy to Merged folder (if enabled)
                if merger:
                     try:
                        shutil.copy2(file_path, merger.output_dir / output_filename)
                     except Exception as e:
                        print(f"Error copying PDF to merged dir: {e}")
                
                continue 

            # 3. Universal Text Loader (Any text file)
            # (Strictly else if not handled above. Note: Office files handled above continue or set markdown_content)
            elif ext not in ['.docx', '.xlsx', '.pptx', '.pdf']:
                # Check known extensions OR try to read as UTF-8
                is_known_text = ext in TEXT_EXTENSIONS
                is_readable_text = False
                if not is_known_text:
                    if is_text_file(file_path):
                        is_readable_text = True
                
                if is_known_text or is_readable_text:
                    print(f"Processing Text[{ext}]: {file}")
                    try:
                        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                            markdown_content = f.read()
                            if not markdown_content.strip(): markdown_content = "(Empty File)"
                    except Exception as e:
                        print(f"Error reading text file {file}: {e}")
                        markdown_content = ""
                else:
                    # Binary or unknown -> Log and Skip
                    print(f"[Skipped Binary] {file}")
                    continue

            # --- Post-Processing (Report & Write) ---

            # Report Logic (For converted files that might still have some visuals)
            if vis_count > 0 and ext in ['.docx', '.xlsx', '.pptx']:
                 # If we are here, it wasn't high density enough to switch, but maybe worth noting?
                 # Actually, let's strictly log the ones we converted too if they had visuals
                 ratio = char_count / vis_count if vis_count > 0 else 9999
                 report_items.append((file, vis_count, char_count, ratio, "Converted"))

            if markdown_content:
                # 構造を維持したファイル名を生成
                output_filename = get_output_filename(root_path, file_path, extension=".md")
                output_path = output_dir / output_filename
                
                try:
                    rel_path_str = str(file_path.relative_to(root_path))
                except:
                    rel_path_str = file if isinstance(file, str) else file.name

                metadata_header = f"""# File Info
- Original Filename: {file}
- Relative Path: {rel_path_str}
- Context: {rel_path_str.replace(os.sep, ' > ')}

---
"""
                final_content = metadata_header + markdown_content + "\n\n---\n\n"

                # 1. Write individual file (Always)
                try:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(final_content)
                except Exception as e:
                    print(f"Failed to write individual {output_path}: {e}")
                
                # 2. Add to Smart Merger (Only if enabled)
                if merger:
                    merger.add_content(output_filename, final_content)


def main():
    args = setup_args()
    target_path = Path(args.target_dir)
    
    if not target_path.exists():
        print(f"Error: Path '{target_path}' not found.")
        return

    if target_path.is_dir():
        output_dir = target_path / OUTPUT_DIR_NAME
        root_processing_path = target_path
    else:
        output_dir = target_path.parent / OUTPUT_DIR_NAME
        root_processing_path = target_path.parent
        
    output_dir.mkdir(exist_ok=True)

    # Merged Output Setup (Only if --merge is specified)
    merger = None
    if args.merge:
        if target_path.is_dir():
            merged_dir = target_path / (OUTPUT_DIR_NAME + "_merged")
        else:
            merged_dir = target_path.parent / (OUTPUT_DIR_NAME + "_merged")
        
        print(f"\nTarget: {target_path}")
        print(f"Output: {output_dir}")
        print(f"Merged: {merged_dir}")
        print("-" * 50)
        
        merger = MergedOutputManager(merged_dir)
    else:
        print(f"\nTarget: {target_path}")
        print(f"Output: {output_dir}")
        print("(Use --merge to create combined output)")
        print("-" * 50)
    
    report_items = [] 
    
    process_directory(target_path, root_processing_path, output_dir, args, report_items, merger)
    
    # Finalize Merge
    if merger:
        merger.finalize()

    print("\n" + "="*60)
    print(" COMPLETED")
    print("="*60)
    
    if report_items:
        print("\n[!] PROCESSED FILE REPORT (Visual Density)")
        print(f" {'Filename':<40} | {'Visuals':<7} | {'Density':<7} | {'Status'}")
        print("-" * 90)
        for (fname, v, c, r, status) in report_items:
             rating = "High" if r < TEXT_PER_VISUAL_THRESHOLD else "Low"
             print(f" {fname:<40} | {v:>7} | {int(r):>5} ({rating}) | {status}")
        print("-" * 90)
        print(" * 'Converted to PDF': High visual density -> Auto-converted to PDF via LibreOffice.")
        print(" * 'Kept Original (PDF Fail)': PDF Conversion failed -> Copied original file.")



if __name__ == "__main__":
    main()
