# notebooklm_loader/main.py
"""メイン処理モジュール"""

import os
import shutil
import tempfile
import logging
from pathlib import Path
from typing import List, Tuple, Optional, Set
from tqdm import tqdm

from .config import Config
from .logger import setup_logging, get_logger
from .summary import ProcessingSummary, FileResult
from .merger import MergedOutputManager
from .cli import setup_args
from .utils import get_output_filename, sanitize_content
from .extractors import extract_zip_with_encoding, extract_7z, extract_rar, extract_tar, extract_lzh
from .converters import (
    analyze_docx, analyze_xlsx, analyze_pptx, 
    convert_with_markitdown, convert_image_to_pdf, convert_to_pdf_via_libreoffice
)
from .processors import is_text_file, is_likely_text_by_mime


# 定数
OUTPUT_DIR_NAME = "converted_files"



def process_directory(
    current_path: Path,
    root_path: Path,
    output_dir: Path,
    config: Config,
    report_items: List,
    merger: Optional[MergedOutputManager],
    summary: ProcessingSummary,
    processed_archives: Optional[Set] = None,
    password_protected_files: Optional[List] = None,
    show_progress: bool = True
) -> List[str]:
    """
    ディレクトリを再帰的に処理する
    
    Args:
        current_path: 現在処理中のパス
        root_path: ルートパス
        output_dir: 出力ディレクトリ
        config: 設定オブジェクト
        report_items: レポート項目リスト
        merger: マージマネージャー
        summary: 処理サマリー
        processed_archives: 処理済みアーカイブセット
        password_protected_files: パスワード保護ファイルリスト
        
    Returns:
        パスワード保護ファイルのリスト
    """
    logger = get_logger()
    
    if processed_archives is None:
        processed_archives = set()
    if password_protected_files is None:
        password_protected_files = []

    # アーカイブファイルの場合
    if current_path.is_file():
        ext = current_path.suffix.lower()
        if ext in config.archive_extensions:
            if current_path in processed_archives:
                return password_protected_files
            processed_archives.add(current_path)
            
            logger.info(f"Extracting Archive [{ext}]: {current_path.name} ...")
            try:
                with tempfile.TemporaryDirectory() as temp_dir:
                    result = _extract_archive(current_path, temp_dir, ext)
                    
                    if result == "PASSWORD_PROTECTED":
                        logger.warning(f"    [!] Password protected: {current_path.name}")
                        password_protected_files.append(str(current_path))
                        summary.add_result(FileResult(
                            path=str(current_path),
                            status="password_protected",
                            file_type=ext
                        ))
                    elif result == "OK":
                        process_directory(
                            Path(temp_dir), Path(temp_dir), output_dir, config,
                            report_items, merger, summary, processed_archives, password_protected_files,
                            show_progress=False  # アーカイブ内は進捗表示しない
                        )
            except Exception as e:
                logger.error(f"Error processing archive {current_path}: {e}")
            return password_protected_files

    # ディレクトリ処理 - まずファイル一覧を収集
    all_files = []
    for root, dirs, files in os.walk(current_path):
        if OUTPUT_DIR_NAME in root or "converted_files_merged" in root:
            continue
        for file in files:
            if not file.startswith('.'):
                all_files.append((root, file))
    
    # プログレスバー付きでファイルを処理
    file_iterator = all_files
    if show_progress and all_files:
        file_iterator = tqdm(all_files, desc="Processing files", unit="file", 
                            leave=True, dynamic_ncols=True)
    
    for root, file in file_iterator:
        file_path = Path(root) / file
        
        if show_progress and isinstance(file_iterator, tqdm):
            file_iterator.set_postfix_str(file[:30] + '...' if len(file) > 30 else file)
            
        # 隠しファイルをスキップ（既にフィルタ済みだが念のため）
        if file.startswith('.'):
            continue
            
            # シンボリックリンクをスキップ
            if file_path.is_symlink():
                logger.debug(f"[Skipped Symlink] {file}")
                summary.add_result(FileResult(path=str(file_path), status="skipped", file_type="symlink"))
                continue
            
            # 注: 巨大ファイルはテキストならmergerで自動分割、バイナリならMIME判定でスキップ
            
            ext = file_path.suffix.lower()
            
            # スキップ対象
            if ext in config.skip_extensions:
                logger.debug(f"[Skipped Unsupported] {file}")
                summary.add_result(FileResult(path=str(file_path), status="skipped", file_type=ext))
                continue
            
            # アーカイブファイルの再帰処理
            if ext in config.archive_extensions:
                if file_path not in processed_archives:
                    processed_archives.add(file_path)
                    logger.info(f"Extracting Archive [{ext}]: {file} ...")
                    try:
                        with tempfile.TemporaryDirectory() as temp_dir:
                            result = _extract_archive(file_path, temp_dir, ext)
                            
                            if result == "PASSWORD_PROTECTED":
                                logger.warning(f"    [!] Password protected: {file}")
                                password_protected_files.append(str(file_path))
                                summary.add_result(FileResult(
                                    path=str(file_path),
                                    status="password_protected",
                                    file_type=ext
                                ))
                            elif result == "OK":
                                process_directory(
                                    Path(temp_dir), Path(temp_dir), output_dir, config,
                                    report_items, merger, summary, processed_archives, password_protected_files
                                )
                    except Exception as e:
                        logger.error(f"Error processing archive {file}: {e}")
                continue

            # ファイル処理
            result = _process_single_file(
                file_path, file, ext, root_path, output_dir, config, 
                report_items, merger, summary
            )
    
    return password_protected_files


def _extract_archive(archive_path: Path, extract_to: str, ext: str) -> str:
    """アーカイブを展開"""
    if ext == '.zip':
        return extract_zip_with_encoding(archive_path, extract_to)
    elif ext == '.7z':
        return extract_7z(archive_path, extract_to)
    elif ext == '.rar':
        return extract_rar(archive_path, extract_to)
    elif ext in ['.tar', '.gz', '.tgz']:
        return extract_tar(archive_path, extract_to)
    elif ext == '.lzh':
        return extract_lzh(archive_path, extract_to)
    return "UNSUPPORTED"


def _process_single_file(
    file_path: Path,
    file: str,
    ext: str,
    root_path: Path,
    output_dir: Path,
    config: Config,
    report_items: List,
    merger: Optional[MergedOutputManager],
    summary: ProcessingSummary
) -> bool:
    """単一ファイルを処理"""
    logger = get_logger()
    vis_count = 0
    char_count = 0
    markdown_content = ""

    # 1. 新形式Office (.docx, .xlsx, .pptx)
    if ext == '.docx':
        logger.info(f"Processing: {file}")
        vis_count, char_count = analyze_docx(file_path)
    elif ext == '.xlsx':
        logger.info(f"Processing: {file}")
        vis_count, char_count = analyze_xlsx(file_path)
    elif ext == '.pptx':
        if config.skip_ppt:
            logger.info(f"Skipping PPT: {file}")
            return False
        logger.info(f"Processing: {file}")
        vis_count, char_count = analyze_pptx(file_path)

    # 視覚密度チェック（新形式Office）
    if ext in config.office_extensions_new:
        ratio = char_count / vis_count if vis_count > 0 else 9999
        is_dense_visual = ratio < config.visual_density_threshold
        if is_dense_visual or vis_count >= 5:
            logger.info(f"  [Auto-Switch] High density detected (Visuals: {vis_count}). Converting to PDF...")
            target_pdf_name = get_output_filename(root_path, file_path, extension=".pdf")
            final_pdf_path = output_dir / target_pdf_name
            
            pdf_result = convert_to_pdf_via_libreoffice(file_path, output_dir)
            
            if pdf_result:
                try:
                    if pdf_result != final_pdf_path:
                        if final_pdf_path.exists():
                            final_pdf_path.unlink()
                        pdf_result.rename(final_pdf_path)
                    
                    report_items.append((file, vis_count, char_count, ratio, "Converted to PDF"))
                    logger.info(f"    -> Success: {target_pdf_name}")
                    summary.add_result(FileResult(path=str(file_path), status="converted", output=target_pdf_name, file_type=ext))
                    
                    if merger:
                        try:
                            shutil.copy2(final_pdf_path, merger.output_dir / target_pdf_name)
                        except Exception as e:
                            logger.error(f"Error copying PDF: {e}")
                except Exception as e:
                    logger.error(f"    Error renaming PDF: {e}")
            else:
                logger.warning("    [Fallback] PDF conversion failed.")
                report_items.append((file, vis_count, char_count, ratio, "Kept Original (PDF Fail)"))
            return True
        else:
            markdown_content = convert_with_markitdown(file_path)

    # 2. PDF Files
    elif ext == '.pdf':
        logger.info(f"Copying PDF: {file}")
        output_filename = get_output_filename(root_path, file_path, extension=".pdf")
        try:
            shutil.copy2(file_path, output_dir / output_filename)
            summary.add_result(FileResult(path=str(file_path), status="converted", output=output_filename, file_type=ext))
        except Exception:
            pass
        if merger:
            try:
                shutil.copy2(file_path, merger.output_dir / output_filename)
            except Exception as e:
                logger.error(f"Error copying PDF: {e}")
        return True

    # 3. Legacy Office
    elif ext in config.office_extensions_legacy:
        if ext == '.ppt' and config.skip_ppt:
            logger.info(f"Skipping PPT (Legacy): {file}")
            return False
        logger.info(f"Processing Legacy Office[{ext}]: {file}")
        markdown_content = convert_with_markitdown(file_path)
        if markdown_content is None:
            logger.warning(f"    [Warning] Could not convert: {file}")
            return False

    # 4. MarkItDown対応形式
    elif ext in config.markitdown_extensions:
        logger.info(f"Processing MarkItDown[{ext}]: {file}")
        markdown_content = convert_with_markitdown(file_path)
        if markdown_content is None:
            logger.warning(f"    [Warning] Could not convert: {file}")
            return False

    # 5. Visio
    elif ext in config.visio_extensions:
        logger.info(f"Processing Visio: {file}")
        target_pdf_name = get_output_filename(root_path, file_path, extension=".pdf")
        final_pdf_path = output_dir / target_pdf_name
        
        pdf_result = convert_to_pdf_via_libreoffice(file_path, output_dir)
        if pdf_result:
            try:
                if pdf_result != final_pdf_path:
                    if final_pdf_path.exists():
                        final_pdf_path.unlink()
                    pdf_result.rename(final_pdf_path)
                logger.info(f"    -> Success: {target_pdf_name}")
                summary.add_result(FileResult(path=str(file_path), status="converted", output=target_pdf_name, file_type=ext))
                if merger:
                    try:
                        shutil.copy2(final_pdf_path, merger.output_dir / target_pdf_name)
                    except Exception as e:
                        logger.error(f"Error copying Visio PDF: {e}")
            except Exception as e:
                logger.error(f"    Error renaming Visio PDF: {e}")
        else:
            logger.warning(f"    [Warning] Could not convert Visio: {file}")
        return True

    # 6. 画像
    elif ext in config.image_extensions:
        logger.info(f"Processing Image: {file}")
        target_pdf_name = get_output_filename(root_path, file_path, extension=".pdf")
        final_pdf_path = output_dir / target_pdf_name
        
        pdf_result = convert_image_to_pdf(file_path, output_dir)
        if pdf_result:
            try:
                if pdf_result != final_pdf_path:
                    if final_pdf_path.exists():
                        final_pdf_path.unlink()
                    pdf_result.rename(final_pdf_path)
                logger.info(f"    -> Success: {target_pdf_name}")
                summary.add_result(FileResult(path=str(file_path), status="converted", output=target_pdf_name, file_type=ext))
                if merger:
                    try:
                        shutil.copy2(final_pdf_path, merger.output_dir / target_pdf_name)
                    except Exception as e:
                        logger.error(f"Error copying image PDF: {e}")
            except Exception as e:
                logger.error(f"    Error renaming image PDF: {e}")
        else:
            logger.warning(f"    [Warning] Could not convert image: {file}")
        return True

    # 7. テキストファイル
    elif ext not in config.office_extensions_all and ext != '.pdf':
        mime_is_text = is_likely_text_by_mime(file_path)
        is_known_text = ext in config.text_extensions
        detected_encoding = None
        
        if is_known_text or mime_is_text == True:
            is_readable, detected_encoding = is_text_file(file_path)
            if not detected_encoding:
                detected_encoding = 'utf-8'
        elif mime_is_text == False:
            logger.debug(f"[Skipped Binary] {file}")
            summary.add_result(FileResult(path=str(file_path), status="skipped", file_type="binary"))
            return False
        else:
            is_readable, detected_encoding = is_text_file(file_path)
            if not is_readable:
                logger.debug(f"[Skipped Binary] {file}")
                summary.add_result(FileResult(path=str(file_path), status="skipped", file_type="binary"))
                return False
        
        logger.info(f"Processing Text[{ext}] ({detected_encoding}): {file}")
        try:
            with open(file_path, 'r', encoding=detected_encoding, errors='replace') as f:
                markdown_content = f.read()
                # 不可視文字（ゼロ幅スペース等）を除去
                markdown_content = sanitize_content(markdown_content)
                if not markdown_content.strip():
                    markdown_content = "(Empty File)"
        except Exception as e:
            logger.error(f"Error reading text file {file}: {e}")
            markdown_content = ""

    # Markdown出力
    if markdown_content:
        # 全てのコンテンツから不可視文字を除去（MarkItDown変換後も含む）
        markdown_content = sanitize_content(markdown_content)
        output_filename = get_output_filename(root_path, file_path, extension=".md")
        output_path = output_dir / output_filename
        
        try:
            rel_path_str = str(file_path.relative_to(root_path))
        except ValueError:
            rel_path_str = file if isinstance(file, str) else file.name

        metadata_header = f"""# File Info
- Original Filename: {file}
- Relative Path: {rel_path_str}
- Context: {rel_path_str.replace(os.sep, ' > ')}

---
"""
        final_content = metadata_header + markdown_content + "\n\n---\n\n"

        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(final_content)
            summary.add_result(FileResult(path=str(file_path), status="converted", output=output_filename, file_type=ext))
        except Exception as e:
            logger.error(f"Failed to write {output_path}: {e}")
            summary.add_result(FileResult(path=str(file_path), status="error", error_message=str(e), file_type=ext))
        
        if merger:
            merger.add_content(output_filename, final_content)
    
    return True


def run() -> int:
    """
    メインエントリーポイント
    
    Returns:
        終了コード（0=成功）
    """
    args = setup_args()
    target_path = Path(args.target_dir)
    
    if not target_path.exists():
        print(f"Error: Path '{target_path}' not found.")
        return 1

    # 設定作成
    config = Config.from_args(args)

    # 出力ディレクトリ設定
    if target_path.is_dir():
        output_dir = target_path / OUTPUT_DIR_NAME
        root_processing_path = target_path
    else:
        output_dir = target_path.parent / OUTPUT_DIR_NAME
        root_processing_path = target_path.parent
        
    output_dir.mkdir(exist_ok=True)

    # ログ設定
    logger = setup_logging(output_dir, verbose=config.verbose)
    
    if config.quiet:
        for handler in logger.handlers:
            if isinstance(handler, logging.StreamHandler) and not isinstance(handler, logging.FileHandler):
                handler.setLevel(logging.CRITICAL)
    
    # 処理サマリー初期化
    summary = ProcessingSummary(target_path=str(target_path))
    
    # Dry-run モード
    if config.dry_run:
        logger.info("=== DRY-RUN MODE ===")
        logger.info("Following files would be processed:")
        for root, dirs, files in os.walk(target_path):
            for file in files:
                logger.info(f"  - {Path(root) / file}")
        logger.info("Dry-run complete. No files were actually processed.")
        return 0

    # Merged Output Setup
    merger = None
    if config.merge:
        if target_path.is_dir():
            merged_dir = target_path / (OUTPUT_DIR_NAME + "_merged")
        else:
            merged_dir = target_path.parent / (OUTPUT_DIR_NAME + "_merged")
        
        logger.info(f"Target: {target_path}")
        logger.info(f"Output: {output_dir}")
        logger.info(f"Merged: {merged_dir}")
        logger.info("-" * 50)
        
        merger = MergedOutputManager(merged_dir, max_chars_per_volume=config.max_chars_per_volume)
    else:
        logger.info(f"Target: {target_path}")
        logger.info(f"Output: {output_dir}")
        logger.info("(Use --merge to create combined output)")
        logger.info("-" * 50)
    
    report_items = []
    password_protected_files = []
    
    password_protected_files = process_directory(
        target_path, root_processing_path, output_dir, config,
        report_items, merger, summary, password_protected_files=password_protected_files,
        show_progress=not config.quiet  # quietモード時はプログレスバー無効
    )
    
    # Finalize Merge
    if merger:
        merger.finalize()

    # サマリー保存
    summary_file = summary.save(output_dir)

    logger.info("")
    logger.info("=" * 60)
    logger.info(" COMPLETED")
    logger.info("=" * 60)
    
    # 統計サマリー
    logger.info(f"\nProcessing Summary:")
    logger.info(f"  Total files:        {summary.total_files}")
    logger.info(f"  Processed:          {summary.processed}")
    logger.info(f"  Skipped:            {summary.skipped}")
    logger.info(f"  Errors:             {summary.errors}")
    logger.info(f"  Password protected: {summary.password_protected}")
    logger.info(f"\nReport saved to: {summary_file}")
    
    # パスワード保護ファイルレポート
    if password_protected_files:
        logger.warning("\n[!] PASSWORD PROTECTED FILES (Could not process)")
        logger.warning("-" * 60)
        for pf in password_protected_files:
            logger.warning(f"  - {pf}")
        logger.warning("-" * 60)
        logger.warning(f"  Total: {len(password_protected_files)} file(s)")
    
    if report_items:
        logger.info("\n[!] PROCESSED FILE REPORT (Visual Density)")
        logger.info(f" {'Filename':<40} | {'Visuals':<7} | {'Density':<7} | {'Status'}")
        logger.info("-" * 90)
        for (fname, v, c, r, status) in report_items:
            rating = "High" if r < config.visual_density_threshold else "Low"
            logger.info(f" {fname:<40} | {v:>7} | {int(r):>5} ({rating}) | {status}")
        logger.info("-" * 90)
    
    return 0
