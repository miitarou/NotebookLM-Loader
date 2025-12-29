# notebooklm_loader/merger.py
"""Smart Chunking & Merged Outputモジュール"""

from pathlib import Path
from typing import List

from .logger import get_logger

# 安全弁: 1ファイルあたりの最大分割数
MAX_PARTS = 10000


class MergedOutputManager:
    """
    ファイルを結合してNotebookLM用に最適化されたサイズで出力する
    
    Attributes:
        output_dir: 出力ディレクトリ
        max_chars_per_volume: ボリュームあたりの最大文字数
        current_vol: 現在のボリューム番号
    """
    
    def __init__(self, output_dir: Path, max_chars_per_volume: int = 10500000):
        """
        初期化
        
        Args:
            output_dir: 出力ディレクトリ
            max_chars_per_volume: ボリュームあたりの最大文字数（デフォルト約35MB）
        """
        self.output_dir = output_dir
        self.output_dir.mkdir(exist_ok=True)
        self.max_chars_per_volume = max_chars_per_volume
        self.current_vol = 1
        self.current_content: List[str] = []
        self.current_char_count = 0
        self.file_index: List[str] = []

    def add_content(self, filename: str, content: str):
        """
        コンテンツを追加する
        
        Args:
            filename: ファイル名
            content: コンテンツ
        """
        content_len = len(content)

        # 巨大ファイルの場合は分割
        if content_len > self.max_chars_per_volume:
            self._handle_huge_file(filename, content)
            return

        # バッファオーバーフローの場合はフラッシュ
        if self.current_char_count + content_len > self.max_chars_per_volume:
            self._flush_volume()

        self.current_content.append(content)
        self.current_char_count += content_len
        self.file_index.append(filename)

    def _handle_huge_file(self, filename: str, content: str):
        """
        巨大ファイルを行単位で分割して登録する
        
        文字の途中で切断されないよう、改行位置で分割を行う。
        これによりマルチバイト文字（日本語など）が途中で切れることを防ぐ。
        """
        logger = get_logger()
        remaining = content
        part_num = 1
        
        while remaining and part_num <= MAX_PARTS:
            part_header = f"\n\n# {filename} (Part {part_num})\n\n"
            header_len = len(part_header)
            
            if self.current_char_count + header_len > self.max_chars_per_volume:
                 self._flush_volume()
            
            available_space = max(1, self.max_chars_per_volume - self.current_char_count - header_len)
            
            if len(remaining) > available_space:
                # 行単位で分割：available_space以内で最後の改行位置を探す
                split_pos = remaining.rfind('\n', 0, available_space)
                
                if split_pos == -1:
                    # 改行が見つからない場合、CSVのレコード境界（カンマ+改行相当）を探す
                    # CSV形式では各フィールドがカンマで区切られているため、
                    # カンマの後で分割すればレコードの途中で切れにくい
                    split_pos = remaining.rfind(',\n', 0, available_space)
                    if split_pos != -1:
                        split_pos += 2  # カンマと改行の次から
                
                if split_pos == -1:
                    # TSV（タブ区切り）のレコード境界を探す
                    split_pos = remaining.rfind('\t\n', 0, available_space)
                    if split_pos != -1:
                        split_pos += 2  # タブと改行の次から
                
                if split_pos == -1:
                    # カンマ+改行もない場合、カンマのみを探す
                    split_pos = remaining.rfind(',', 0, available_space)
                    if split_pos != -1:
                        split_pos += 1  # カンマの次から
                
                if split_pos == -1:
                    # タブのみを探す（TSV対応）
                    split_pos = remaining.rfind('\t', 0, available_space)
                    if split_pos != -1:
                        split_pos += 1  # タブの次から
                
                if split_pos == -1:
                    # カンマ・タブも見つからない場合、スペースで分割
                    split_pos = remaining.rfind(' ', 0, available_space)
                    if split_pos != -1:
                        split_pos += 1
                
                if split_pos == -1 or split_pos == 0:
                    # どれも見つからない場合は、そのまま分割（最終手段）
                    split_pos = available_space
                
                c_chunk = remaining[:split_pos]
                remaining = remaining[split_pos:]
            else:
                c_chunk = remaining
                remaining = ""
            
            full_chunk = part_header + c_chunk
            
            self.current_content.append(full_chunk)
            self.file_index.append(f"{filename} (Part {part_num})")
            self.current_char_count += len(full_chunk)
            
            if self.current_char_count >= self.max_chars_per_volume:
                self._flush_volume()
            
            part_num += 1
        
        if part_num > MAX_PARTS:
            logger.warning(f"Max parts ({MAX_PARTS}) reached for {filename}. File may be truncated.")

    def _flush_volume(self):
        """現在のバッファをファイルに書き出す"""
        if not self.current_content:
            return

        vol_filename = f"Merged_Files_Vol{self.current_vol:02d}.md"
        output_path = self.output_dir / vol_filename
        
        # 目次生成
        index_text = "# Table of Contents\n" + "\n".join([f"- {name}" for name in self.file_index]) + "\n\n---\n\n"
        full_text = index_text + "\n".join(self.current_content)
        
        logger = get_logger()
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(full_text)
            logger.info(f"[Merged Created] {vol_filename} ({len(full_text)} chars)")
        except Exception as e:
            logger.error(f"Error writing volume {vol_filename}: {e}")

        # リセット
        self.current_vol += 1
        self.current_content = []
        self.current_char_count = 0
        self.file_index = []

    def finalize(self):
        """最後に残っているバッファを書き出す"""
        self._flush_volume()
