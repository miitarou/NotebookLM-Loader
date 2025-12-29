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
        
        行の途中で切断されないよう、行単位で処理を行う。
        1行を追加するとサイズオーバーになる場合は、先にボリュームを閉じてから追加する。
        これにより、行が途中で切れることを完全に防ぐ。
        """
        logger = get_logger()
        lines = content.split('\n')
        part_num = 1
        current_part_lines = []
        current_part_size = 0
        
        for line in lines:
            line_with_newline = line + '\n'
            line_len = len(line_with_newline)
            
            # Partヘッダーのサイズを計算
            part_header = f"\n\n# {filename} (Part {part_num})\n\n"
            header_len = len(part_header) if not current_part_lines else 0
            
            # この行を追加した場合の合計サイズ
            projected_total = self.current_char_count + header_len + current_part_size + line_len
            
            if projected_total > self.max_chars_per_volume and current_part_lines:
                # 現在のPartを確定して追加
                part_header = f"\n\n# {filename} (Part {part_num})\n\n"
                full_chunk = part_header + ''.join(current_part_lines)
                
                self.current_content.append(full_chunk)
                self.file_index.append(f"{filename} (Part {part_num})")
                self.current_char_count += len(full_chunk)
                
                if self.current_char_count >= self.max_chars_per_volume:
                    self._flush_volume()
                
                # 新しいPartを開始
                part_num += 1
                current_part_lines = [line_with_newline]
                current_part_size = line_len
                
                if part_num > MAX_PARTS:
                    logger.warning(f"Max parts ({MAX_PARTS}) reached for {filename}. File may be truncated.")
                    break
            else:
                # 行を現在のPartに追加
                current_part_lines.append(line_with_newline)
                current_part_size += line_len
        
        # 残りの行を追加
        if current_part_lines:
            part_header = f"\n\n# {filename} (Part {part_num})\n\n"
            
            if self.current_char_count + len(part_header) > self.max_chars_per_volume:
                self._flush_volume()
            
            full_chunk = part_header + ''.join(current_part_lines)
            
            self.current_content.append(full_chunk)
            self.file_index.append(f"{filename} (Part {part_num})")
            self.current_char_count += len(full_chunk)
            
            if self.current_char_count >= self.max_chars_per_volume:
                self._flush_volume()

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
