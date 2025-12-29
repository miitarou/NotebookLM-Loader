# notebooklm_loader/merger.py
"""Smart Chunking & Merged Outputモジュール"""

from pathlib import Path
from typing import List


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
        remaining = content
        part_num = 1
        
        while remaining:
            part_header = f"\n\n# {filename} (Part {part_num})\n\n"
            header_len = len(part_header)
            
            if self.current_char_count + header_len > self.max_chars_per_volume:
                 self._flush_volume()
            
            available_space = self.max_chars_per_volume - self.current_char_count - header_len
            
            if len(remaining) > available_space:
                # 行単位で分割：available_space以内で最後の改行位置を探す
                split_pos = remaining.rfind('\n', 0, available_space)
                
                if split_pos == -1:
                    # 改行が見つからない場合（非常に長い1行）
                    # スペースで分割を試みる
                    split_pos = remaining.rfind(' ', 0, available_space)
                
                if split_pos == -1:
                    # スペースも見つからない場合は、そのまま分割（最終手段）
                    split_pos = available_space
                else:
                    # 改行またはスペースの次の文字から新しいチャンクを開始
                    split_pos += 1
                
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

    def _flush_volume(self):
        """現在のバッファをファイルに書き出す"""
        if not self.current_content:
            return

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

        # リセット
        self.current_vol += 1
        self.current_content = []
        self.current_char_count = 0
        self.file_index = []

    def finalize(self):
        """最後に残っているバッファを書き出す"""
        self._flush_volume()
