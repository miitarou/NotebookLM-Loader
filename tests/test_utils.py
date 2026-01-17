# tests/test_utils.py
"""utilsモジュールのユニットテスト"""

import pytest
from notebooklm_loader.utils import sanitize_content, sanitize_filename, get_output_filename, INVISIBLE_CHARS
from pathlib import Path


class TestSanitizeContent:
    """sanitize_content関数のテスト"""
    
    def test_removes_zero_width_space(self):
        """ゼロ幅スペース (U+200B) を除去できること"""
        input_text = "Hello\u200bWorld"
        result = sanitize_content(input_text)
        assert result == "HelloWorld"
        assert "\u200b" not in result
    
    def test_removes_bom(self):
        """BOM (U+FEFF) を除去できること"""
        input_text = "\ufeffHello World"
        result = sanitize_content(input_text)
        assert result == "Hello World"
        assert "\ufeff" not in result
    
    def test_removes_null_character(self):
        """NULL文字 (U+0000) を除去できること"""
        input_text = "Hello\u0000World"
        result = sanitize_content(input_text)
        assert result == "HelloWorld"
        assert "\u0000" not in result
    
    def test_removes_all_invisible_chars(self):
        """全ての不可視文字を除去できること"""
        # 全ての不可視文字を含むテスト文字列を作成
        input_text = "Start"
        for char in INVISIBLE_CHARS:
            input_text += char
        input_text += "End"
        
        result = sanitize_content(input_text)
        assert result == "StartEnd"
        
        # 全ての不可視文字が除去されていることを確認
        for char in INVISIBLE_CHARS:
            assert char not in result
    
    def test_preserves_normal_text(self):
        """通常テキストを変更しないこと"""
        input_text = "これは日本語テキストです。Hello World! 123"
        result = sanitize_content(input_text)
        assert result == input_text
    
    def test_preserves_newlines_and_tabs(self):
        """改行とタブを保持すること"""
        input_text = "Line1\nLine2\tTabbed"
        result = sanitize_content(input_text)
        assert result == input_text
        assert "\n" in result
        assert "\t" in result
    
    def test_empty_string(self):
        """空文字列を正しく処理すること"""
        result = sanitize_content("")
        assert result == ""
    
    def test_multiple_occurrences(self):
        """複数回出現する不可視文字を全て除去すること"""
        input_text = "\u200bA\u200bB\u200bC\u200b"
        result = sanitize_content(input_text)
        assert result == "ABC"
    
    def test_real_world_example(self):
        """実際に問題になったケース（ASUSTeK末尾のゼロ幅スペース）"""
        input_text = "ASUSTeK COMPUTER INC.\u200b"
        result = sanitize_content(input_text)
        assert result == "ASUSTeK COMPUTER INC."


class TestSanitizeFilename:
    """sanitize_filename関数のテスト"""
    
    def test_removes_invalid_characters(self):
        """無効な文字を除去すること"""
        input_name = 'file<name>with:invalid*chars?.txt'
        result = sanitize_filename(input_name)
        assert '<' not in result
        assert '>' not in result
        assert ':' not in result
        assert '*' not in result
        assert '?' not in result
    
    def test_preserves_valid_filename(self):
        """有効なファイル名を変更しないこと"""
        input_name = "valid_filename.txt"
        result = sanitize_filename(input_name)
        assert result == input_name
    
    def test_japanese_filename(self):
        """日本語ファイル名を保持すること"""
        input_name = "日本語ファイル名.txt"
        result = sanitize_filename(input_name)
        assert result == input_name


class TestGetOutputFilename:
    """get_output_filename関数のテスト"""
    
    def test_basic_path(self):
        """基本的なパス変換"""
        root = Path("/root/folder")
        file = Path("/root/folder/subfolder/file.docx")
        result = get_output_filename(root, file, ".md")
        assert result == "subfolder_file.md"
    
    def test_root_level_file(self):
        """ルートレベルのファイル"""
        root = Path("/root/folder")
        file = Path("/root/folder/file.docx")
        result = get_output_filename(root, file, ".md")
        assert result == "file.md"
    
    def test_pdf_extension(self):
        """PDF拡張子への変換"""
        root = Path("/root/folder")
        file = Path("/root/folder/document.pptx")
        result = get_output_filename(root, file, ".pdf")
        assert result.endswith(".pdf")
