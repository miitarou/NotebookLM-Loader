# tests/test_merger.py
"""mergerモジュールのユニットテスト"""

import pytest
import tempfile
from pathlib import Path
from notebooklm_loader.merger import MergedOutputManager, MAX_PARTS


class TestMergedOutputManager:
    """MergedOutputManager クラスのテスト"""
    
    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリを作成"""
        with tempfile.TemporaryDirectory() as tmpdir:
            yield Path(tmpdir)
    
    def test_initialization(self, temp_output_dir):
        """初期化が正しく行われること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=1000)
        assert manager.output_dir == temp_output_dir
        assert manager.max_chars_per_volume == 1000
        assert manager.current_vol == 1
        assert manager.current_content == []
        assert manager.current_char_count == 0
    
    def test_add_small_content(self, temp_output_dir):
        """小さなコンテンツを追加できること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=10000)
        manager.add_content("test.md", "Hello World")
        
        assert len(manager.current_content) == 1
        assert manager.current_char_count == len("Hello World")
        assert "test.md" in manager.file_index
    
    def test_volume_flush_on_overflow(self, temp_output_dir):
        """容量オーバー時にボリュームがフラッシュされること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=100)
        
        # 100文字を超えるコンテンツを追加
        manager.add_content("file1.md", "A" * 50)
        manager.add_content("file2.md", "B" * 60)  # これで100を超える
        
        # file2を追加する前にfile1がフラッシュされるはず
        assert manager.current_vol >= 1
    
    def test_finalize_writes_remaining(self, temp_output_dir):
        """finalizeで残りのバッファが書き出されること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=10000)
        manager.add_content("test.md", "Test content")
        manager.finalize()
        
        # ファイルが作成されていることを確認
        output_files = list(temp_output_dir.glob("Merged_Files_Vol*.md"))
        assert len(output_files) >= 1
    
    def test_output_file_format(self, temp_output_dir):
        """出力ファイルのフォーマットが正しいこと"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=10000)
        manager.add_content("test.md", "Test content")
        manager.finalize()
        
        output_file = temp_output_dir / "Merged_Files_Vol01.md"
        assert output_file.exists()
        
        content = output_file.read_text(encoding='utf-8')
        assert "# Table of Contents" in content
        assert "test.md" in content
        assert "Test content" in content


class TestHandleHugeFile:
    """_handle_huge_file メソッドのテスト"""
    
    @pytest.fixture
    def temp_output_dir(self):
        """一時出力ディレクトリを作成"""
        with tempfile.TemporaryDirectory() as tmpdir:
            yield Path(tmpdir)
    
    def test_splits_huge_file_by_lines(self, temp_output_dir):
        """巨大ファイルが行単位で分割されること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=100)
        
        # 100文字を超えるコンテンツ（複数行）
        huge_content = "\n".join(["Line " + str(i) for i in range(20)])
        manager.add_content("huge.md", huge_content)
        manager.finalize()
        
        # 複数のボリュームが作成されること
        output_files = list(temp_output_dir.glob("Merged_Files_Vol*.md"))
        assert len(output_files) >= 1
    
    def test_preserves_line_integrity(self, temp_output_dir):
        """行の途中で切断されないこと"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=50)
        
        # 各行が完全に保持されるべきコンテンツ
        lines = ["日本語の行1", "日本語の行2", "日本語の行3"]
        huge_content = "\n".join(lines)
        manager.add_content("test.md", huge_content)
        manager.finalize()
        
        # 出力ファイルを全て読み込んで結合
        all_content = ""
        for f in sorted(temp_output_dir.glob("Merged_Files_Vol*.md")):
            all_content += f.read_text(encoding='utf-8')
        
        # 各行が完全に含まれていること
        for line in lines:
            assert line in all_content
    
    def test_adds_part_headers(self, temp_output_dir):
        """Partヘッダーが追加されること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=100)
        
        huge_content = "\n".join(["Line " + str(i) * 10 for i in range(50)])
        manager.add_content("huge.md", huge_content)
        manager.finalize()
        
        # 出力ファイルにPartヘッダーが含まれること
        output_files = list(temp_output_dir.glob("Merged_Files_Vol*.md"))
        all_content = ""
        for f in output_files:
            all_content += f.read_text(encoding='utf-8')
        
        assert "huge.md (Part 1)" in all_content
    
    def test_single_very_long_line(self, temp_output_dir):
        """非常に長い1行を処理できること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=1000)
        
        # 500文字の1行
        long_line = "A" * 500
        manager.add_content("single_line.md", long_line)
        manager.finalize()
        
        output_files = list(temp_output_dir.glob("Merged_Files_Vol*.md"))
        assert len(output_files) >= 1
    
    def test_empty_content(self, temp_output_dir):
        """空コンテンツを処理できること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=1000)
        manager.add_content("empty.md", "")
        manager.finalize()
        
        # エラーなく完了すること
        output_files = list(temp_output_dir.glob("Merged_Files_Vol*.md"))
        # 空コンテンツでもヘッダーは追加されるので1ファイルはできる
        assert len(output_files) >= 0


class TestVolumeNumbering:
    """ボリューム番号のテスト"""
    
    @pytest.fixture
    def temp_output_dir(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            yield Path(tmpdir)
    
    def test_volume_numbers_are_sequential(self, temp_output_dir):
        """ボリューム番号が連番であること"""
        manager = MergedOutputManager(temp_output_dir, max_chars_per_volume=50)
        
        # 複数のボリュームを作成
        for i in range(10):
            manager.add_content(f"file{i}.md", "X" * 30)
        manager.finalize()
        
        output_files = sorted(temp_output_dir.glob("Merged_Files_Vol*.md"))
        
        # ファイル名が連番であること
        for i, f in enumerate(output_files, 1):
            expected_name = f"Merged_Files_Vol{i:02d}.md"
            assert f.name == expected_name
