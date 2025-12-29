# notebooklm_loader/summary.py
"""処理サマリーモジュール"""

import datetime
import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional, Dict, List, Any


@dataclass
class FileResult:
    """
    個別ファイルの処理結果
    
    Attributes:
        path: ファイルパス
        status: 処理ステータス（converted, skipped, error, password_protected）
        output: 出力ファイルパス
        error_message: エラーメッセージ
        file_type: ファイルタイプ
    """
    path: str
    status: str  # converted, skipped, error, password_protected
    output: Optional[str] = None
    error_message: Optional[str] = None
    file_type: Optional[str] = None


@dataclass
class ProcessingSummary:
    """
    処理サマリー
    
    Attributes:
        run_time: 実行時刻
        target_path: 処理対象パス
        total_files: 総ファイル数
        processed: 処理済み数
        skipped: スキップ数
        errors: エラー数
        password_protected: パスワード保護ファイル数
        files: ファイル別処理結果リスト
    """
    run_time: str = field(default_factory=lambda: datetime.datetime.now().isoformat())
    target_path: str = ""
    total_files: int = 0
    processed: int = 0
    skipped: int = 0
    errors: int = 0
    password_protected: int = 0
    files: List[Dict[str, Any]] = field(default_factory=list)
    
    def add_result(self, result: FileResult):
        """処理結果を追加"""
        self.files.append(asdict(result))
        self.total_files += 1
        
        if result.status == "converted":
            self.processed += 1
        elif result.status == "skipped":
            self.skipped += 1
        elif result.status == "error":
            self.errors += 1
        elif result.status == "password_protected":
            self.password_protected += 1
    
    def save(self, output_dir: Path) -> Path:
        """
        サマリーをJSONファイルに保存
        
        Args:
            output_dir: 出力ディレクトリ
            
        Returns:
            保存したファイルのパス
        """
        summary_file = output_dir / "processing_report.json"
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(asdict(self), f, ensure_ascii=False, indent=2)
        return summary_file
