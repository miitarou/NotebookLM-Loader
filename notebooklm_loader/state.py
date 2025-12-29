# notebooklm_loader/state.py
"""差分処理用の状態管理モジュール"""

import json
import hashlib
from pathlib import Path
from typing import Dict, Optional, Any
from dataclasses import dataclass, asdict, field
from datetime import datetime


@dataclass
class FileState:
    """ファイルの状態情報"""
    hash: str
    mtime: float
    output: str
    processed_at: str
    file_type: str


@dataclass
class ProcessingState:
    """
    処理状態を管理するクラス
    
    差分処理のために、処理済みファイルのハッシュと更新日時を記録する。
    """
    version: str = "1.0"
    files: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    
    @classmethod
    def load(cls, state_file: Path) -> 'ProcessingState':
        """
        状態ファイルから読み込む
        
        Args:
            state_file: 状態ファイルのパス
            
        Returns:
            ProcessingState インスタンス
        """
        if not state_file.exists():
            return cls()
        
        try:
            with open(state_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            state = cls(version=data.get('version', '1.0'))
            state.files = data.get('files', {})
            return state
        except Exception:
            return cls()
    
    def save(self, state_file: Path):
        """
        状態ファイルに保存する
        
        Args:
            state_file: 状態ファイルのパス
        """
        data = {
            'version': self.version,
            'last_updated': datetime.now().isoformat(),
            'files': self.files
        }
        with open(state_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    def get_file_hash(self, file_path: Path) -> str:
        """
        ファイルのハッシュを計算する
        
        Args:
            file_path: 対象ファイルのパス
            
        Returns:
            SHA256ハッシュ文字列
        """
        sha256 = hashlib.sha256()
        try:
            with open(file_path, 'rb') as f:
                for chunk in iter(lambda: f.read(8192), b''):
                    sha256.update(chunk)
            return sha256.hexdigest()
        except Exception:
            return ""
    
    def needs_processing(self, file_path: Path, file_key: str) -> bool:
        """
        ファイルが処理を必要とするか判定する
        
        Args:
            file_path: 対象ファイルのパス
            file_key: 状態管理用のキー
            
        Returns:
            処理が必要な場合True
        """
        if file_key not in self.files:
            return True
        
        stored = self.files[file_key]
        
        # 更新日時チェック（高速）
        try:
            current_mtime = file_path.stat().st_mtime
            if current_mtime != stored.get('mtime'):
                # 更新日時が異なる場合、ハッシュも確認
                current_hash = self.get_file_hash(file_path)
                if current_hash != stored.get('hash'):
                    return True
        except Exception:
            return True
        
        return False
    
    def record_processed(self, file_path: Path, file_key: str, output: str, file_type: str):
        """
        処理完了を記録する
        
        Args:
            file_path: 処理したファイルのパス
            file_key: 状態管理用のキー
            output: 出力ファイル名
            file_type: ファイルタイプ
        """
        try:
            mtime = file_path.stat().st_mtime
            file_hash = self.get_file_hash(file_path)
            
            self.files[file_key] = {
                'hash': file_hash,
                'mtime': mtime,
                'output': output,
                'processed_at': datetime.now().isoformat(),
                'file_type': file_type
            }
        except Exception:
            pass
    
    def remove_deleted(self, current_files: set):
        """
        削除されたファイルの記録を削除する
        
        Args:
            current_files: 現在存在するファイルのキーセット
        """
        to_remove = [k for k in self.files if k not in current_files]
        for key in to_remove:
            del self.files[key]
