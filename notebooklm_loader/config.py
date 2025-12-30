# notebooklm_loader/config.py
"""設定管理モジュール"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Set, Optional

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False


@dataclass
class Config:
    """
    アプリケーション設定
    
    Attributes:
        max_file_size_mb: スキップする最大ファイルサイズ（MB）
        merge_volume_mb: マージボリュームの最大サイズ（MB）
        visual_density_threshold: 視覚密度判定の閾値
        verbose: 詳細ログ出力
        quiet: コンソール出力抑制
        dry_run: 実行計画のみ表示
        merge: マージモード有効
        skip_ppt: PowerPointスキップ
    """
    # ファイル処理設定
    max_file_size_mb: int = 100
    merge_volume_mb: int = 35
    max_chars_per_volume: int = 5000000  # マージボリュームの最大文字数（デフォルト500万文字）
    visual_density_threshold: int = 300
    
    # CLI オプション
    verbose: bool = False
    quiet: bool = False
    dry_run: bool = False
    merge: bool = False
    skip_ppt: bool = False
    
    # 拡張子設定
    office_extensions_new: Set[str] = field(default_factory=lambda: {'.docx', '.xlsx', '.pptx', '.xls'})
    office_extensions_legacy: Set[str] = field(default_factory=lambda: {'.doc', '.ppt'})
    markitdown_extensions: Set[str] = field(default_factory=lambda: {'.rtf', '.epub', '.msg', '.eml'})
    visio_extensions: Set[str] = field(default_factory=lambda: {'.vsdx', '.vsd'})
    image_extensions: Set[str] = field(default_factory=lambda: {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp'})
    archive_extensions: Set[str] = field(default_factory=lambda: {'.zip', '.7z', '.rar', '.tar', '.gz', '.tgz', '.lzh'})
    skip_extensions: Set[str] = field(default_factory=lambda: {
        '.one', '.onetoc2', '.accdb', '.mdb',
        '.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv', '.webm',
        '.mp3', '.wav', '.aac', '.flac', '.ogg', '.wma', '.m4a',
        '.dwg', '.dxf', '.exe', '.dll', '.so', '.dylib',
        '.bin', '.dat', '.iso', '.img',
    })
    text_extensions: Set[str] = field(default_factory=lambda: {
        '.txt', '.md', '.py', '.js', '.jsx', '.ts', '.tsx', '.html', '.css', '.json',
        '.yaml', '.yml', '.org', '.sh', '.bat', '.zsh', '.rb', '.java', '.c', '.cpp',
        '.h', '.go', '.rs', '.php', '.pl', '.swift', '.kt', '.sql', '.xml', '.csv',
        '.log', '.ini', '.cfg', '.conf', '.properties', '.env', '.toml', '.tsv', '.rst'
    })
    
    @property
    def office_extensions_all(self) -> Set[str]:
        """全Office拡張子"""
        return self.office_extensions_new | self.office_extensions_legacy
    
    @property
    def max_file_size(self) -> int:
        """最大ファイルサイズ（バイト）"""
        return self.max_file_size_mb * 1024 * 1024
    
    @property
    def get_max_chars_per_volume(self) -> int:
        """マージボリュームの最大文字数（設定値を優先）"""
        # max_chars_per_volumeが明示的に設定されていればそれを使用
        return self.max_chars_per_volume
    
    @classmethod
    def from_yaml(cls, path: Path) -> 'Config':
        """YAMLファイルから設定を読み込む"""
        if not HAS_YAML:
            raise ImportError("PyYAML is required to load config from YAML. Install with: pip install pyyaml")
        with open(path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        
        config = cls()
        if 'processing' in data:
            proc = data['processing']
            if 'max_file_size_mb' in proc:
                config.max_file_size_mb = proc['max_file_size_mb']
            if 'merge_volume_mb' in proc:
                config.merge_volume_mb = proc['merge_volume_mb']
            if 'visual_density_threshold' in proc:
                config.visual_density_threshold = proc['visual_density_threshold']
            if 'max_chars_per_volume' in proc:
                config.max_chars_per_volume = proc['max_chars_per_volume']
        
        if 'skip_extensions' in data:
            config.skip_extensions = set(data['skip_extensions'])
        
        return config
    
    @classmethod
    def from_args(cls, args) -> 'Config':
        """
        argparse引数から設定を作成
        
        Args:
            args: argparseの結果オブジェクト
            
        Returns:
            Config インスタンス
        """
        config = cls(
            verbose=getattr(args, 'verbose', False),
            quiet=getattr(args, 'quiet', False),
            dry_run=getattr(args, 'dry_run', False),
            merge=getattr(args, 'merge', False),
            skip_ppt=getattr(args, 'skip_ppt', False),
        )
        
        # --configオプションで設定ファイルが指定された場合
        config_path = getattr(args, 'config', None)
        if config_path:
            from pathlib import Path
            yaml_config = cls.from_yaml(Path(config_path))
            # YAMLの設定をマージ
            config.max_file_size_mb = yaml_config.max_file_size_mb
            config.merge_volume_mb = yaml_config.merge_volume_mb
            config.visual_density_threshold = yaml_config.visual_density_threshold
            config.skip_extensions = yaml_config.skip_extensions
        
        return config

