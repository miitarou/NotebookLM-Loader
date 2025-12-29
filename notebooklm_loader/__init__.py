# NotebookLM Loader
# Office files to Markdown/PDF converter for NotebookLM

__version__ = "2.0.0"

from .config import Config
from .logger import setup_logging
from .summary import ProcessingSummary, FileResult
from .merger import MergedOutputManager
from .state import ProcessingState
from .main import run

__all__ = [
    'Config',
    'setup_logging', 
    'ProcessingSummary',
    'FileResult',
    'MergedOutputManager',
    'ProcessingState',
    'run',
]
