#!/usr/bin/env python3
"""
NotebookLM Loader - Office files to Markdown/PDF converter for NotebookLM

Usage:
    python office_to_notebooklm.py /path/to/folder [--merge] [--verbose] [--dry-run]
    
For more options, run:
    python office_to_notebooklm.py --help
"""

from notebooklm_loader import run

if __name__ == "__main__":
    exit(run())
