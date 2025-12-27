# Office to NotebookLM Converter (Powered by MarkItDown)

This is a Python tool designed to convert Microsoft Office files (Word, Excel, PowerPoint) into a Markdown format optimized for **Google NotebookLM**.

It utilizes **Microsoft's official `MarkItDown`** conversion engine for high-fidelity text extraction, effectively handling tables and lists.
Crucially, it features "Smart Chunking" and "Auto-Switch to PDF" to handle real-world document sets that contain both text-heavy and visual-heavy files.

## Key Features

1.  **Smart Chunking (Merged Output)**:
    *   Automatically merges converted text files into larger "Volume" files (approx. 200,000 chars each).
    *   This drastically reduces the number of file uploads required for NotebookLM (e.g., turning 1,000 small docs into 5 large ones).
    *   Recursive splitting ensures no single file exceeds the token limit.

2.  **Auto-Switch to PDF (High Density Visuals)**:
    *   If a file (like a PowerPoint slide deck) is determined to be "High Visual Density" (many images, little text), the tool **automatically converts it to PDF** using LibreOffice (instead of Markdown).
    *   This ensures NotebookLM can "see" the charts and diagrams, rather than just receiving meaningless text fragments.

3.  **All-in-One Loader**:
    *   Recursively scans folders and **ZIP files**.
    *   Supports Office docs (`.docx`, `.xlsx`, `.pptx`), PDFs, and source code/text files (`.py`, `.txt`, `.md`, etc.) mixed together.

4.  **Universal Processing**:
    *   Handles encoding issues (Japanese filenames in ZIPs).
    *   Logs skipped binary files.

## Requirements

- Python 3.10+
- **LibreOffice** (Required for PDF conversion of PPTX/DOCX)
    - Mac: `/Applications/LibreOffice.app` check
    - Linux: `soffice` command must be in PATH

## Installation

1.  Clone this repository.
2.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

Specify the target **input folder** or **ZIP file**.

### 1. Basic Usage (Recommended)
This generates individual Markdown files in `converted_files` AND merged "Volume" files in `converted_files_merged`.

```bash
# Process a folder
python office_to_notebooklm.py /path/to/documents --merge

# Process a ZIP file
python office_to_notebooklm.py /path/to/archive.zip --merge
```

The output in `converted_files_merged` will contain:
- `Merged_Files_Vol01.md`: Merged text content.
- `Presentation.pdf`: Visual-heavy files (auto-converted).
- `Manual.pdf`: Original PDF files (passed through).

**Just drag and drop the contents of `converted_files_merged` into NotebookLM.**

### Options

- `--merge`: (Recommended) Enables Smart Chunking and Auto-Switch to PDF. Generates `converted_files_merged` folder.
- `--skip-ppt`: Skips PowerPoint files entirely.

## License

MIT
