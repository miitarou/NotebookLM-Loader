# NotebookLM Loader (Powered by MarkItDown)

[ **English** | [日本語](README.md) | [中文](README_CN.md) ]


This is a Python tool designed to convert Microsoft Office files (Word, Excel, PowerPoint) into a Markdown format optimized for **Google NotebookLM**.
It aims to **structure unstructured data** (tables, lists, headings) within documents to maximize the accuracy of RAG (Retrieval-Augmented Generation).
It utilizes **Microsoft's official `MarkItDown`** conversion engine for high-fidelity text extraction.

## Key Features

1.  **Smart Chunking (Merged Output)**:
    *   Automatically merges converted text files into larger **`Merged_Files_VolXX.md`** files (approx. 200,000 chars each).
    *   These merged files, along with auto-converted PDFs, are output to the **`converted_files_merged` folder**.
    *   Users just need to drag and drop the contents of this folder into NotebookLM.
    *   Recursive splitting ensures no single file exceeds the token limit.

2.  **Auto-Switch to PDF (High Density Visuals)**:
    *   If a file (like a PowerPoint slide deck) is determined to be "High Visual Density" (many images, little text), the tool **automatically converts it to PDF** using LibreOffice (instead of Markdown).
    *   This eliminates the manual effort of converting files to PDF specifically for NotebookLM registration.

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

- `--merge`: **(Recommended: Smart Mode)**
    - In addition to standard 1-to-1 conversion (`converted_files`), generating a **`converted_files_merged` folder**.
    - This folder contains optimized "Volume" files (under 200k chars) and auto-converted PDFs. Use this folder for uploading to NotebookLM.
- `--skip-ppt`:
    - **Excludes** PowerPoint (.pptx) files from the dataset entirely.
    - These files will not be converted to Markdown nor PDF. Use this only if you intentionally want to ignore PowerPoint files.

## Visual Density Report

The report displayed after execution shows whether each file was processed as "Text (Markdown)" or "Visual (PDF)".
Files marked as "High Visual Density" are **automatically exported as PDF**, so no further action is required. Simply upload them to NotebookLM.

MIT
