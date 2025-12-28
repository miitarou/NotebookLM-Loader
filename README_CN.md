# NotebookLM Loader (Powered by MarkItDown)

[ [English](README_EN.md) | [日本語](README.md) | **中文** ]

这是一个 Python 工具，旨在将 Microsoft Office 文件（Word、Excel、PowerPoint）转换为针对 **Google NotebookLM** 优化的 Markdown 格式。
其目的是对文档中的非结构化数据（表格、列表、标题）进行**明确的结构化 (Structuring)**，从而最大化 RAG（检索增强生成）的准确性。
它采用 **Microsoft 官方的 `MarkItDown`** 转换引擎，实现高保真的文本提取。

## 主要功能

1.  **智能分块 (Smart Chunking - Merged Output)**:
    *   自动将转换后的文本文件合并为较大的 **`Merged_Files_VolXX.md`** 文件（每个约 40MB / 1200万字符）。
    *   这些合并文件以及自动转换的 PDF 都将输出到 **`converted_files_merged` 文件夹** 中。
    *   用户只需将该文件夹的内容拖放到 NotebookLM 中即可。
    *   递归分割确保没有单个文件超过上传限制。

2.  **自动转换为 PDF (Auto-Switch to PDF - High Density Visuals)**:
    *   如果文件（如 PowerPoint 演示文稿）被判定为 “高视觉密度 (High Visual Density)”（图片多，文本少），该工具会使用 LibreOffice **自动将其转换为 PDF**（而不是 Markdown）。
    *   这消除了专门为 NotebookLM 注册而手动将文件转换为 PDF 的工作。

3.  **多合一加载器 (All-in-One Loader)**:
    *   递归扫描文件夹和 **ZIP 文件**。
    *   支持混合处理 Office 文档 (`.docx`, `.xlsx`, `.pptx`)、PDF 以及源代码/文本文件 (`.py`, `.txt`, `.md` 等)。

4.  **通用处理**:
    *   处理编码问题（ZIP 中的日语/中文文件名）。
    *   记录跳过的二进制文件。

## 环境要求

- Python 3.10+
- **LibreOffice** (自动 PDF 转换功能必须)
    - 推荐 **7.0 或更高版本**
    - Mac: 必须安装在 `/Applications/LibreOffice.app`
    - Linux/Windows: `soffice` 命令必须在 PATH 中

## 安装

1.  克隆此仓库。
2.  安装依赖:
    ```bash
    pip install -r requirements.txt
    ```

## 用法

指定目标 **输入文件夹** 或 **ZIP 文件**。

### 1. 基本用法（推荐）
这将在 `converted_files` 中生成单独的 Markdown 文件，并在 `converted_files_merged` 中生成合并的 “卷 (Volume)” 文件。

```bash
# 处理文件夹
python office_to_notebooklm.py /path/to/documents --merge

# 处理 ZIP 文件
python office_to_notebooklm.py /path/to/archive.zip --merge
```

`converted_files_merged` 中的输出将包含：
- `Merged_Files_Vol01.md`: 合并的文本内容。
- `Presentation.pdf`: 视觉密集型文件（自动转换）。
- `Manual.pdf`: 原始 PDF 文件（直接通过）。

**只需将 `converted_files_merged` 的内容拖放到 NotebookLM 中即可。**

### 选项

- `--merge`: **（推荐：智能模式）**
    - 除了普通模式（`converted_files` 中的一对一输出）外，还会生成 **`converted_files_merged` 文件夹**。
    - 该文件夹包含针对文件大小限制（40MB）优化的合并文件和自动转换的 PDF 文件。上传到 NotebookLM 时请使用此文件夹。
- `--skip-ppt`:
    - 将 PowerPoint (.pptx) 文件**从数据集中完全排除**。
    - 指定此选项后，将不会执行 Markdown 转换或 PDF 转换。仅在您有意忽略 PowerPoint 文件时使用此选项。

## 视觉密度报告 (Visual Density Report)

执行后显示的报告是信息性的，指出每个文件是作为 “文本 (Markdown)” 还是 “视觉 (PDF)” 处理的。
标记为 “High Visual Density” 的文件**已自动导出为 PDF**，因此无需用户进行额外操作。直接上传到 NotebookLM 即可。

MIT
