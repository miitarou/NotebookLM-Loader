# NotebookLM Loader (Powered by MarkItDown)

[ [English](README_EN.md) | [日本語](README.md) | **中文** ]

这是一个 Python 工具，旨在将 Microsoft Office 文件（Word、Excel、PowerPoint）转换为针对 **Google NotebookLM** 优化的 Markdown 格式。

它采用 **Microsoft 官方的 `MarkItDown`** 转换引擎，实现高保真的文本提取，有效处理表格和列表。
关键在于，它具备 “智能分块 (Smart Chunking)” 和 “自动转换为 PDF (Auto-Switch to PDF)” 功能，能够处理包含大量文本和视觉内容的现实文档集。

## 主要功能

1.  **智能分块 (Smart Chunking - Merged Output)**:
    *   自动将转换后的文本文件合并为较大的 “卷 (Volume)” 文件（每个约 200,000 字符）。
    *   这大大减少了上传到 NotebookLM 所需的文件数量（例如，将 1,000 个小文档合并为 5 个大文件）。
    *   递归分割确保没有单个文件超过 Token 限制。

2.  **自动转换为 PDF (Auto-Switch to PDF - High Density Visuals)**:
    *   如果文件（如 PowerPoint 演示文稿）被判定为 “高视觉密度 (High Visual Density)”（图片多，文本少），该工具会使用 LibreOffice **自动将其转换为 PDF**（而不是 Markdown）。
    *   这确保 NotebookLM 能够 “看到” 图表和图解，而不仅仅是接收无意义的文本片段。

3.  **多合一加载器 (All-in-One Loader)**:
    *   递归扫描文件夹和 **ZIP 文件**。
    *   支持混合处理 Office 文档 (`.docx`, `.xlsx`, `.pptx`)、PDF 以及源代码/文本文件 (`.py`, `.txt`, `.md` 等)。

4.  **通用处理**:
    *   处理编码问题（ZIP 中的日语/中文文件名）。
    *   记录跳过的二进制文件。

## 环境要求

- Python 3.10+
- **LibreOffice** (用于将 PPTX/DOCX 转换为 PDF)
    - Mac: 检查 `/Applications/LibreOffice.app`
    - Linux: `soffice` 命令必须在 PATH 中

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

- `--merge`: （推荐）启用智能分块和自动 PDF 转换。生成 `converted_files_merged` 文件夹。
- `--skip-ppt`: 完全跳过 PowerPoint 文件。

## 许可证

MIT
