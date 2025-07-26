# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

dangerously-skip-permission: false

## Project Overview

This is a Markdown to Word converter that transforms Markdown files into Word documents compliant with GB/T 9704-2012 (Chinese government document standards). The converter uses a three-stage pipeline: MarkdownPreprocessor → PandocProcessor → WordPostprocessor.

## Build and Run Commands

```bash
# Install dependencies
pip3 install -r requirements.txt

# Basic usage
python3 md_to_word.py document.md

# Specify output file
python3 md_to_word.py input.md -o output.docx

# Use with Obsidian Vault environment variables
OBSIDIAN_VAULT_NAME="我的笔记" python3 md_to_word.py document.md
```

## Architecture Overview

### Processing Pipeline
1. **MarkdownPreprocessor** (`src/core/markdown_preprocessor.py`): Cleans Markdown content, extracts metadata, and neutralizes formatting for controlled conversion
2. **PandocProcessor** (`src/core/pandoc_processor.py`): Converts Markdown to DOCX using Pandoc with MathML support
3. **WordPostprocessor** (`src/core/word_postprocessor.py`): Applies GB/T 9704-2012 formatting using 7 specialized formatters

### Key Components
- **Formatters** (`src/formatters/`): 7 specialized modules for different document elements (page, paragraph, title, table, list, image, base)
- **Utils** (`src/utils/`): Support utilities including path validation, XPath caching, and security checks
- **Config** (`src/config/config.py`): Central configuration for fonts, margins, and document standards

### Critical Implementation Details

#### Ordered List Handling
The preprocessor converts ordered lists from `1. item` to `` `1.` item`` to prevent Pandoc auto-numbering and ensure consistent font formatting per GB/T 9704-2012.

#### Image Processing
- Supports both standard Markdown `![](path)` and Obsidian `![[filename]]` formats
- Searches multiple paths: relative to MD file, Obsidian attachments, ./images, ./assets, current directory
- Auto-adjusts images to full page width (156mm) while maintaining aspect ratio

#### Math Formula Preservation
- LaTeX formulas (`$...$` and `$$...$$`) are converted to MathML by Pandoc
- Special handling throughout pipeline to preserve MathML content in captions and paragraphs

## Development Notes

### Dependencies
- **Pandoc**: Required system dependency for Markdown conversion
- **python-docx**: Python library for Word document manipulation

### Security Considerations
- Path traversal protection in `path_validator.py`
- Safe subprocess execution with proper escaping
- XML injection prevention in formatters

### GB/T 9704-2012 Standards
- Page margins: Top 37mm, Bottom 35mm, Left 28mm, Right 26mm
- Document grid: 28 characters per line, 22 lines per page
- Fonts: 小标宋体 (title), 仿宋 (body), 黑体 (H2), 楷体 (H3)