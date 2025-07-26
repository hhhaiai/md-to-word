# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Markdown到Word公文格式转换工具，符合GB/T 9704-2012《党政机关公文格式》国家标准。

**重要更新**：基于Pandoc引擎，支持LaTeX公式、表格和列表，采用模块化设计。

## Common Commands

```bash
# Run the converter
python3 md_to_word.py document.md
python3 md_to_word.py input.md -o output.docx

# Install dependencies  
pip3 install -r requirements.txt

# With custom Obsidian vault
OBSIDIAN_VAULT_NAME="我的笔记" python3 md_to_word.py document.md
```

## Architecture Overview

### Processing Pipeline
```
Input .md → MarkdownPreprocessor → PandocProcessor → WordPostprocessor → Output .docx
              ↓                        ↓                 ↓
        (cleaned text +          (raw DOCX)      (7 specialized formatters)
         metadata dict)
```

### Core Processing Flow

1. **MarkdownPreprocessor** (`src/core/markdown_preprocessor.py`):
   - Removes YAML frontmatter and metadata
   - **Critical**: Converts ordered lists `1. ` → `` `1.` `` to prevent Pandoc auto-numbering
   - Repositions captions after images/tables
   - Extracts filename as document title

2. **PandocProcessor** (`src/core/pandoc_processor.py`):
   - Executes Pandoc with `--mathml` for LaTeX formula support
   - Creates temporary files with proper cleanup
   - Returns path to generated DOCX

3. **WordPostprocessor** (`src/core/word_postprocessor.py`):
   - Orchestrates 7 formatters in specific order
   - Special handling for MathML formulas
   - Advanced image processing with Obsidian support

### Formatter Modules

Located in `src/formatters/`:
- **PageFormatter**: Document grid (22行×28字), page margins
- **ParagraphFormatter**: Font/size per GB/T 9704-2012, grid alignment
- **DocumentTitleFormatter**: Title from filename, attachment formatting
- **TableFormatter**: Table styling (仿宋、三号)
- **ListFormatter**: List indentation and formatting
- **ImageFormatter**: Full-width images (156mm), aspect ratio preservation

### Key Design Patterns

1. **"Controlled Conversion" Pattern**: Preprocessor neutralizes certain Markdown features (like ordered lists) to ensure post-processor has full control over formatting.

2. **Composition Pattern**: WordPostprocessor delegates to specialized formatters instead of implementing everything in one class.

3. **Math Formula Preservation**: Special detection and handling of MathML content throughout the pipeline.

### Important Data Structures

**Metadata Dictionary**:
```python
{
    'title': str,           # From filename
    'content': str,         # Preprocessed Markdown  
    'attachments': list     # Attachment descriptions
}
```

**Image Info Structure**:
```python
{
    'path': str,            # Actual file path
    'title': str,           # Alt text/caption
    'type': str,            # 'markdown' or 'obsidian'
    'original': str,        # Original syntax
    'paragraph': Paragraph, # Document paragraph
}
```

### Image Path Search Order
1. Source Markdown file directory
2. Obsidian attachments folder
3. `./images` directory
4. `./assets` directory
5. Current directory

### Security Considerations
- Path traversal protection in `PathValidator`
- Command injection prevention using `shlex.quote()`
- XML injection prevention in formatters

### Development Notes

- Python 3 + Pandoc required (system install)
- Main entry: `md_to_word.py`
- Config: `src/config/config.py`
- All formatters inherit from `BaseFormatter`
- XPath queries cached for performance
- Regex patterns precompiled in `constants.py`