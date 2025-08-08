# Markdown to Word Converter - Core Logic Documentation

## Overview

This project converts Markdown files to Word documents that strictly follow the GB/T 9704-2012 Chinese government document standards. The system uses a three-stage processing pipeline with modular formatters to achieve precise control over the final document formatting.

## Architecture

### Processing Pipeline
```
Input .md → MarkdownPreprocessor → PandocProcessor → WordPostprocessor → Output .docx
              ↓                        ↓                 ↓
        (cleaned text +          (raw DOCX)      (7 specialized formatters)
         metadata dict)
```

### 4-Layer Architecture
1. **Core Layer** (`src/core/`): Main processing controllers
2. **Formatters Layer** (`src/formatters/`): 7 specialized formatting modules  
3. **Utils Layer** (`src/utils/`): Support utilities and optimizations
4. **Config Layer** (`src/config/`): Configuration management

## Core Processors Deep Dive

### 1. MarkdownPreprocessor (`src/core/markdown_preprocessor.py`)

**Purpose**: Cleans and prepares Markdown content for Pandoc conversion while preserving critical formatting information.

#### Key Responsibilities:
- **Content Cleaning**: Removes YAML frontmatter and unwanted metadata
- **Format Neutralization**: Strategically disables certain Markdown features to give post-processing full control
- **Metadata Extraction**: Collects document title and attachment information
- **Content Restructuring**: Repositions captions and merges broken lines

#### Core Functions:

**`preprocess_file(file_path: str) -> Dict[str, Any]`**
- Main entry point for file processing
- Returns structured metadata: `{'title': str, 'content': str, 'attachments': list}`
- Extracts filename as document title (removing extension)

**`preprocess_content(content: str, file_path: str = '') -> str`**  
- Applies 8 sequential preprocessing filters:
  1. `_filter_yaml_frontmatter()` - Removes YAML front matter blocks
  2. `_filter_ending_metadata()` - Strips ending metadata (Date, tags)
  3. `_remove_bold_formatting()` - Converts `**text**` to plain text
  4. `_reposition_captions()` - Moves image/table captions after their elements
  5. `_fix_unordered_list_asterisks()` - Converts `* item` to `- item`
  6. `_merge_broken_lines()` - Intelligently merges accidentally split lines
  7. `_skip_first_level_headers()` - Removes/adjusts header levels
  8. `_convert_ordered_lists_to_text()` - **Critical**: Converts `1. item` to `` `1.` item``

#### Advanced Features:

**Dynamic Header Level Detection**
```python
def _skip_first_level_headers(self, lines: List[str]) -> List[str]:
    # If multiple H1 headers exist, downshift all header levels
    # If single/no H1, skip it (use filename as document title)
    h1_count = sum(1 for line in lines if line.strip().startswith('# ') and not line.strip().startswith('##'))
    
    if h1_count > 1:
        return self._adjust_header_levels(lines)  # # → ##, ## → ###, ### → text
    else:
        return [line for line in lines if not (line.strip().startswith('# ') and not line.strip().startswith('##'))]
```

**Ordered List Neutralization (Critical for GB/T 9704-2012 Compliance)**
```python
def _convert_ordered_lists_to_text(self, lines: List[str]) -> List[str]:
    # Converts: "1. 项目内容" → "`1.` 项目内容"  
    # Converts: "2.1.1 详细说明" → "`2.1.1` 详细说明"
    # This prevents Pandoc from auto-numbering and ensures consistent fonts
```

**Caption Repositioning System**
- Searches up to 10 lines before and 20 lines after for matching elements
- Handles complex cases like tables spanning multiple lines
- Preserves math formulas in captions during repositioning

### 2. PandocProcessor (`src/core/pandoc_processor.py`)

**Purpose**: Handles the actual Markdown to DOCX conversion using Pandoc as the conversion engine.

#### Key Responsibilities:
- **Pandoc Integration**: Safely executes Pandoc subprocess with proper error handling
- **Math Formula Processing**: Enables MathML rendering for LaTeX formulas
- **Temporary File Management**: Creates and cleans up temporary files
- **Document Loading**: Prepares DOCX for post-processing

#### Core Functions:

**`convert_markdown_to_docx(markdown_content: str, output_path: str, title: str = None) -> str`**
- Creates temporary Markdown file with optional title prefix
- Executes Pandoc with optimized arguments
- Returns path to generated DOCX file

**Pandoc Configuration**:
```python
def _get_pandoc_args(self) -> list:
    return [
        '--mathml',           # Enable MathML for LaTeX formulas
        # extra args are defined in config and appended
    ]
```

**Security & Error Handling**:
- Uses subprocess with capture_output=True for safe execution
- Comprehensive exception handling with custom `PandocError`
- Automatic cleanup of temporary files via `_cleanup_temp_files()`

<!-- load_docx_for_postprocessing was removed from code-path usage; retained here historically. -->

### 3. WordPostprocessor (`src/core/word_postprocessor.py`)

**Purpose**: Orchestrates all formatting operations to transform raw Pandoc output into GB/T 9704-2012 compliant documents.

#### Key Responsibilities:
- **Formatter Orchestration**: Coordinates 7 specialized formatters
- **Image Processing**: Advanced image insertion with Obsidian support
- **Math Formula Preservation**: Special handling for MathML content
- **Final Formatting**: Applies all document standards

#### Architecture Pattern:
Uses the **Composition Pattern** - delegates specific formatting tasks to specialized formatter classes instead of implementing everything in one large class.

#### Core Functions:

**`apply_formatting(docx_path: str, metadata: Dict[str, Any], original_markdown: str = None) -> str`**
- Main orchestration method
- Applies formatters in specific order:
  1. `PageFormatter` - Document grid and page setup
  2. `ParagraphFormatter` - Text and heading formatting  
  3. `DocumentTitleFormatter` - Title and attachments
  4. `TableFormatter` - Table styling
  5. `ListFormatter` - List formatting
  6. `ImageFormatter` - Image processing

**Advanced Image Processing System**:
```python
def process_and_insert_images(self):
    # 3-stage process:
    # Stage 1: Identify all images and captions
    # Stage 2: Replace image syntax with actual images
    # Stage 3: Format all captions with proper styling
```

**Math Formula Safe Processing**:
```python
def _has_math_formula(self, paragraph) -> bool:
    # Detects both MathML (converted) and LaTeX (original) formulas
    xml_str = paragraph._element.xml
    if 'oMath' in xml_str or 'oMathPara' in xml_str:
        return True
    # Also checks for LaTeX patterns: $...$ and $$...$$
    return Patterns.LATEX_INLINE_MATH_PATTERN.search(paragraph.text)
```

**Special Handling for Math Captions**:
```python
def _insert_image_before_math_caption(self, caption_paragraph, image_info):
    # When caption contains math formulas:
    # 1. Create new paragraph above for image
    # 2. Preserve original paragraph with math content intact
    # 3. Remove only image syntax, keep MathML elements
```

#### Formatter Integration:

**Initialization**:
```python
def __init__(self):
    self.config = DocumentConfig()
    
    # Initialize specialized formatters
    self.page_formatter = PageFormatter(self.config)
    self.paragraph_formatter = ParagraphFormatter(self.config)
    self.title_formatter = DocumentTitleFormatter(self.config)
    self.table_formatter = TableFormatter(self.config)
    self.list_formatter = ListFormatter(self.config)
    self.image_formatter = ImageFormatter(self.config)
```

## Data Flow and Contracts

### Data Structures

**Metadata Dictionary** (passed between processors):
```python
{
    'title': str,           # Document title from filename
    'content': str,         # Preprocessed Markdown content  
    'attachments': list     # List of attachment descriptions
}
```

**Image Information Structure**:
```python
{
    'path': str,                    # Actual image file path
    'title': str,                   # Image alt text or caption
    'type': str,                    # 'markdown' or 'obsidian'
    'original': str,                # Original syntax text
    'number': int,                  # Image counter
    'paragraph': Paragraph,         # Document paragraph object
    'caption_in_same_para': str,    # Caption in same paragraph
    'caption_paragraph': Paragraph  # Separate caption paragraph
}
```

### Processing Flow

1. **Input Phase**: 
   - File reading and initial parsing
   - Filename extraction for title

2. **Preprocessing Phase**:
   - Content cleaning and format neutralization
   - Metadata extraction
   - Strategic list and header processing

3. **Conversion Phase**:
   - Pandoc subprocess execution
   - MathML formula rendering
   - Basic DOCX structure creation

4. **Post-processing Phase**:
   - Document grid application (22×28 character grid)
   - Typography application (GB/T 9704-2012 fonts)
   - Image insertion and formatting
   - Caption processing and positioning
   - Final document validation

## Key Design Decisions

### 1. "Controlled Conversion" Pattern
The system intentionally disables certain Markdown features during preprocessing (like ordered lists) to ensure the post-processing stage has complete control over final formatting. This guarantees compliance with strict government document standards.

### 2. Modular Formatter Architecture  
Instead of a monolithic formatter class, the system uses 7 specialized formatters:
- **PageFormatter**: Page layout and grid
- **ParagraphFormatter**: Text and headings
- **DocumentTitleFormatter**: Titles and attachments
- **TableFormatter**: Table styling
- **ListFormatter**: List formatting
- **ImageFormatter**: Image processing
- **BaseFormatter**: Common functionality

### 3. Math Formula Preservation
Special handling throughout the pipeline to preserve LaTeX formulas:
- Preprocessor skips math blocks during line merging
- Pandoc renders to MathML format
- Post-processor detects and safely handles MathML content

### 4. Obsidian Integration
Full support for Obsidian vault integration:
- `![[image.png]]` syntax recognition (post-processor reinserts for this syntax)
- Standard Markdown `![alt](path)` images are handled directly by Pandoc
- Multi-path image search (vault, attachments, relative paths)
- Intelligent image title extraction

## Configuration System

### Document Standards (GB/T 9704-2012)
```python
# Page Layout
PAGE_MARGINS = {
    'top': Mm(37),     # 上边距37mm
    'bottom': Mm(35),  # 下边距35mm  
    'left': Mm(28),    # 左边距28mm
    'right': Mm(26)    # 右边距26mm
}

# Typography
FONTS = {
    'xiaobiaosong': '小标宋体',  # Document title
    'fangsong': '仿宋',          # Body text
    'heiti': '黑体',             # Level 1 headings
    'kaiti': '楷体'              # Level 2 headings
}

# Document Grid
CHARS_PER_LINE = 28    # 每行28字
LINES_PER_PAGE = 22    # 每页22行
```

### Performance Optimizations
- **XPath Caching**: Query results cached to avoid repeated XML parsing
- **Regex Precompilation**: Patterns compiled once in `constants.py`
- **Batch Processing**: Single-pass document traversal where possible

## Error Handling

### Exception Hierarchy
```python
FileProcessingError   # File I/O issues
PandocError           # Pandoc conversion failures
PathSecurityError     # Path traversal attempts
```

### Safety Mechanisms
- Path traversal protection in image processing
- XML injection prevention in formatters
- Safe subprocess execution for Pandoc calls
- Comprehensive temporary file cleanup

## Extension Points

### Adding New Preprocessing Rules
```python
def preprocess_content(self, content: str, file_path: str = '') -> str:
    lines = content.split('\n')
    
    # Add custom filter in the processing chain
    lines = self._my_custom_filter(lines)
    
    return '\n'.join(lines)
```

### Creating New Formatters
```python
from .base_formatter import BaseFormatter

class MyFormatter(BaseFormatter):
    def format_elements(self, doc: Document):
        for paragraph in doc.paragraphs:
            self._apply_custom_formatting(paragraph)
```

This documentation provides a comprehensive understanding of the core logic and architecture, enabling developers to maintain, extend, and troubleshoot the system effectively.