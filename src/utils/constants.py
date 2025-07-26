"""
共享常量模块 - 存放各个模块共用的正则表达式模式和常量
避免代码重复，保证一致性
"""
import re

# 正则表达式模式 - 预编译提高性能
class Patterns:
    """预编译的正则表达式模式集合"""
    
    # 标题相关模式
    HEADING_PATTERNS = [
        re.compile(r'^[一二三四五六七八九十]+、'),          # 一、二、三、
        re.compile(r'^（[一二三四五六七八九十]+）'),        # （一）（二）（三）
        re.compile(r'^[0-9]+\.'),                         # 1. 2. 3.
        re.compile(r'^[0-9]+、'),                         # 1、2、3、
    ]
    
    # 列表相关模式
    ORDERED_LIST_PATTERN = re.compile(r'^\d+\.\s*')          # 匹配有序列表 "1. "
    ORDERED_LIST_PATTERN_PREPROCESSOR = re.compile(r'^\d+\.\s+')  # 预处理器使用的版本（包含空格）
    ORDERED_LIST_DOT_REPLACE_PATTERN = re.compile(r'^(\d+)\.\s+')  # 替换点号模式
    
    NUMBERED_LIST_PATTERN = re.compile(r'^(\s*)(\d+)\.\s+')  # 数字列表模式（带缩进）
    UNORDERED_LIST_PATTERN = re.compile(r'^\s*\*\s+')        # 无序列表模式 "* "
    UNORDERED_LIST_REPLACE_PATTERN = re.compile(r'^(\s*)\*\s+')  # 无序列表替换模式
    INDENTED_LIST_PATTERN = re.compile(r'^\s+[-*]')          # 缩进列表模式
    
    # 图片相关模式
    MARKDOWN_IMAGE_PATTERN = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')  # ![alt](url)
    OBSIDIAN_IMAGE_PATTERN = re.compile(r'!\[\[([^\]]+)\]\]')         # ![[filename]]
    
    # 图片文件名模式
    IMAGE_FILENAME_PATTERNS = [
        re.compile(r'Pasted image \d+'),                              # Obsidian粘贴图片
        re.compile(r'^006Fd7o3gy1.*\.(png|jpg|jpeg|gif|bmp)$'),      # 微博图片
        re.compile(r'^Screenshot.*\.(png|jpg|jpeg|gif|bmp)$'),        # 截图文件
        re.compile(r'^.*\.(png|jpg|jpeg|gif|bmp)$'),                  # 通用图片文件
    ]
    
    # 文本处理模式
    BOLD_PATTERN = re.compile(r'\*\*(.*?)\*\*')              # **粗体**
    BOLD_UNDERSCORE_PATTERN = re.compile(r'__(.*?)__')       # __粗体__
    CHINESE_CHAR_PATTERN = re.compile(r'[\u4e00-\u9fff]')    # 中文字符
    
    # 数学公式模式
    LATEX_INLINE_MATH_PATTERN = re.compile(r'\$[^$]+\$')     # 行内LaTeX数学公式 $...$
    LATEX_BLOCK_MATH_PATTERN = re.compile(r'\$\$[^$]+\$\$')  # 块级LaTeX数学公式 $$...$$
    
    # 表格相关模式
    TABLE_ROW_PATTERN = re.compile(r'^\s*\|')                # 表格行 "|..."
    
    # 图表标题模式 - 匹配 图/图片/表/表格/图表 + 可选空格 + 数字 + 可选空格 + 标点(:：.) + 描述
    CAPTION_PATTERN = re.compile(r'^(图片?|表格?|图表)\s*(\d+)\s*[:：.]\s*(.*)$')
    CAPTION_PREFIX_PATTERN = re.compile(r'^(图片?|表格?|图表)\s*(\d+)\s*[:：.]\s*')  # 仅匹配前缀部分
    
    # 多级编号模式
    MULTI_LEVEL_NUMBER_PATTERN = re.compile(r'^(\d+\.\d+(?:\.\d+)*)\s+(.+)$')      # 2.1.1 内容
    SIMPLE_ORDERED_LIST_WITH_CONTENT = re.compile(r'^(\s*)(\d+)\.\s+(.+)$')        # 1. 内容（捕获内容）
    
    # 图片语法清理模式
    OBSIDIAN_IMAGE_CLEANUP_PATTERN = re.compile(r'!\[\[[^\]]+\]\]')                # 清理 ![[filename]]
    MARKDOWN_IMAGE_CLEANUP_PATTERN = re.compile(r'!\[[^\]]*\]\([^)]+\)')           # 清理 ![alt](path)
    
    # 通用文本清理模式
    WHITESPACE_CLEANUP_PATTERN = re.compile(r'\s+')                                # 清理多余空格
    PASTED_IMAGE_CLEANUP_PATTERN = re.compile(r'Pasted image \d{14}')              # 清理粘贴图片名称

# 文档格式常量
class DocumentFormats:
    """文档格式相关常量"""
    
    # 图片文件名模式（用于清理）
    IMAGE_CLEANUP_PATTERNS = [
        'Pasted image',  # Obsidian粘贴的图片
        '006Fd7o3gy1',   # 微博图片ID
        '.png',          # PNG文件扩展名
        '.jpg',          # JPG文件扩展名 
        '.jpeg',         # JPEG文件扩展名
        '.gif',          # GIF文件扩展名
        '.bmp',          # BMP文件扩展名
        '.webp'          # WEBP文件扩展名
    ]