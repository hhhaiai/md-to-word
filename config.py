from docx.shared import Mm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

class DocumentConfig:
    """公文格式配置类，基于GB/T 9704-2012标准"""
    
    # 页面设置 (单位: 毫米)
    PAGE_MARGINS = {
        'top': Mm(37),      # 天头
        'bottom': Mm(35),   # 地脚
        'left': Mm(28),     # 左边距
        'right': Mm(26)     # 右边距
    }
    
    # 字体配置
    FONTS = {
        'fangsong': 'FangSong',      # 正文字体
        'xiaobiaosong': '小标宋体',      # 标题字体
        'heiti': '黑体',               # 一级标题
        'kaiti': 'Kaiti TC'            # 二级标题
    }
    
    # 字号配置 (单位: 磅)
    FONT_SIZES = {
        'title': Pt(22),        # 二号 - 标题
        'body': Pt(16),         # 三号 - 正文
        'table': Pt(12),        # 四号 - 表格
        'page_num': Pt(14),     # 四号 - 页码
        'header': Pt(16)        # 三号 - 发文字号
    }
    
    # 行距配置
    LINE_SPACING = Pt(25.5)     # 固定值25.5磅
    
    # 版心配置
    CHARS_PER_LINE = 28         # 每行字符数
    LINES_PER_PAGE = 22         # 每页行数
    
    # 段落缩进
    FIRST_LINE_INDENT = Pt(32)  # 首行缩进2字符
    
    # 列表缩进配置
    LIST_INDENT = {
        'level1_left': Mm(12),       # 一级列表文字左缩进：增大数值→符号和文字整体右移
        'level1_first_line': -Mm(6), # 一级列表符号相对位置：减小数值→符号左移，增大符号与文字间距
        'level2_left': Mm(18),       # 二级列表文字左缩进：增大数值→符号和文字整体右移
        'level2_first_line': -Mm(6)  # 二级列表符号相对位置：减小数值→符号左移，增大符号与文字间距
    }
    
    # 对齐方式
    ALIGNMENTS = {
        'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
        'justify': WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        'left': WD_PARAGRAPH_ALIGNMENT.LEFT,
        'right': WD_PARAGRAPH_ALIGNMENT.RIGHT
    }
    
    # 表格自动适应配置
    TABLE_CONFIG = {
        'auto_fit': True,              # 启用表格自动适应
        'auto_fit_mode': 'window',     # 适应模式：'window'(窗口)/'contents'(内容)
        'preferred_width_percent': 100, # 首选宽度百分比
        'allow_row_breaks': True,      # 允许跨页断行
        'row_height_rule': 'auto'      # 行高规则：'auto'/'exact'/'at_least'
    }
    
    # Pandoc相关配置
    PANDOC_CONFIG = {
        # 数学公式处理方式
        'math_method': 'mathml',  # 使用MathML渲染数学公式
        
        # 表格处理
        'table_style': 'grid',  # 表格样式
        'table_caption': True,  # 是否显示表格标题
        
        # 列表处理
        'list_style': 'chinese',  # 中文列表样式
        
        # 图片处理
        'image_width': '100%',  # 图片宽度
        'image_dpi': 300,  # 图片DPI
        
        # 引用处理
        'citation_style': 'gb7714',  # 中文引用格式
        
        # 其他pandoc参数
        'extra_args': [
            '--preserve-tabs',
            '--wrap=none',
            '--reference-links'
        ],
        
        # 图片文字环绕设置
        'image_wrap_text': True,  # 是否启用图片文字环绕
        'image_wrap_type': 'topAndBottom'  # 环绕类型：topAndBottom, square, tight, through, none
    }
    
    # 图片路径配置
    IMAGE_CONFIG = {
        # Obsidian附件默认路径
        'obsidian_attachments_path': '/Users/yishentu/Library/Mobile Documents/iCloud~md~obsidian/Documents/YT\'s Obsidian/- Attachments',
        
        # 图片搜索路径列表（按优先级排序）
        'search_paths': [
            '/Users/yishentu/Library/Mobile Documents/iCloud~md~obsidian/Documents/YT\'s Obsidian/- Attachments',
            './images',
            './assets',
            './'
        ],
        
        # 支持的图片格式
        'supported_formats': ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'],
        
        # 是否复制图片到输出目录
        'copy_images': True,
        
        # 图片输出目录（相对于Word文档）
        'output_dir': 'images'
    }