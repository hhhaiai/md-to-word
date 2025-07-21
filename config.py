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
        'page_num': Pt(14),     # 四号 - 页码
        'header': Pt(16)        # 三号 - 发文字号
    }
    
    # 行距配置
    LINE_SPACING = Pt(28.8)     # 固定值28.8磅
    
    # 版心配置
    CHARS_PER_LINE = 28         # 每行字符数
    LINES_PER_PAGE = 22         # 每页行数
    
    # 段落缩进
    FIRST_LINE_INDENT = Pt(32)  # 首行缩进2字符
    
    # 对齐方式
    ALIGNMENTS = {
        'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
        'justify': WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        'left': WD_PARAGRAPH_ALIGNMENT.LEFT,
        'right': WD_PARAGRAPH_ALIGNMENT.RIGHT
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
        ]
    }