"""
列表格式化器 - 负责有序和无序列表的格式化
"""
from docx import Document
from docx.shared import Pt
from docx.oxml.shared import qn
from .base_formatter import BaseFormatter
from ..utils.constants import Patterns
from ..utils.xpath_cache import OptimizedXMLProcessor


class ListFormatter(BaseFormatter):
    """列表格式化器 - 负责有序和无序列表的格式化"""
    
    def __init__(self, config=None):
        super().__init__(config)
        self.xml_processor = OptimizedXMLProcessor()
    
    def format_lists(self, doc: Document):
        """格式化列表 - 保留Word列表格式，只调整缩进和字体
        
        注意：有序列表已在预处理阶段转换为正文格式，这里只处理Word生成的列表
        """
        for paragraph in doc.paragraphs:
            # 跳过标题段落，避免将包含数字的标题误判为列表
            if paragraph.style.name.startswith('Heading'):
                continue
                
            # 只处理Word内置列表
            if self._is_word_list_item(paragraph):
                self._format_word_list_item(paragraph)
    
    
    def _is_word_list_item(self, paragraph) -> bool:
        """判断是否为Word内置列表项"""
        # 检查Compact样式
        if paragraph.style.name == 'Compact':
            return True
            
        # 检查是否有列表编号
        if hasattr(paragraph, '_element'):
            numPr = self.xml_processor.cache.find_first(paragraph._element, './/w:numPr')
            return numPr is not None
            
        return False
    
    def _format_word_list_item(self, paragraph):
        """格式化Word内置列表项"""
        is_sub_item = self._is_sub_list_item(paragraph)
        
        # 应用缩进和格式设置
        if is_sub_item:
            self._apply_sub_item_format(paragraph)
        else:
            self._apply_main_item_format(paragraph)
        
        # 设置字体格式
        for run in paragraph.runs:
            run.font.name = self.config.FONTS['fangsong']
            run.font.size = self.config.FONT_SIZES['body']
            self._set_chinese_font(run, self.config.FONTS['fangsong'])
    
    def _is_sub_list_item(self, paragraph) -> bool:
        """判断是否为二级列表项"""
        if hasattr(paragraph, '_element'):
            ilvl_elem = self.xml_processor.cache.find_first(paragraph._element, './/w:ilvl')
            if ilvl_elem is not None:
                val = ilvl_elem.get(qn('w:val'))
                return bool(val and int(val) > 0)
        return False
    
    def _apply_sub_item_format(self, paragraph):
        """应用二级列表格式"""
        paragraph_format = paragraph.paragraph_format
        paragraph_format.left_indent = self.config.LIST_INDENT['level2_left']
        paragraph_format.first_line_indent = self.config.LIST_INDENT['level2_first_line']
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        
        # 启用文档网格对齐
        self._enable_snap_to_grid(paragraph)
    
    def _apply_main_item_format(self, paragraph):
        """应用一级列表格式"""
        paragraph_format = paragraph.paragraph_format
        paragraph_format.left_indent = self.config.LIST_INDENT['level1_left']
        paragraph_format.first_line_indent = self.config.LIST_INDENT['level1_first_line']
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        
        # 启用文档网格对齐
        self._enable_snap_to_grid(paragraph)
