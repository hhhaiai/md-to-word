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
        """格式化列表 - 保留Word列表格式，只调整缩进和字体"""
        for paragraph in doc.paragraphs:
            # 跳过标题段落，避免将包含数字的标题误判为列表
            if paragraph.style.name.startswith('Heading'):
                continue
                
            list_type = self._detect_list_type(paragraph)
            
            if list_type == 'word_list':
                self._format_word_list_item(paragraph)
            elif list_type == 'ordered_list':
                self._format_ordered_list_paragraph(paragraph)
    
    def _detect_list_type(self, paragraph):
        """检测段落的列表类型"""
        # 检查是否为Word内置列表（Compact样式或有列表编号）
        if self._is_word_list_item(paragraph):
            return 'word_list'
        
        return 'none'
    
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
        self._apply_list_font_format(paragraph)
    
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
    
    def _apply_list_font_format(self, paragraph):
        """应用列表字体格式"""
        for run in paragraph.runs:
            run.font.name = self.config.FONTS['fangsong']
            run.font.size = self.config.FONT_SIZES['body']
            self._set_chinese_font(run, self.config.FONTS['fangsong'])
    
    def _format_ordered_list_paragraph(self, paragraph):
        """格式化有序列表段落"""
        paragraph_format = paragraph.paragraph_format
        
        # 检查是否需要分割段落（包含多个列表项）
        text = paragraph.text.strip()
        items = Patterns.ORDERED_LIST_SPLIT_PATTERN.split(text)[1:]  # 分割并保留序号
        
        if len(items) > 2:  # 包含多个列表项
            self._split_multi_item_paragraph(paragraph, items)
        else:
            # 单个有序列表项，直接格式化
            self._format_single_ordered_item(paragraph)
    
    def _split_multi_item_paragraph(self, paragraph, items):
        """分割包含多个列表项的段落"""
        paired_items = self._organize_list_items(items)
        
        if not paired_items:
            return
            
        # 更新第一个段落为第一项
        self._update_first_paragraph(paragraph, paired_items[0])
        
        # 为剩余项目创建新段落
        self._create_additional_paragraphs(paragraph, paired_items[1:])
    
    def _organize_list_items(self, items):
        """将分割的列表项重新组织为完整的项目"""
        paired_items = []
        for i in range(0, len(items), 2):
            if i + 1 < len(items):
                paired_items.append(items[i] + items[i + 1])
        return paired_items
    
    def _update_first_paragraph(self, paragraph, first_item_text):
        """更新第一个段落的内容和格式"""
        paragraph.clear()
        first_run = paragraph.add_run(first_item_text.strip())
        self._apply_list_run_format(first_run)
        self._set_ordered_list_format(paragraph)
    
    def _create_additional_paragraphs(self, original_paragraph, remaining_items):
        """为剩余的列表项创建新段落"""
        current_p = original_paragraph._element
        parent = current_p.getparent()
        
        for item_text in remaining_items:
            # 创建新的XML段落元素
            new_p = original_paragraph._element.makeelement(
                original_paragraph._element.tag, 
                original_paragraph._element.attrib
            )
            parent.insert(parent.index(current_p) + 1, new_p)
            current_p = new_p
            
            # 创建新的段落对象并设置内容
            new_paragraph = self._create_paragraph_object(new_p, original_paragraph._parent)
            self._set_paragraph_content_and_format(new_paragraph, item_text.strip())
    
    def _create_paragraph_object(self, xml_element, parent):
        """创建新的段落对象"""
        from docx.text.paragraph import Paragraph
        return Paragraph(xml_element, parent)
    
    def _set_paragraph_content_and_format(self, paragraph, text):
        """设置段落内容和格式"""
        new_run = paragraph.add_run(text)
        self._apply_list_run_format(new_run)
        self._set_ordered_list_format(paragraph)
    
    def _apply_list_run_format(self, run):
        """应用列表运行的字体格式"""
        run.font.name = self.config.FONTS['fangsong']
        run.font.size = self.config.FONT_SIZES['body']
        self._set_chinese_font(run, self.config.FONTS['fangsong'])
    
    def _format_single_ordered_item(self, paragraph):
        """格式化单个有序列表项"""
        self._set_ordered_list_format(paragraph)
        
        # 确保列表项使用正确字体
        for run in paragraph.runs:
            self._apply_list_run_format(run)
    
    def _set_ordered_list_format(self, paragraph):
        """设置有序列表的段落格式"""
        paragraph_format = paragraph.paragraph_format
        paragraph_format.left_indent = Pt(0)
        paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        
        # 启用文档网格对齐
        self._enable_snap_to_grid(paragraph)
    def _enable_snap_to_grid(self, paragraph):
        """启用段落的文档网格对齐"""
        from docx.oxml.ns import qn as qn_func
        from docx.oxml.shared import OxmlElement
        
        pPr = paragraph._element.get_or_add_pPr()
        
        # 检查是否已有snapToGrid元素
        snapToGrid = pPr.find(qn_func("w:snapToGrid"))
        if snapToGrid is None:
            # 创建新的snapToGrid元素
            snapToGrid = OxmlElement("w:snapToGrid")
            pPr.append(snapToGrid)
