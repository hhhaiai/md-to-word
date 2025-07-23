"""
页面格式化器 - 负责页面设置、页眉页脚、页码等
"""
from docx import Document
from docx.shared import Mm
from docx.oxml.shared import OxmlElement, qn
from .base_formatter import BaseFormatter


class PageFormatter(BaseFormatter):
    """页面格式化器 - 负责页面设置、页眉页脚、页码等"""
    
    def setup_page_format(self, doc: Document):
        """设置页面格式"""
        section = doc.sections[0]
        
        # 设置页边距
        section.top_margin = self.config.PAGE_MARGINS['top']
        section.bottom_margin = self.config.PAGE_MARGINS['bottom']
        section.left_margin = self.config.PAGE_MARGINS['left']
        section.right_margin = self.config.PAGE_MARGINS['right']
        
        # 设置页眉页脚距离
        section.header_distance = Mm(25)
        section.footer_distance = Mm(25)
    
    def add_page_numbers(self, doc: Document):
        """添加页码"""
        section = doc.sections[0]
        footer = section.footer
        
        # 清除现有内容
        footer.paragraphs[0].clear()
        
        # 添加页码
        paragraph = footer.paragraphs[0]
        paragraph.alignment = self.config.ALIGNMENTS['center']
        
        run = paragraph.add_run("- ")
        run.font.name = self.config.FONTS['fangsong']
        run.font.size = self.config.FONT_SIZES['page_num']
        self._set_chinese_font(run, self.config.FONTS['fangsong'])
        
        # 插入页码字段
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = 'PAGE'
        
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        
        run2 = paragraph.add_run()
        run2._r.append(fldChar1)
        run2._r.append(instrText)
        run2._r.append(fldChar2)
        run2.font.name = self.config.FONTS['fangsong']
        run2.font.size = self.config.FONT_SIZES['page_num']
        self._set_chinese_font(run2, self.config.FONTS['fangsong'])
        
        run3 = paragraph.add_run(" -")
        run3.font.name = self.config.FONTS['fangsong']
        run3.font.size = self.config.FONT_SIZES['page_num']
        self._set_chinese_font(run3, self.config.FONTS['fangsong'])