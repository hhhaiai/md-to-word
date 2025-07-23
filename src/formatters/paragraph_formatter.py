"""
段落格式化器 - 负责标题和正文段落的格式化
"""
from docx import Document
from docx.shared import Pt, RGBColor
from typing import Dict, Any
from .base_formatter import BaseFormatter
from ..utils.constants import Patterns


class ParagraphFormatter(BaseFormatter):
    """段落格式化器 - 负责标题和正文段落的格式化"""
    
    def format_document_content(self, doc: Document, metadata: Dict[str, Any]):
        """格式化文档内容，应用公文格式要求"""
        for paragraph in doc.paragraphs:
            # 跳过空段落
            if not paragraph.text.strip():
                continue
            
            # 判断段落类型并应用相应格式
            if self._is_heading(paragraph):
                self._format_heading(paragraph)
            else:
                self._format_body_paragraph(paragraph)
    
    def _is_heading(self, paragraph) -> bool:
        """判断段落是否为标题"""
        # 检查是否使用了标题样式
        if paragraph.style.name.startswith('Heading'):
            return True
        
        # 检查段落内容是否以中文数字或序号开头（如"一、"、"（一）"等）
        text = paragraph.text.strip()
        
        for pattern in Patterns.HEADING_PATTERNS:
            if pattern.match(text):
                return True
        
        return False
    
    def _format_heading(self, paragraph):
        """格式化标题段落"""
        text = paragraph.text.strip()
        level = self._get_heading_level(paragraph, text)
        
        # 检查是否包含数学公式（通过检查XML元素）
        if self._has_math_formula(paragraph):
            # 如果包含数学公式，只修改字体格式，不清除内容
            self._format_paragraph_with_math(paragraph, level, is_heading=True)
            return
        
        # 对于不包含数学公式的标题，使用原有的方法
        # 清除现有格式
        for run in paragraph.runs:
            run.clear()
        
        # 重新添加文本
        run = paragraph.add_run(text)
        
        # 根据级别应用格式
        if level == 1:
            # 一级标题：黑体，三号，不加粗
            run.font.name = self.config.FONTS['heiti']
            run.font.size = self.config.FONT_SIZES['body']
            run.bold = False
            # 设置字体颜色为黑色
            run.font.color.rgb = RGBColor(0, 0, 0)
            self._set_chinese_font(run, self.config.FONTS['heiti'])
        elif level == 2:
            # 二级标题：楷体，三号，不加粗
            run.font.name = self.config.FONTS['kaiti']
            run.font.size = self.config.FONT_SIZES['body']
            run.bold = False
            # 设置字体颜色为黑色
            run.font.color.rgb = RGBColor(0, 0, 0)
            self._set_chinese_font(run, self.config.FONTS['kaiti'])
        else:
            # 三级及以下标题：仿宋，三号，不加粗
            run.font.name = self.config.FONTS['fangsong']
            run.font.size = self.config.FONT_SIZES['body']
            run.bold = False
            # 设置字体颜色为黑色
            run.font.color.rgb = RGBColor(0, 0, 0)
            self._set_chinese_font(run, self.config.FONTS['fangsong'])
        
        # 设置段落格式
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        paragraph_format = paragraph.paragraph_format
        paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
        paragraph_format.line_spacing = self.config.LINE_SPACING
        paragraph_format.space_after = Pt(6)
        paragraph_format.space_before = Pt(6)
    
    def _get_heading_level(self, paragraph, text: str) -> int:
        """获取标题级别"""
        # 检查Word内置标题样式 - pandoc生成的样式映射（修正版）
        # 因为pandoc将#作为Heading 1，我们跳过#，所以：
        # ## → Heading 2 → 应该是我们的一级标题（黑体）
        # ### → Heading 3 → 应该是我们的二级标题（楷体）  
        # #### → Heading 4 → 应该是我们的三级标题（仿宋）
        if paragraph.style.name == 'Heading 1':
            return 1  # # → 黑体 (但我们通常跳过这个)
        elif paragraph.style.name == 'Heading 2':
            return 1  # ## → 黑体
        elif paragraph.style.name == 'Heading 3':
            return 2  # ### → 楷体
        elif paragraph.style.name == 'Heading 4':
            return 3  # #### → 仿宋
        elif paragraph.style.name == 'Heading 5':
            return 3  # ##### → 仿宋 (fallback)
        
        # 根据文本内容判断级别（处理中文标题格式）
        if Patterns.HEADING_PATTERNS[0].match(text):
            return 1  # 一、二、三、 → 黑体
        elif Patterns.HEADING_PATTERNS[1].match(text):
            return 2  # （一）（二）（三） → 楷体
        elif Patterns.HEADING_PATTERNS[2].match(text) or Patterns.HEADING_PATTERNS[3].match(text):
            return 3  # 1. 2. 3. 或 1、2、3、 → 仿宋
        
        return 3  # 默认三级标题
    
    def _format_body_paragraph(self, paragraph):
        """格式化正文段落"""
        # 检查是否包含数学公式
        if self._has_math_formula(paragraph):
            # 如果包含数学公式，使用特殊的格式化方法
            self._format_paragraph_with_math(paragraph, level=0, is_heading=False)
            return
        
        # 为所有运行应用仿宋格式
        for run in paragraph.runs:
            run.font.name = self.config.FONTS['fangsong']
            run.font.size = self.config.FONT_SIZES['body']
            self._set_chinese_font(run, self.config.FONTS['fangsong'])
        
        # 设置段落格式
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        paragraph_format = paragraph.paragraph_format
        paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
        paragraph_format.line_spacing = self.config.LINE_SPACING
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
    
    def _format_paragraph_with_math(self, paragraph, level: int, is_heading: bool):
        """格式化包含数学公式的段落，保留数学公式内容"""
        # 设置段落格式
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        paragraph_format = paragraph.paragraph_format
        
        if is_heading:
            paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
            paragraph_format.space_after = Pt(6)
            paragraph_format.space_before = Pt(6)
        else:
            paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
            paragraph_format.space_after = Pt(0)
            paragraph_format.space_before = Pt(0)
        
        paragraph_format.line_spacing = self.config.LINE_SPACING
        
        # 只格式化文本run，跳过数学公式
        for run in paragraph.runs:
            if run._element.tag.endswith('r'):  # 普通文本run
                try:
                    # 检查run的XML内容，只处理不包含数学内容的run
                    if 'oMath' not in run._element.xml:
                        if is_heading:
                            # 根据标题级别设置字体
                            if level == 1:
                                run.font.name = self.config.FONTS['heiti']
                                self._set_chinese_font(run, self.config.FONTS['heiti'])
                            elif level == 2:
                                run.font.name = self.config.FONTS['kaiti']
                                self._set_chinese_font(run, self.config.FONTS['kaiti'])
                            else:
                                run.font.name = self.config.FONTS['fangsong']
                                self._set_chinese_font(run, self.config.FONTS['fangsong'])
                            run.font.size = self.config.FONT_SIZES['body']
                            run.bold = False
                            run.font.color.rgb = RGBColor(0, 0, 0)
                        else:
                            # 正文段落格式
                            run.font.name = self.config.FONTS['fangsong']
                            run.font.size = self.config.FONT_SIZES['body']
                            self._set_chinese_font(run, self.config.FONTS['fangsong'])
                except:
                    # 如果出错，跳过这个run
                    continue