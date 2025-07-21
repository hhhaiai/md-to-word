from docx import Document
from docx.shared import Pt, Mm
from docx.oxml.shared import OxmlElement, qn
from typing import Dict, Any
import re

from config import DocumentConfig

class WordPostprocessor:
    """Word文档后处理器，对pandoc生成的Word文档应用GB/T 9704-2012格式要求"""
    
    def __init__(self):
        self.config = DocumentConfig()
        self.doc = None
    
    def apply_formatting(self, docx_path: str, metadata: Dict[str, Any], original_markdown: str = None) -> str:
        """
        对pandoc生成的Word文档应用公文格式
        
        Args:
            docx_path: pandoc生成的Word文档路径
            metadata: 包含标题、日期、附件等元数据的字典
            original_markdown: 原始markdown内容，用于判断列表层级
            
        Returns:
            处理后的Word文档路径
        """
        # 加载pandoc生成的文档
        self.doc = Document(docx_path)
        
        # 保存原始markdown用于列表层级判断
        self.original_markdown = original_markdown
        
        # 设置页面格式
        self._setup_page_format()
        
        # 处理文档内容格式
        self._format_document_content(metadata)
        
        # 添加文档标题（如果有）
        if metadata.get('title'):
            self._add_document_title(metadata['title'])
        
        # 添加附件说明
        if metadata.get('attachments'):
            for attachment in metadata['attachments']:
                self._add_attachment(attachment)
        
        # 添加成文日期
        if metadata.get('date'):
            self._add_date(metadata['date'])
        
        # 添加页码
        self._add_page_numbers()
        
        # 格式化列表
        self.format_lists()
        
        # 格式化表格
        self.format_tables()
        
        # 保存格式化后的文档
        self.doc.save(docx_path)
        return docx_path
    
    def _setup_page_format(self):
        """设置页面格式"""
        section = self.doc.sections[0]
        
        # 设置页边距
        section.top_margin = self.config.PAGE_MARGINS['top']
        section.bottom_margin = self.config.PAGE_MARGINS['bottom']
        section.left_margin = self.config.PAGE_MARGINS['left']
        section.right_margin = self.config.PAGE_MARGINS['right']
        
        # 设置页眉页脚距离
        section.header_distance = Mm(25)
        section.footer_distance = Mm(25)
    
    def _format_document_content(self, metadata: Dict[str, Any]):
        """格式化文档内容，应用公文格式要求"""
        for paragraph in self.doc.paragraphs:
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
        heading_patterns = [
            r'^[一二三四五六七八九十]+、',  # 一、二、三、
            r'^（[一二三四五六七八九十]+）',  # （一）（二）（三）
            r'^[0-9]+\.',  # 1. 2. 3.
            r'^[0-9]+、',  # 1、2、3、
        ]
        
        for pattern in heading_patterns:
            if re.match(pattern, text):
                return True
        
        return False
    
    def _format_heading(self, paragraph):
        """格式化标题段落"""
        text = paragraph.text.strip()
        level = self._get_heading_level(paragraph, text)
        
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
            self._set_chinese_font(run, self.config.FONTS['heiti'])
        elif level == 2:
            # 二级标题：楷体，三号，不加粗
            run.font.name = self.config.FONTS['kaiti']
            run.font.size = self.config.FONT_SIZES['body']
            run.bold = False
            self._set_chinese_font(run, self.config.FONTS['kaiti'])
        else:
            # 三级及以下标题：仿宋，三号，不加粗
            run.font.name = self.config.FONTS['fangsong']
            run.font.size = self.config.FONT_SIZES['body']
            run.bold = False
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
        if re.match(r'^[一二三四五六七八九十]+、', text):
            return 1  # 一、二、三、 → 黑体
        elif re.match(r'^（[一二三四五六七八九十]+）', text):
            return 2  # （一）（二）（三） → 楷体
        elif re.match(r'^[0-9]+\.', text) or re.match(r'^[0-9]+、', text):
            return 3  # 1. 2. 3. 或 1、2、3、 → 仿宋
        
        return 3  # 默认三级标题
    
    def _format_body_paragraph(self, paragraph):
        """格式化正文段落"""
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
    
    def _add_document_title(self, title: str):
        """在文档开头添加标题"""
        # 在文档开头插入标题段落
        title_paragraph = self.doc.paragraphs[0].insert_paragraph_before()
        title_paragraph.alignment = self.config.ALIGNMENTS['center']
        
        run = title_paragraph.add_run(title)
        run.font.name = self.config.FONTS['xiaobiaosong']
        run.font.size = self.config.FONT_SIZES['title']
        run.bold = True
        
        # 设置中文字体
        self._set_chinese_font(run, self.config.FONTS['xiaobiaosong'])
        
        # 添加空行
        self.doc.paragraphs[1].insert_paragraph_before()
    
    def _add_attachment(self, attachment: str):
        """添加附件说明"""
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = self.config.ALIGNMENTS['justify']
        
        run = paragraph.add_run(attachment)
        run.font.name = self.config.FONTS['fangsong']
        run.font.size = self.config.FONT_SIZES['body']
        
        self._set_chinese_font(run, self.config.FONTS['fangsong'])
        
        # 设置段落格式
        paragraph_format = paragraph.paragraph_format
        paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
        paragraph_format.line_spacing = self.config.LINE_SPACING
        paragraph_format.space_after = Pt(0)
    
    def _add_date(self, date: str):
        """添加成文日期"""
        paragraph = self.doc.add_paragraph()
        paragraph.alignment = self.config.ALIGNMENTS['right']
        
        run = paragraph.add_run(date)
        run.font.name = self.config.FONTS['fangsong']
        run.font.size = self.config.FONT_SIZES['body']
        
        self._set_chinese_font(run, self.config.FONTS['fangsong'])
        
        # 设置段落格式
        paragraph_format = paragraph.paragraph_format
        paragraph_format.line_spacing = self.config.LINE_SPACING
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(12)
    
    def _add_page_numbers(self):
        """添加页码"""
        section = self.doc.sections[0]
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
    
    def _set_chinese_font(self, run, font_name: str):
        """设置中文字体"""
        rPr = run._element.rPr
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            run._element.insert(0, rPr)
        
        # 设置中文字体
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        
        rFonts.set(qn('w:eastAsia'), font_name)
        rFonts.set(qn('w:hint'), 'eastAsia')
    
    def format_tables(self):
        """格式化表格"""
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = self.config.FONTS['fangsong']
                            run.font.size = self.config.FONT_SIZES['body']
                            self._set_chinese_font(run, self.config.FONTS['fangsong'])
                        
                        # 设置单元格段落格式
                        paragraph.alignment = self.config.ALIGNMENTS['justify']
                        paragraph_format = paragraph.paragraph_format
                        paragraph_format.line_spacing = self.config.LINE_SPACING
    
    def format_lists(self):
        """格式化列表 - 无序列表首级不缩进，有序列表作为段落处理"""
        for i, paragraph in enumerate(self.doc.paragraphs):
            
            # 检查是否为列表项 - Compact样式通常是列表
            if paragraph.style.name == 'Compact':
                paragraph_format = paragraph.paragraph_format
                full_text = paragraph.text.strip()
                
                print(f"DEBUG: 处理列表 '{full_text}'")
                
                # 通过文本内容判断是否为二级列表
                is_sub_item = '子项' in full_text
                
                if is_sub_item:
                    print("DEBUG: 识别为二级列表（子项）")
                    # 二级列表：○ 符号，缩进0.6cm
                    paragraph_format.left_indent = Mm(6)
                    paragraph_format.first_line_indent = Pt(0)
                    paragraph_format.space_after = Pt(0)
                    paragraph_format.space_before = Pt(0)
                    paragraph.alignment = self.config.ALIGNMENTS['justify']
                    
                    final_text = '○' + full_text
                else:
                    print("DEBUG: 识别为一级列表")
                    # 一级列表：• 符号，不缩进
                    paragraph_format.left_indent = Pt(0)
                    paragraph_format.first_line_indent = Pt(0)
                    paragraph_format.space_after = Pt(0)
                    paragraph_format.space_before = Pt(0)
                    paragraph.alignment = self.config.ALIGNMENTS['justify']
                    
                    final_text = '•' + full_text
                
                print(f"DEBUG: 最终文本: '{final_text}'")
                
                # 重新设置段落内容
                paragraph.clear()
                run = paragraph.add_run(final_text)
                run.font.name = self.config.FONTS['fangsong']
                run.font.size = self.config.FONT_SIZES['body']
                self._set_chinese_font(run, self.config.FONTS['fangsong'])
            
            # 检查有序列表 - 包含数字序号的段落
            elif re.match(r'^\d+\.', paragraph.text.strip()):
                paragraph_format = paragraph.paragraph_format
                
                # 检查是否需要分割段落（包含多个列表项）
                text = paragraph.text.strip()
                items = re.split(r'(\d+\.)', text)[1:]  # 分割并保留序号
                
                if len(items) > 2:  # 包含多个列表项
                    
                    # 保存当前段落位置
                    current_p = paragraph._element
                    parent = current_p.getparent()
                    
                    # 重新组织项目
                    paired_items = []
                    for i in range(0, len(items), 2):
                        if i + 1 < len(items):
                            paired_items.append(items[i] + items[i + 1])
                    
                    # 更新第一个段落为第一项
                    if paired_items:
                        paragraph.clear()
                        first_run = paragraph.add_run(paired_items[0].strip())
                        first_run.font.name = self.config.FONTS['fangsong']
                        first_run.font.size = self.config.FONT_SIZES['body']
                        self._set_chinese_font(first_run, self.config.FONTS['fangsong'])
                        
                        paragraph_format.left_indent = Pt(0)
                        paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
                        paragraph_format.space_after = Pt(0)
                        paragraph_format.space_before = Pt(0)
                        paragraph.alignment = self.config.ALIGNMENTS['justify']
                        
                        # 为剩余项目创建新段落
                        for item_text in paired_items[1:]:
                            new_p = paragraph._element.makeelement(paragraph._element.tag, paragraph._element.attrib)
                            parent.insert(parent.index(current_p) + 1, new_p)
                            current_p = new_p
                            
                            # 创建新的段落对象
                            from docx.text.paragraph import Paragraph
                            new_paragraph = Paragraph(new_p, paragraph._parent)
                            
                            # 添加文本
                            new_run = new_paragraph.add_run(item_text.strip())
                            new_run.font.name = self.config.FONTS['fangsong']
                            new_run.font.size = self.config.FONT_SIZES['body']
                            self._set_chinese_font(new_run, self.config.FONTS['fangsong'])
                            
                            # 设置格式
                            new_paragraph_format = new_paragraph.paragraph_format
                            new_paragraph_format.left_indent = Pt(0)
                            new_paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
                            new_paragraph_format.space_after = Pt(0)
                            new_paragraph_format.space_before = Pt(0)
                            new_paragraph.alignment = self.config.ALIGNMENTS['justify']
                else:
                    # 单个有序列表项，直接格式化
                    paragraph_format.left_indent = Pt(0)
                    paragraph_format.first_line_indent = self.config.FIRST_LINE_INDENT
                    paragraph_format.space_after = Pt(0)
                    paragraph_format.space_before = Pt(0)
                    paragraph.alignment = self.config.ALIGNMENTS['justify']
                    
                    # 确保列表项使用正确字体
                    for run in paragraph.runs:
                        run.font.name = self.config.FONTS['fangsong']
                        run.font.size = self.config.FONT_SIZES['body']
                        self._set_chinese_font(run, self.config.FONTS['fangsong'])