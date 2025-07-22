from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, Any, List
import os
import re

from config import DocumentConfig
from formatters import (
    PageFormatter, 
    ParagraphFormatter, 
    DocumentTitleFormatter, 
    TableFormatter, 
    ListFormatter, 
    ImageFormatter
)
from markdown_preprocessor import MarkdownPreprocessor


class WordPostprocessor:
    """
    重构后的Word文档后处理器
    使用组合模式，将不同的格式化功能委托给专门的格式化器类
    解决了原有的"God Object"反模式问题
    """
    
    def __init__(self):
        self.config = DocumentConfig()
        self.doc = None
        
        # 初始化专门的格式化器
        self.page_formatter = PageFormatter(self.config)
        self.paragraph_formatter = ParagraphFormatter(self.config)
        self.title_formatter = DocumentTitleFormatter(self.config)
        self.table_formatter = TableFormatter(self.config)
        self.list_formatter = ListFormatter(self.config)
        self.image_formatter = ImageFormatter(self.config)
    
    def apply_formatting(self, docx_path: str, metadata: Dict[str, Any], original_markdown: str = None) -> str:
        """
        对pandoc生成的Word文档应用公文格式
        
        Args:
            docx_path: pandoc生成的Word文档路径
            metadata: 包含标题、附件等元数据的字典
            original_markdown: 原始markdown内容，用于判断列表层级
            
        Returns:
            处理后的Word文档路径
        """
        # 加载pandoc生成的文档
        self.doc = Document(docx_path)
        
        # 保存原始markdown用于列表层级判断
        self.original_markdown = original_markdown
        
        # 保存文档路径的目录，用于查找图片
        self.doc_dir = os.path.dirname(os.path.abspath(docx_path))
        
        # 使用专门的格式化器处理不同方面的格式化
        self.page_formatter.setup_page_format(self.doc)
        self.paragraph_formatter.format_document_content(self.doc, metadata)
        
        # 添加文档标题（如果有）
        if metadata.get('title'):
            self.title_formatter.add_document_title(self.doc, metadata['title'])
        
        # 添加附件说明
        if metadata.get('attachments'):
            for attachment in metadata['attachments']:
                self.title_formatter.add_attachment(self.doc, attachment)
        
        # 应用各种格式化
        self.page_formatter.add_page_numbers(self.doc)
        self.list_formatter.format_lists(self.doc)
        self.table_formatter.format_tables(self.doc)
        
        # 新的图片处理方式：直接查找并替换图片语法
        self.process_and_insert_images()
        
        # 保存格式化后的文档
        self.doc.save(docx_path)
        return docx_path
    
    # 保留原有的公共方法以维持向后兼容性
    def format_tables(self):
        """格式化表格（向后兼容方法）"""
        if self.doc:
            self.table_formatter.format_tables(self.doc)
    
    def format_lists(self):
        """格式化列表（向后兼容方法）"""
        if self.doc:
            self.list_formatter.format_lists(self.doc)
    
    def format_images(self):
        """格式化图片（向后兼容方法）"""
        if self.doc:
            self.image_formatter.format_images(self.doc)
    
    def process_and_insert_images(self):
        """处理文档中的图片语法并插入实际图片"""
        # 初始化图片查找器
        self.image_finder = MarkdownPreprocessor()
        
        # 定义图片语法模式
        markdown_image_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')
        obsidian_image_pattern = re.compile(r'!\[\[([^\]]+)\]\]')
        
        # 更宽泛的caption识别模式，匹配各种图表格式
        caption_pattern = re.compile(r'^(\*\*)?[图表](?:片|格|表)?\s*\d+\s*[.:：]\s*')
        
        # 图片计数器
        image_counter = 0
        # 收集需要处理的图片信息
        images_to_process = []
        
        # 第一遍：识别所有图片和caption
        all_paragraphs = list(self.doc.paragraphs)
        for i, paragraph in enumerate(all_paragraphs):
            text = paragraph.text.strip()
            
            # 检查是否包含图片语法
            markdown_match = markdown_image_pattern.search(text)
            obsidian_match = obsidian_image_pattern.search(text)
            
            if markdown_match or obsidian_match:
                image_counter += 1
                
                # 检查是否在同一段落包含caption
                caption_in_same_para = None
                if markdown_match or obsidian_match:
                    # 提取图片语法后的文本作为potential caption
                    if markdown_match:
                        image_syntax_end = markdown_match.end()
                        remaining_text = text[image_syntax_end:].strip()
                        alt_text = markdown_match.group(1)
                        image_path = markdown_match.group(2)
                        image_type = 'markdown'
                    else:
                        image_syntax_end = obsidian_match.end()
                        remaining_text = text[image_syntax_end:].strip()
                        image_path = obsidian_match.group(1)
                        # 确定标题
                        if 'Pasted image' in image_path:
                            alt_text = ""  # 空标题，不显示
                        else:
                            alt_text = os.path.splitext(os.path.basename(image_path))[0]
                        image_type = 'obsidian'
                    
                    # 检查剩余文本是否为caption
                    if remaining_text and caption_pattern.match(remaining_text):
                        caption_in_same_para = remaining_text
                
                # 检查下一段落是否为caption（跳过空行）
                caption_paragraph = None
                for j in range(i + 1, min(i + 3, len(all_paragraphs))):  # 最多检查后面2个段落
                    next_text = all_paragraphs[j].text.strip()
                    if not next_text:  # 跳过空段落
                        continue
                    if caption_pattern.match(next_text):
                        caption_paragraph = all_paragraphs[j]
                        break
                    else:
                        break  # 遇到非空、非caption段落就停止
                
                # 查找图片实际路径
                actual_path = self._find_image_actual_path(image_path)
                
                # 构建图片信息
                image_info = {
                    'path': actual_path,
                    'title': alt_text,
                    'type': image_type,
                    'original': text,
                    'number': image_counter,
                    'paragraph': paragraph,
                    'caption_in_same_para': caption_in_same_para,
                    'caption_paragraph': caption_paragraph
                }
                
                images_to_process.append(image_info)
        
        # 第二遍：处理图片插入
        for image_info in images_to_process:
            self._replace_paragraph_with_image(image_info['paragraph'], image_info)
        
        # 第三遍：处理caption格式化
        self._process_captions()
    
    def _find_image_actual_path(self, image_path: str) -> str:
        """查找图片的实际路径"""
        # 如果是绝对路径或URL，直接返回
        if os.path.isabs(image_path) or image_path.startswith(('http://', 'https://')):
            return image_path
        
        # 使用图片查找器查找实际路径
        actual_path = self.image_finder._find_image_path(image_path, self.doc_dir)
        return actual_path
    
    def _replace_paragraph_with_image(self, paragraph, image_info: Dict):
        """将包含图片语法的段落替换为实际图片"""
        try:
            # 检查图片路径是否存在
            if not image_info['path'] or not os.path.exists(image_info['path']):
                # 如果图片不存在，保留原文本但显示警告
                print(f"警告：找不到图片 {image_info.get('original', 'unknown')}")
                return
            
            # 处理同一段落中的caption分离
            if image_info.get('caption_in_same_para'):
                # 在当前段落后插入caption段落
                p_element = paragraph._element
                parent = p_element.getparent()
                
                # 创建caption段落
                caption_p = self.doc.add_paragraph()
                caption_p.text = image_info['caption_in_same_para']
                
                # 将caption段落移动到图片段落之后
                parent.insert(parent.index(p_element) + 1, caption_p._element)
            
            # 清空原段落文本
            paragraph.clear()
            
            # 添加图片到段落
            run = paragraph.add_run()
            picture = run.add_picture(image_info['path'], width=Inches(5))  # 默认宽度5英寸
            
            # 设置段落居中
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 设置图片文字环绕
            drawing_elements = run._element.xpath('.//w:drawing')
            if drawing_elements:
                self.image_formatter._set_image_wrap(drawing_elements[0])
                
        except Exception as e:
            print(f"插入图片时出错: {e}")
            # 保留原始文本
            paragraph.text = image_info['original']
    
    def _process_captions(self):
        """处理并格式化所有图片和表格caption，统一放在图表下方"""
        try:
            # 更宽泛的caption识别和标准化模式
            # 匹配: 图/图片/图表 + 可选空格 + 数字 + 可选空格 + [.:：] + 内容
            image_pattern = re.compile(r'^(\*\*)?图(?:片|表)?\s*(\d+)\s*[.:：]\s*(.*?)(\*\*)?$')
            # 匹配: 表/表格 + 可选空格 + 数字 + 可选空格 + [.:：] + 内容  
            table_pattern = re.compile(r'^(\*\*)?表(?:格)?\s*(\d+)\s*[.:：]\s*(.*?)(\*\*)?$')
            
            # 第一步：标准化所有caption格式，但不移动位置
            for paragraph in self.doc.paragraphs:
                text = paragraph.text.strip()
                if text:
                    # 处理图片caption
                    image_match = image_pattern.match(text)
                    if image_match:
                        number = image_match.group(2)
                        content = image_match.group(3)
                        # 标准化为: 图1. 内容
                        paragraph.text = f"图{number}. {content}"
                        self.image_formatter._format_image_caption(paragraph)
                    
                    else:
                        # 处理表格caption
                        table_match = table_pattern.match(text)
                        if table_match:
                            number = table_match.group(2)
                            content = table_match.group(3)
                            # 标准化为: 表1. 内容
                            paragraph.text = f"表{number}. {content}"
                            self._format_table_caption(paragraph)
            
            print("Caption格式标准化完成，位置保持不变")
                        
        except Exception as e:
            print(f"处理caption时出错: {e}")
    
    def _format_table_caption(self, paragraph):
        """格式化表格caption，与图片caption相同的格式"""
        try:
            # 使用与图片caption相同的格式设置
            for run in paragraph.runs:
                run.font.name = self.config.FONTS['fangsong']
                run.font.size = self.config.FONT_SIZES['table']  # 4号字体
                run.bold = False
                # 设置中文字体
                from docx.oxml.ns import qn
                run._element.rPr.rFonts.set(qn('w:eastAsia'), self.config.FONTS['fangsong'])
            
            # 设置段落格式 - 表格caption居中显示
            paragraph.alignment = self.config.ALIGNMENTS['center']
            paragraph_format = paragraph.paragraph_format
            paragraph_format.line_spacing = self.config.LINE_SPACING
            paragraph_format.space_after = Pt(6)
            paragraph_format.space_before = Pt(3)
            paragraph_format.first_line_indent = Pt(0)  # caption不缩进
            
        except AttributeError as e:
            print(f"警告：格式化表格caption时出现属性错误: {e}")
        except Exception as e:
            print(f"警告：格式化表格caption时出现错误: {e}")
    
