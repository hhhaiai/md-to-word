from docx import Document
from typing import Dict, Any

from config import DocumentConfig
from formatters import (
    PageFormatter, 
    ParagraphFormatter, 
    DocumentTitleFormatter, 
    TableFormatter, 
    ListFormatter, 
    ImageFormatter
)


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
        self.image_formatter.format_images(self.doc)
        self.image_formatter.remove_image_captions(self.doc)
        
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