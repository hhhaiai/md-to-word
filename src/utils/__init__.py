"""
Utils module - 工具模块
导出常用的工具类和异常
"""

from .constants import Patterns, DocumentFormats
from .exceptions import (
    Md2WordError,
    ConfigurationError, 
    FileProcessingError,
    MarkdownParsingError,
    PandocError,
    DocumentFormattingError,
    ImageProcessingError,
    TableFormattingError,
    ListFormattingError,
    XMLProcessingError,
    PathSecurityError
)
from .xpath_cache import XPathCache, OptimizedXMLProcessor

__all__ = [
    # 常量和模式
    'Patterns',
    'DocumentFormats',
    
    # 异常类
    'Md2WordError',
    'ConfigurationError', 
    'FileProcessingError',
    'MarkdownParsingError', 
    'PandocError',
    'DocumentFormattingError',
    'ImageProcessingError',
    'TableFormattingError',
    'ListFormattingError',
    'XMLProcessingError',
    'PathSecurityError',
    
    # XML处理工具
    'XPathCache',
    'OptimizedXMLProcessor'
]