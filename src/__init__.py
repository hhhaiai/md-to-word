"""
md-to-word package - Markdown到Word公文格式转换工具
符合GB/T 9704-2012《党政机关公文格式》国家标准
"""

__version__ = "2.1.0"
__author__ = "md-to-word Project"
__description__ = "Markdown到Word公文格式转换工具"

# 导出主要组件
from .config import DocumentConfig
from .utils import (
    Patterns, 
    DocumentFormats,
    Md2WordError,
    FileProcessingError,
    PandocError,
    ImageProcessingError,
    XMLProcessingError,
    PathSecurityError,
    XPathCache,
    OptimizedXMLProcessor,
    ConfigValidator
)

__all__ = [
    # 配置
    'DocumentConfig',
    
    # 常量和模式
    'Patterns',
    'DocumentFormats',
    
    # 异常类
    'Md2WordError',
    'FileProcessingError', 
    'PandocError',
    'ImageProcessingError',
    'XMLProcessingError', 
    'PathSecurityError',
    
    # XML处理工具
    'XPathCache',
    'OptimizedXMLProcessor',
    
    # 配置验证
    'ConfigValidator'
]