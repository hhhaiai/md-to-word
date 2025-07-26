"""
Utils module - 工具模块
导出常用的工具类和异常
"""

from .constants import Patterns, DocumentFormats
from .exceptions import (
    Md2WordError,
    FileProcessingError,
    PandocError,
    ImageProcessingError,
    XMLProcessingError,
    PathSecurityError
)
from .xpath_cache import XPathCache, OptimizedXMLProcessor
from .config_validator import ConfigValidator
from .path_validator import validate_safe_path, is_safe_relative_path

__all__ = [
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
    'ConfigValidator',
    
    # 路径验证
    'validate_safe_path',
    'is_safe_relative_path'
]