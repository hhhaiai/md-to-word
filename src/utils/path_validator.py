"""
路径验证工具模块 - 提供安全的路径验证功能
"""
import os
from pathlib import Path
from typing import Optional
from .exceptions import PathSecurityError


def validate_safe_path(path: str, base_dir: Optional[str] = None, allow_absolute: bool = True) -> Path:
    """
    验证路径安全性，防止路径遍历攻击
    
    Args:
        path: 要验证的路径
        base_dir: 基础目录，如果提供则验证路径是否在基础目录内
        allow_absolute: 是否允许绝对路径
        
    Returns:
        Path: 安全的Path对象
        
    Raises:
        PathSecurityError: 如果路径不安全
    """
    if not path:
        raise PathSecurityError("路径不能为空")
    
    try:
        # 转换为Path对象并解析
        path_obj = Path(path)
        
        # 检查路径是否为绝对路径
        if path_obj.is_absolute() and not allow_absolute:
            raise PathSecurityError(f"不允许使用绝对路径: {path}")
        
        # 解析路径（包括符号链接）
        resolved_path = path_obj.resolve()
        
        # 检查路径组件
        parts = path_obj.parts
        for part in parts:
            # 检查危险的路径组件
            if part in [".", "..", "~"]:
                raise PathSecurityError(f"路径包含不安全的组件: {part}")
            # 检查隐藏文件/目录（以.开头）
            if part.startswith(".") and len(part) > 1:
                # 允许一些常见的配置文件
                if part not in [".gitignore", ".env", ".vscode", ".github"]:
                    raise PathSecurityError(f"路径包含隐藏文件/目录: {part}")
        
        # 如果提供了基础目录，确保路径在基础目录内
        if base_dir:
            base_path = Path(base_dir).resolve()
            try:
                # Python 3.9+
                if hasattr(resolved_path, 'is_relative_to'):
                    if not resolved_path.is_relative_to(base_path):
                        raise PathSecurityError(f"路径不在允许的目录内: {resolved_path}")
                else:
                    # Python 3.8 兼容性
                    try:
                        resolved_path.relative_to(base_path)
                    except ValueError:
                        raise PathSecurityError(f"路径不在允许的目录内: {resolved_path}")
            except ValueError:
                raise PathSecurityError(f"路径不在允许的目录内: {resolved_path}")
        
        # 检查路径是否包含特殊字符（Windows）
        if os.name == 'nt':
            invalid_chars = '<>:"|?*'
            filename = resolved_path.name
            if any(char in filename for char in invalid_chars):
                raise PathSecurityError(f"文件名包含无效字符: {filename}")
        
        return resolved_path
        
    except PathSecurityError:
        raise
    except Exception as e:
        raise PathSecurityError(f"路径验证失败: {e}")


def is_safe_relative_path(path: str) -> bool:
    """
    检查是否为安全的相对路径
    
    Args:
        path: 要检查的路径
        
    Returns:
        bool: 如果是安全的相对路径返回True
    """
    try:
        path_obj = Path(path)
        
        # 必须是相对路径
        if path_obj.is_absolute():
            return False
        
        # 不能包含危险的路径组件
        parts = path_obj.parts
        for part in parts:
            if part in [".", "..", "~"]:
                return False
                
        return True
    except Exception:
        return False