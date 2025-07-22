import pypandoc
import tempfile
import os
import shutil
from pathlib import Path
from typing import Dict, Any, Optional
from docx import Document
from config import DocumentConfig
from exceptions import PandocError

class PandocProcessor:
    """使用Pandoc进行Markdown到Word的转换处理器"""
    
    def __init__(self):
        self.temp_files = []  # 用于跟踪临时文件以便清理
        self.image_config = DocumentConfig.IMAGE_CONFIG
        self.pandoc_config = DocumentConfig.PANDOC_CONFIG
    
    def convert_markdown_to_docx(self, markdown_content: str, output_path: str, 
                                title: str = None, extra_args: Optional[list] = None) -> str:
        """
        将预处理后的Markdown内容转换为Word文档
        
        Args:
            markdown_content: 预处理后的Markdown内容
            output_path: 输出的Word文档路径
            title: 文档标题（可选）
            extra_args: 额外的pandoc参数
            
        Returns:
            生成的Word文档路径
        """
        try:
            # 创建临时Markdown文件
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', 
                                           suffix='.md', delete=False) as temp_md:
                # 如果提供了标题，添加到内容前面
                if title:
                    # 使用一级标题格式添加文档标题
                    content_with_title = f"# {title}\n\n{markdown_content}"
                else:
                    content_with_title = markdown_content
                    
                temp_md.write(content_with_title)
                temp_md_path = temp_md.name
                self.temp_files.append(temp_md_path)
            
            # 设置pandoc转换参数
            pandoc_args = self._get_pandoc_args()
            if extra_args:
                pandoc_args.extend(extra_args)
            
            # 使用subprocess直接调用pandoc（确保数学公式正确处理）
            import subprocess
            
            cmd = ['pandoc', temp_md_path, '-t', 'docx', '-o', output_path] + pandoc_args
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                raise PandocError(f"Pandoc命令执行失败: {result.stderr}")
            
            return output_path
            
        except pypandoc.PandocError as e:
            raise PandocError(f"Pandoc转换失败: {e}")
        except OSError as e:
            raise PandocError(f"文件操作失败: {e}")
        except Exception as e:
            raise PandocError(f"Pandoc处理时发生未知错误: {e}")
        finally:
            # 清理临时文件
            self._cleanup_temp_files()
    
    def _get_pandoc_args(self) -> list:
        """获取pandoc转换参数"""
        args = [
            # 数学公式支持
            f'--{self.pandoc_config["math_method"]}',
            # 保持原始格式的某些方面
            '--preserve-tabs',
            # 处理换行
            '--wrap=none',
        ]
        
        # 添加配置文件中的额外参数
        if 'extra_args' in self.pandoc_config:
            args.extend(self.pandoc_config['extra_args'])
        
        return args
    
    
    def _cleanup_temp_files(self):
        """清理临时文件"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except OSError:
                pass  # 忽略清理时的文件系统错误
        self.temp_files.clear()
    
    def load_docx_for_postprocessing(self, docx_path: str) -> Document:
        """
        加载pandoc生成的docx文件以便进行后处理
        
        Args:
            docx_path: Word文档路径
            
        Returns:
            python-docx Document对象
        """
        return Document(docx_path)
    
    def check_pandoc_available(self) -> bool:
        """检查pandoc是否可用"""
        try:
            pypandoc.get_pandoc_version()
            return True
        except (OSError, RuntimeError):
            return False
    
