import pypandoc
import tempfile
import os
from pathlib import Path
from typing import Dict, Any, Optional
from docx import Document

class PandocProcessor:
    """使用Pandoc进行Markdown到Word的转换处理器"""
    
    def __init__(self):
        self.temp_files = []  # 用于跟踪临时文件以便清理
    
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
            
            # 使用pypandoc进行转换
            pypandoc.convert_file(
                temp_md_path,
                'docx',
                outputfile=output_path,
                extra_args=pandoc_args
            )
            
            return output_path
            
        except Exception as e:
            raise Exception(f"Pandoc转换失败: {e}")
        finally:
            # 清理临时文件
            self._cleanup_temp_files()
    
    def _get_pandoc_args(self) -> list:
        """获取pandoc转换参数"""
        args = [
            # 数学公式支持
            '--mathml',  # 使用MathML渲染数学公式
            # 保持原始格式的某些方面
            '--preserve-tabs',
            # 处理换行
            '--wrap=none',
        ]
        # 先不使用lua filter，避免复杂性
        # filter_path = self._get_chinese_filter_path()
        # if filter_path:
        #     args.append(f'--lua-filter={filter_path}')
        return args
    
    def _get_chinese_filter_path(self) -> Optional[str]:
        """获取中文处理Lua过滤器路径（如果存在）"""
        current_dir = Path(__file__).parent
        filter_path = current_dir / 'chinese_filter.lua'
        if filter_path.exists():
            return str(filter_path)
        return None
    
    def _cleanup_temp_files(self):
        """清理临时文件"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception:
                pass  # 忽略清理错误
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
        except Exception:
            return False
    
    def get_supported_formats(self) -> Dict[str, Any]:
        """获取支持的格式信息"""
        try:
            from_formats = pypandoc.get_pandoc_formats()[0]
            to_formats = pypandoc.get_pandoc_formats()[1]
            return {
                'from': from_formats,
                'to': to_formats,
                'version': pypandoc.get_pandoc_version()
            }
        except Exception as e:
            return {'error': str(e)}