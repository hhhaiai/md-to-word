import re
import os
import shutil
from typing import List, Dict, Any
from config import DocumentConfig

class MarkdownPreprocessor:
    """Markdown预处理器，用于清理和过滤Markdown内容后交给pandoc处理"""
    
    def __init__(self):
        self.image_config = DocumentConfig.IMAGE_CONFIG
    
    def preprocess_file(self, file_path: str) -> Dict[str, Any]:
        """预处理Markdown文件，返回处理结果"""
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 获取文件名作为标题（去掉扩展名）
        filename = os.path.basename(file_path)
        title_from_filename = os.path.splitext(filename)[0]
        
        # 预处理内容
        processed_content = self.preprocess_content(content, file_path)
        
        # 提取元数据
        metadata = self._extract_metadata(content)
        
        return {
            'title': title_from_filename,
            'content': processed_content,
            'attachments': metadata.get('attachments', [])
        }
    
    def preprocess_content(self, content: str, file_path: str = '') -> str:
        """预处理Markdown内容"""
        lines = content.split('\n')
        
        # 应用所有过滤器
        lines = self._filter_yaml_frontmatter(lines)
        lines = self._filter_ending_metadata(lines)
        lines = self._remove_bold_formatting(lines)
        lines = self._remove_numbered_list_spaces(lines)
        lines = self._convert_ordered_lists_to_paragraphs(lines)
        lines = self._fix_unordered_list_asterisks(lines)
        lines = self._process_math_formulas(lines)
        if file_path:
            lines = self._process_images(lines, os.path.dirname(file_path))
        lines = self._merge_broken_lines(lines)
        lines = self._skip_first_level_headers(lines)
        
        # 重新组合内容
        processed_content = '\n'.join(lines)
        
        return processed_content.strip()
    
    def _extract_metadata(self, content: str) -> Dict[str, Any]:
        """提取附件信息"""
        lines = content.split('\n')
        metadata = {
            'attachments': []
        }
        
        for line in lines:
            line = line.strip()
            
            # 检测附件说明
            if line.startswith('附件') or '附件' in line:
                metadata['attachments'].append(line)
        
        return metadata
    

    
    def _filter_yaml_frontmatter(self, lines: List[str]) -> List[str]:
        """过滤YAML front matter"""
        if not lines or not lines[0].strip() == '---':
            return lines
        
        # 找到第二个---的位置
        end_index = -1
        for i in range(1, len(lines)):
            if lines[i].strip() == '---':
                end_index = i
                break
        
        if end_index != -1:
            # 跳过YAML front matter部分
            return lines[end_index + 1:]
        
        return lines
    
    def _filter_ending_metadata(self, lines: List[str]) -> List[str]:
        """过滤结尾的Date和标签"""
        # 从后往前查找最后一个实质性内容的位置
        last_content_index = len(lines) - 1
        
        for i in range(len(lines) - 1, -1, -1):
            line_stripped = lines[i].strip()
            
            # 跳过空行、Date行、单词标签（如#work）、---分隔符
            if (line_stripped == '' or 
                line_stripped.startswith('Date:') or
                line_stripped == '---' or
                (line_stripped.startswith('#') and ' ' not in line_stripped and not line_stripped.startswith('##'))):
                continue
            else:
                last_content_index = i
                break
        
        # 返回到最后实质内容位置的所有行
        return lines[:last_content_index + 1]
    
    def _remove_bold_formatting(self, lines: List[str]) -> List[str]:
        """去除加粗标记，保留文字内容"""
        processed_lines = []
        
        for line in lines:
            # 去除 **加粗** 标记
            processed_line = re.sub(r'\*\*(.*?)\*\*', r'\1', line)
            # 去除 __加粗__ 标记（另一种Markdown加粗语法）
            processed_line = re.sub(r'__(.*?)__', r'\1', processed_line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _remove_numbered_list_spaces(self, lines: List[str]) -> List[str]:
        """去除数字列表项中的空格，如将'1. '转换为'1.'"""
        processed_lines = []
        
        for line in lines:
            # 匹配数字后跟点和空格的模式，如 "1. ", "2. ", "10. " 等
            processed_line = re.sub(r'^(\s*)(\d+)\.\s+', r'\1\2.', line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _process_math_formulas(self, lines: List[str]) -> List[str]:
        """处理数学公式 - 保持LaTeX格式让pandoc处理"""
        # 现在我们让pandoc处理数学公式，所以保持原始格式
        # 但我们仍然可以做一些清理工作
        processed_lines = []
        
        for line in lines:
            # 保持原始的LaTeX数学公式格式
            # pandoc会正确处理$...$和$$...$$格式
            processed_lines.append(line)
        
        return processed_lines
    
    def _process_images(self, lines: List[str], source_dir: str) -> List[str]:
        """处理图片链接，将相对路径转换为绝对路径或复制图片"""
        processed_lines = []
        
        for line in lines:
            # 处理标准Markdown图片语法: ![alt](path)
            markdown_pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
            # 处理Obsidian图片语法: ![[path]]
            obsidian_pattern = r'!\[\[([^\]]+)\]\]'
            
            def replace_markdown_image(match):
                alt_text = match.group(1)
                image_path = match.group(2)
                
                # 如果已经是绝对路径或URL，不处理
                if os.path.isabs(image_path) or image_path.startswith(('http://', 'https://')):
                    return match.group(0)
                
                # 尝试找到图片的实际路径
                actual_path = self._find_image_path(image_path, source_dir)
                
                if actual_path:
                    # 返回绝对路径格式的图片引用
                    return f'![{alt_text}]({actual_path})'
                else:
                    # 如果找不到图片，保持原始引用
                    print(f"警告：找不到图片 {image_path}")
                    return match.group(0)
            
            def replace_obsidian_image(match):
                image_path = match.group(1)
                
                # 尝试找到图片的实际路径
                actual_path = self._find_image_path(image_path, source_dir)
                
                if actual_path:
                    # 转换为标准Markdown格式，使用文件名作为alt text
                    filename = os.path.splitext(os.path.basename(image_path))[0]
                    return f'![{filename}]({actual_path})'
                else:
                    # 如果找不到图片，保持原始引用
                    print(f"警告：找不到图片 {image_path}")
                    return match.group(0)
            
            # 先处理Obsidian格式，再处理标准Markdown格式
            processed_line = re.sub(obsidian_pattern, replace_obsidian_image, line)
            processed_line = re.sub(markdown_pattern, replace_markdown_image, processed_line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _find_image_path(self, image_name: str, source_dir: str) -> str:
        """在配置的搜索路径中查找图片"""
        
        # 构建搜索路径列表
        search_paths = [
            # 首先在源文件目录查找
            source_dir,
            # 然后在配置的搜索路径中查找
            *self.image_config['search_paths']
        ]
        
        # 支持的图片格式
        supported_formats = self.image_config['supported_formats']
        
        for search_path in search_paths:
            if not os.path.exists(search_path):
                continue
                
            # 直接匹配文件名
            full_path = os.path.join(search_path, image_name)
            if os.path.isfile(full_path):
                return os.path.abspath(full_path)
            
            # 如果没有扩展名，尝试添加支持的格式
            if not os.path.splitext(image_name)[1]:
                for ext in supported_formats:
                    full_path_with_ext = os.path.join(search_path, image_name + ext)
                    if os.path.isfile(full_path_with_ext):
                        return os.path.abspath(full_path_with_ext)
        
        return None
    
    def _merge_broken_lines(self, lines: List[str]) -> List[str]:
        """合并被意外分割的行，但保持列表缩进"""
        merged_lines = []
        i = 0
        
        while i < len(lines):
            original_line = lines[i]  # 保持原始缩进
            current_line_stripped = lines[i].strip()
            
            # 如果当前行不为空
            if current_line_stripped:
                # 检查下一行是否是单词片段（通常被意外分割的行）
                if (i + 1 < len(lines) and 
                    lines[i + 1].strip() and
                    not lines[i + 1].strip().startswith('#') and
                    not lines[i + 1].strip().startswith('附件') and
                    len(lines[i + 1].strip()) < 20 and  # 短行更可能是分割的部分
                    not lines[i + 1].strip().endswith('：') and
                    not lines[i + 1].strip().endswith(':') and
                    not lines[i + 1].strip().startswith('|') and  # 不合并表格行
                    not lines[i + 1].strip().startswith('-') and  # 不合并列表项
                    not lines[i + 1].strip().startswith('*') and  # 不合并列表项
                    not re.match(r'^\d+\.', lines[i + 1].strip()) and  # 不合并有序列表
                    not re.match(r'^\s+[-*]', lines[i + 1])):  # 不合并带缩进的列表项
                    
                    # 合并当前行和下一行，保持原始行的缩进
                    merged_line = original_line.rstrip() + lines[i + 1].strip()
                    merged_lines.append(merged_line)
                    i += 2  # 跳过下一行
                else:
                    # 保持原始缩进
                    merged_lines.append(original_line)
                    i += 1
            else:
                # 保持空行
                merged_lines.append(original_line)
                i += 1
        
        return merged_lines
    
    def _skip_first_level_headers(self, lines: List[str]) -> List[str]:
        """跳过一级标题（#），因为我们使用文件名作为文档标题"""
        processed_lines = []
        
        for line in lines:
            # 跳过一级标题，但保留二级及以上标题
            if line.strip().startswith('# ') and not line.strip().startswith('##'):
                continue
            else:
                processed_lines.append(line)
        
        return processed_lines
    
    def _convert_ordered_lists_to_paragraphs(self, lines: List[str]) -> List[str]:
        """将有序列表转换为普通段落，保留序号但阻止pandoc识别为列表"""
        processed_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            line_stripped = line.strip()
            
            # 检测有序列表项 (例: "1. **标题**")
            if re.match(r'^\d+\.\s+', line_stripped):
                # 保留原始缩进，但添加特殊处理防止pandoc识别为列表
                # 在序号后添加全角空格，这样pandoc不会将其识别为列表
                # 获取原始行的缩进
                indent = len(line) - len(line.lstrip())
                modified_line = ' ' * indent + re.sub(r'^(\d+)\.\s+', r'\1.　', line_stripped)
                processed_lines.append(modified_line)
                
                # 检查后续行是否为列表项的续行内容
                j = i + 1
                while j < len(lines) and lines[j].strip() and not re.match(r'^\d+\.\s+', lines[j].strip()) and not lines[j].strip().startswith('#'):
                    # 这是列表项的续行内容，作为单独段落处理
                    continuation_line = lines[j].strip()
                    if continuation_line:  # 非空行
                        processed_lines.append('')  # 添加空行分隔
                        processed_lines.append(continuation_line)
                    j += 1
                
                i = j  # 跳过已处理的行
            else:
                processed_lines.append(line)
                i += 1
        
        return processed_lines
    
    def _fix_unordered_list_asterisks(self, lines: List[str]) -> List[str]:
        """修复无序列表的星号，避免被误识别为斜体"""
        processed_lines = []
        
        for line in lines:
            # 检测无序列表项 (例: "* 项目内容")
            if re.match(r'^\s*\*\s+', line):
                # 在星号前后加空格，确保pandoc正确识别为列表
                processed_line = re.sub(r'^(\s*)\*\s+', r'\1- ', line)
                processed_lines.append(processed_line)
            else:
                processed_lines.append(line)
        
        return processed_lines