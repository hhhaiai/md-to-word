import re
import os
import shutil
from typing import List, Dict, Any
from pathlib import Path
from config import DocumentConfig
from constants import Patterns, DocumentFormats
from exceptions import FileProcessingError, MarkdownParsingError, PathSecurityError

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
        lines = self._reposition_captions(lines)  # 在合并行之前重新定位标题
        lines = self._remove_numbered_list_spaces(lines)
        lines = self._convert_ordered_lists_to_paragraphs(lines)
        lines = self._fix_unordered_list_asterisks(lines)
        lines = self._process_math_formulas(lines)
        # 不处理图片，保留原始语法
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
            # 使用预编译的正则表达式，提高性能
            processed_line = Patterns.BOLD_PATTERN.sub(r'\1', line)
            processed_line = Patterns.BOLD_UNDERSCORE_PATTERN.sub(r'\1', processed_line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _remove_numbered_list_spaces(self, lines: List[str]) -> List[str]:
        """去除数字列表项中的空格，如将'1. '转换为'1.'"""
        processed_lines = []
        
        for line in lines:
            # 使用预编译的正则表达式
            processed_line = Patterns.NUMBERED_LIST_PATTERN.sub(r'\1\2.', line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _process_math_formulas(self, lines: List[str]) -> List[str]:
        """处理数学公式 - 保持LaTeX格式交给pandoc处理"""
        # 保持原始LaTeX格式，由pandoc负责MathML转换
        return lines
    
    
    def _find_image_path(self, image_name: str, source_dir: str) -> str:
        """在配置的搜索路径中查找图片"""
        
        # 构建搜索路径列表
        search_paths = [
            # 首先在源文件目录查找
            source_dir,
            # 然后在配置的搜索路径中查找（动态获取）
            *DocumentConfig.get_image_search_paths()
        ]
        
        # 支持的图片格式
        supported_formats = self.image_config['supported_formats']
        
        for search_path_str in search_paths:
            search_path = Path(search_path_str).resolve()
            if not search_path.exists() or not search_path.is_dir():
                continue
            
            # 安全的路径构建和验证
            try:
                # 构建目标图片路径
                image_path = (search_path / image_name).resolve()
                
                # 关键安全检查：确保解析后的路径在允许的搜索目录内
                if not self._is_safe_path(image_path, search_path):
                    continue
                
                # 直接匹配文件名
                if image_path.is_file():
                    return str(image_path)
                
                # 如果没有扩展名，尝试添加支持的格式
                if not image_path.suffix:
                    for ext in supported_formats:
                        path_with_ext = image_path.with_suffix(ext)
                        if path_with_ext.is_file() and self._is_safe_path(path_with_ext, search_path):
                            return str(path_with_ext)
                            
            except (OSError, ValueError) as e:
                # 处理无效路径或文件系统错误
                continue
        
        return None
    
    def _is_safe_path(self, target_path: Path, base_path: Path) -> bool:
        """
        验证目标路径是否在基础路径范围内，防止路径遍历攻击
        
        Args:
            target_path: 目标文件的绝对路径
            base_path: 允许的基础目录路径
            
        Returns:
            bool: 如果路径安全则返回True，否则返回False
        """
        try:
            # 确保两个路径都是绝对路径
            target_resolved = target_path.resolve()
            base_resolved = base_path.resolve()
            
            # 检查目标路径是否在基础路径内（包括基础路径本身）
            return base_resolved == target_resolved or base_resolved in target_resolved.parents
            
        except (OSError, ValueError):
            # 如果路径解析失败，认为不安全
            return False
    
    def _merge_broken_lines(self, lines: List[str]) -> List[str]:
        """合并被意外分割的行，但保持列表缩进"""
        merged_lines = []
        i = 0
        
        while i < len(lines):
            current_line = lines[i]
            
            if not current_line.strip():
                # 保持空行
                merged_lines.append(current_line)
                i += 1
                continue
            
            # 检查是否可以与下一行合并
            if self._should_merge_with_next_line(lines, i):
                # 合并当前行和下一行，保持原始行的缩进
                merged_line = current_line.rstrip() + lines[i + 1].strip()
                merged_lines.append(merged_line)
                i += 2  # 跳过下一行
            else:
                # 保持原始缩进
                merged_lines.append(current_line)
                i += 1
        
        return merged_lines
    
    def _should_merge_with_next_line(self, lines: List[str], current_index: int) -> bool:
        """判断当前行是否应该与下一行合并"""
        if current_index + 1 >= len(lines):
            return False
            
        current_line = lines[current_index].strip()
        next_line = lines[current_index + 1].strip()
        
        # 下一行为空则不合并
        if not next_line:
            return False
        
        # 如果当前行是数学公式相关，不合并
        if self._is_math_formula_line(current_line):
            return False
            
        # 检查下一行是否为不应合并的特殊格式
        if self._is_special_format_line(next_line):
            return False
            
        # 短行（少于20字符）更可能是被意外分割的部分
        if len(next_line) >= 20:
            return False
            
        return True
    
    def _is_special_format_line(self, line: str) -> bool:
        """检查是否为特殊格式的行（标题、附件、表格、列表等）"""
        # 检查各种不应合并的格式
        special_checks = [
            line.startswith('#'),              # 标题
            line.startswith('附件'),            # 附件
            line.endswith('：') or line.endswith(':'),  # 冒号结尾
            line.startswith('|'),              # 表格行
            line.startswith('-'),              # 列表项
            line.startswith('*'),              # 列表项
            line == '$$',                      # 数学公式块分隔符
            line.startswith('$$'),             # 数学公式开始/结束行
            line.endswith('$$'),               # 数学公式开始/结束行
            Patterns.CAPTION_PATTERN.match(line),  # 图表caption（使用预编译的模式）
            Patterns.ORDERED_LIST_PATTERN_PREPROCESSOR.match(line),     # 有序列表 (使用预处理器版本)
            Patterns.INDENTED_LIST_PATTERN.match(line),    # 带缩进的列表项
        ]
        
        return any(special_checks)
    
    def _is_math_formula_line(self, line: str) -> bool:
        """检查是否为数学公式相关的行"""
        return (line == '$$' or 
                line.startswith('$$') or 
                line.endswith('$$') or
                ('$' in line and line.count('$') % 2 == 0))  # 行内公式
    
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
            if Patterns.ORDERED_LIST_PATTERN_PREPROCESSOR.match(line_stripped):
                # 保留原始缩进，但添加特殊处理防止pandoc识别为列表
                # 在序号后添加全角空格，这样pandoc不会将其识别为列表
                # 获取原始行的缩进
                indent = len(line) - len(line.lstrip())
                modified_line = ' ' * indent + Patterns.ORDERED_LIST_DOT_REPLACE_PATTERN.sub(r'\1.　', line_stripped)
                processed_lines.append(modified_line)
                
                # 检查后续行是否为列表项的续行内容
                j = i + 1
                while j < len(lines) and lines[j].strip() and not Patterns.ORDERED_LIST_PATTERN_PREPROCESSOR.match(lines[j].strip()) and not lines[j].strip().startswith('#'):
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
            if Patterns.UNORDERED_LIST_PATTERN.match(line):
                # 在星号前后加空格，确保pandoc正确识别为列表
                processed_line = Patterns.UNORDERED_LIST_REPLACE_PATTERN.sub(r'\1- ', line)
                processed_lines.append(processed_line)
            else:
                processed_lines.append(line)
        
        return processed_lines
    
    def _reposition_captions(self, lines: List[str]) -> List[str]:
        """重新定位图表标题，确保标题始终在图表后面"""
        processed_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i].strip()
            
            # 检查是否为图表标题
            caption_match = Patterns.CAPTION_PATTERN.match(line)
            if caption_match:
                caption_type = caption_match.group(1)  # 图/图片/表/表格/图表
                caption_number = caption_match.group(2)  # 数字
                
                # 判断标题类型
                is_figure_caption = caption_type in ['图', '图片', '图表']
                is_table_caption = caption_type in ['表', '表格']
                
                # 首先检查标题是否已经在正确位置（紧跟在图表后面）
                # 向前查找最近的图表元素
                is_after_element = False
                empty_lines_before = 0
                
                for j in range(i - 1, max(i - 10, -1), -1):  # 最多向前查找10行
                    prev_line = lines[j].strip()
                    
                    if not prev_line:  # 空行
                        empty_lines_before += 1
                        continue
                    
                    # 检查是否为对应类型的元素
                    if is_figure_caption and (Patterns.MARKDOWN_IMAGE_PATTERN.match(prev_line) or 
                                            Patterns.OBSIDIAN_IMAGE_PATTERN.match(prev_line)):
                        # 找到了前面的图片，如果中间空行不超过2行，认为caption属于这个图片
                        if empty_lines_before <= 2:
                            is_after_element = True
                        break
                    elif is_table_caption and Patterns.TABLE_ROW_PATTERN.match(prev_line):
                        # 检查是否是完整表格的最后一行
                        # 继续向前查找，确认这是表格的一部分
                        is_table_end = True
                        if j + 1 < len(lines):
                            next_line = lines[j + 1].strip()
                            # 如果下一行也是表格行，说明不是表格结束
                            if next_line and Patterns.TABLE_ROW_PATTERN.match(next_line):
                                is_table_end = False
                        
                        if is_table_end and empty_lines_before <= 2:
                            is_after_element = True
                        break
                    else:
                        # 遇到其他内容，停止查找
                        break
                
                # 如果caption已经在正确位置（某个图表元素后面），保持不动
                if is_after_element:
                    processed_lines.append(lines[i])
                    i += 1
                    continue
                
                # 标题不在正确位置，向后查找对应的图表元素
                found_element = False
                element_index = -1
                
                # 查找范围：从下一行开始，最多查找20行
                for j in range(i + 1, min(i + 21, len(lines))):
                    check_line = lines[j].strip()
                    
                    # 如果遇到另一个标题，停止查找
                    if Patterns.CAPTION_PATTERN.match(check_line):
                        break
                    
                    # 检查是否为图片
                    if is_figure_caption and (Patterns.MARKDOWN_IMAGE_PATTERN.match(check_line) or 
                                            Patterns.OBSIDIAN_IMAGE_PATTERN.match(check_line)):
                        found_element = True
                        element_index = j
                        break
                    
                    # 检查是否为表格（连续的表格行）
                    if is_table_caption and Patterns.TABLE_ROW_PATTERN.match(check_line):
                        # 确认是表格：检查是否有连续的表格行
                        if j + 1 < len(lines) and Patterns.TABLE_ROW_PATTERN.match(lines[j + 1].strip()):
                            found_element = True
                            element_index = j
                            # 找到表格的结束位置
                            while element_index + 1 < len(lines) and Patterns.TABLE_ROW_PATTERN.match(lines[element_index + 1].strip()):
                                element_index += 1
                            break
                
                if found_element:
                    # 标题在元素前面，需要移动到元素后面
                    # 保存原始标题行（包含缩进）
                    caption_line = lines[i]
                    
                    # 跳过当前标题行
                    i += 1
                    
                    # 添加从标题后到元素（含元素）的所有行
                    while i <= element_index:
                        processed_lines.append(lines[i])
                        i += 1
                    
                    # 在元素后添加标题
                    processed_lines.append(caption_line)
                    continue
            
            # 不是需要移动的标题，或标题已经在正确位置，直接添加
            processed_lines.append(lines[i])
            i += 1
        
        return processed_lines