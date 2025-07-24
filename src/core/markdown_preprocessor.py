import re
import os
import shutil
from typing import List, Dict, Any
from pathlib import Path
from ..config import DocumentConfig
from ..utils.constants import Patterns, DocumentFormats
from ..utils.exceptions import FileProcessingError, MarkdownParsingError, PathSecurityError

class MarkdownPreprocessor:
    """Markdown预处理器，用于清理和过滤Markdown内容后交给pandoc处理"""
    
    # Caption处理相关常量
    CAPTION_SEARCH_BEFORE = 10  # 向前查找行数
    CAPTION_SEARCH_AFTER = 20   # 向后查找行数
    CAPTION_MAX_EMPTY_LINES = 2  # 最大允许空行数
    
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
        # 注释掉这一行，不再去除有序列表的空格
        # lines = self._remove_numbered_list_spaces(lines)
        lines = self._fix_unordered_list_asterisks(lines)
        lines = self._merge_broken_lines(lines)
        lines = self._skip_first_level_headers(lines)
        lines = self._convert_ordered_lists_to_text(lines)  # 将所有有序列表转换为正文
        
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
            
        # 如果当前行是特殊格式，不合并
        if self._is_special_format_line(current_line):
            return False
            
        # 检查下一行是否为不应合并的特殊格式
        if self._is_special_format_line(next_line):
            return False
            
        # 短行（少于20字符）更可能是被意外分割的部分
        if len(next_line) >= 20:
            return False
            
        return True
    
    def _is_special_format_line(self, line: str) -> bool:
        """检查是否为特殊格式的行（标题、附件、表格等）"""
        # 检查各种不应合并的格式
        special_checks = [
            line.startswith('#'),              # 标题
            line.startswith('附件'),            # 附件
            line.endswith('：') or line.endswith(':'),  # 冒号结尾
            line.startswith('|'),              # 表格行
            line.startswith('-'),              # 列表项（保留以避免合并破坏列表）
            line.startswith('*'),              # 列表项（保留以避免合并破坏列表）
            line == '$$',                      # 数学公式块分隔符
            line.startswith('$$'),             # 数学公式开始/结束行
            line.endswith('$$'),               # 数学公式开始/结束行
            Patterns.CAPTION_PATTERN.match(line),  # 图表caption（使用预编译的模式）
            re.match(r'^\d+\.', line),         # 数字列表项（1. 2. 或 1.内容 等）
        ]
        
        return any(special_checks)
    
    def _is_math_formula_line(self, line: str) -> bool:
        """检查是否为数学公式相关的行"""
        return (line == '$$' or 
                line.startswith('$$') or 
                line.endswith('$$') or
                ('$' in line and line.count('$') % 2 == 0))  # 行内公式
    
    def _skip_first_level_headers(self, lines: List[str]) -> List[str]:
        """动态检测和调整标题层级
        - 如果检测到多个一级标题（#），将所有标题层级下移
        - 如果只有一个或没有一级标题，则跳过一级标题（使用文件名作为文档标题）
        """
        # 统计一级标题数量
        h1_count = 0
        for line in lines:
            if line.strip().startswith('# ') and not line.strip().startswith('##'):
                h1_count += 1
        
        # 如果有多个一级标题，调整所有标题层级
        if h1_count > 1:
            return self._adjust_header_levels(lines)
        else:
            # 原有逻辑：跳过单个一级标题
            processed_lines = []
            for line in lines:
                if line.strip().startswith('# ') and not line.strip().startswith('##'):
                    continue
                else:
                    processed_lines.append(line)
            return processed_lines
    
    def _adjust_header_levels(self, lines: List[str]) -> List[str]:
        """将标题层级下移一级，但只处理到三级标题
        # -> ##（Heading 2）
        ## -> ###（Heading 3） 
        ### -> 作为正文处理（移除标题标记）
        #### 及更深 -> 作为正文处理（移除标题标记）
        """
        processed_lines = []
        
        for line in lines:
            stripped_line = line.strip()
            
            # 检查是否为标题行
            if stripped_line.startswith('#'):
                # 找到第一个空格的位置（标题级别和内容的分隔）
                space_index = stripped_line.find(' ')
                if space_index > 0:
                    # 获取标题级别（#的数量）
                    header_level = stripped_line[:space_index]
                    if header_level and all(c == '#' for c in header_level):
                        level_count = len(header_level)
                        # 获取原始行的缩进
                        indent = line[:len(line) - len(line.lstrip())]
                        
                        if level_count <= 2:
                            # 一级和二级标题：下移一级
                            new_line = indent + '#' + stripped_line
                            processed_lines.append(new_line)
                        else:
                            # 三级及更深的标题：作为正文处理，移除标题标记
                            content = stripped_line[space_index + 1:]  # 获取标题内容
                            # 检查是否是多级编号格式，如果是，需要保护
                            multi_match = re.match(r'^(\d+\.\d+(?:\.\d+)*)\s+(.+)$', content)
                            if multi_match:
                                # 是多级编号，使用反引号包裹
                                numbering = multi_match.group(1)
                                text = multi_match.group(2)
                                content = f"`{numbering}` {text}"
                            new_line = indent + content
                            processed_lines.append(new_line)
                        continue
            
            # 非标题行或无法识别的格式，保持原样
            processed_lines.append(line)
        
        return processed_lines
    
    def _remove_numbered_list_spaces(self, lines: List[str]) -> List[str]:
        """去除简单数字列表项中的空格，如将'1. '转换为'1.'
        但保留多级编号的格式，如 '2.1.1 xxx' 保持不变
        """
        processed_lines = []
        
        for line in lines:
            # 只匹配简单的一级编号（如 "1. "、"2. " 等）
            # 不匹配多级编号（如 "1.1 "、"2.1.1 " 等）
            processed_line = re.sub(r'^(\s*)(\d+)\.\s+(?!\d)', r'\1\2.', line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _fix_unordered_list_asterisks(self, lines: List[str]) -> List[str]:
        """修复无序列表的星号，避免被误识别为斜体"""
        processed_lines = []
        
        for line in lines:
            # 检测无序列表项 (例: "* 项目内容")
            if re.match(r'^(\s*)\*\s+', line):
                # 将星号替换为短横线
                processed_line = re.sub(r'^(\s*)\*\s+', r'\1- ', line)
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
            if not caption_match:
                # 不是标题，直接添加
                processed_lines.append(lines[i])
                i += 1
                continue
            
            caption_type = caption_match.group(1)  # 图/图片/表/表格/图表
            
            # 检查标题是否已经在正确位置
            if self._is_caption_after_element(lines, i, caption_type):
                processed_lines.append(lines[i])
                i += 1
                continue
            
            # 标题不在正确位置，查找对应的图表元素
            element_info = self._find_element_for_caption(lines, i, caption_type)
            
            if element_info['found']:
                # 需要移动标题到元素后面
                caption_line = lines[i]
                i += 1
                
                # 添加从标题后到元素（含元素）的所有行
                while i <= element_info['index']:
                    processed_lines.append(lines[i])
                    i += 1
                
                # 在元素后添加标题
                processed_lines.append(caption_line)
            else:
                # 没找到对应元素，保持原位置
                processed_lines.append(lines[i])
                i += 1
        
        return processed_lines
    
    def _is_caption_after_element(self, lines: List[str], caption_index: int, caption_type: str) -> bool:
        """检查caption是否已在正确位置（紧跟在对应图表后面）"""
        empty_lines = 0
        
        # 向前查找，最多查找CAPTION_SEARCH_BEFORE行
        for j in range(caption_index - 1, max(caption_index - self.CAPTION_SEARCH_BEFORE, -1), -1):
            prev_line = lines[j].strip()
            
            if not prev_line:  # 空行
                empty_lines += 1
                continue
            
            # 检查是否为匹配的元素
            if self._is_matching_element(prev_line, caption_type):
                # 如果是表格，需要确认是表格的最后一行
                if caption_type in ['表', '表格'] and j + 1 < len(lines):
                    next_line = lines[j + 1].strip()
                    if next_line and Patterns.TABLE_ROW_PATTERN.match(next_line):
                        return False  # 不是表格最后一行
                
                # 检查空行数是否在允许范围内
                return empty_lines <= self.CAPTION_MAX_EMPTY_LINES
            else:
                # 遇到其他内容，停止查找
                break
        
        return False
    
    def _find_element_for_caption(self, lines: List[str], caption_index: int, caption_type: str) -> Dict[str, Any]:
        """向后查找caption对应的图表元素"""
        # 从下一行开始，最多查找CAPTION_SEARCH_AFTER行
        for j in range(caption_index + 1, min(caption_index + self.CAPTION_SEARCH_AFTER + 1, len(lines))):
            check_line = lines[j].strip()
            
            # 如果遇到另一个标题，停止查找
            if Patterns.CAPTION_PATTERN.match(check_line):
                break
            
            # 检查是否为匹配的元素
            if self._is_matching_element(check_line, caption_type):
                element_index = j
                
                # 如果是表格，找到表格的结束位置
                if caption_type in ['表', '表格']:
                    element_index = self._find_table_end(lines, j)
                
                return {'found': True, 'index': element_index}
        
        return {'found': False, 'index': -1}
    
    def _is_matching_element(self, line: str, caption_type: str) -> bool:
        """判断是否为匹配的图表元素"""
        if caption_type in ['图', '图片', '图表']:
            return (Patterns.MARKDOWN_IMAGE_PATTERN.match(line) or 
                    Patterns.OBSIDIAN_IMAGE_PATTERN.match(line))
        elif caption_type in ['表', '表格']:
            return Patterns.TABLE_ROW_PATTERN.match(line)
        return False
    
    def _find_table_end(self, lines: List[str], table_start: int) -> int:
        """找到表格的结束位置"""
        end_index = table_start
        
        # 确认是表格：检查是否有连续的表格行
        if table_start + 1 < len(lines) and Patterns.TABLE_ROW_PATTERN.match(lines[table_start + 1].strip()):
            # 继续向后查找直到表格结束
            while end_index + 1 < len(lines) and Patterns.TABLE_ROW_PATTERN.match(lines[end_index + 1].strip()):
                end_index += 1
        
        return end_index
    
    def _convert_ordered_lists_to_text(self, lines: List[str]) -> List[str]:
        """将所有有序列表和多级编号转换为正文格式
        
        规则：
        1. 所有形如 '1. 内容' 的有序列表都转换为正文
        2. 所有形如 '2.1.1 内容' 的多级编号也转换为正文
        3. 使用特殊标记包裹编号，完全阻止 Pandoc 识别
        """
        processed_lines = []
        
        for line in lines:
            # 先匹配多级编号格式（如 2.1.1 内容）
            multi_match = re.match(r'^(\s*)(\d+\.\d+(?:\.\d+)*)\s+(.+)$', line)
            if multi_match:
                indent = multi_match.group(1)
                numbering = multi_match.group(2)
                content = multi_match.group(3)
                # 使用反引号（inline code）包裹整个编号，阻止 Pandoc 识别为列表
                new_line = f"{indent}`{numbering}` {content}"
                processed_lines.append(new_line)
            else:
                # 再匹配简单有序列表格式（数字 + 点 + 空格）
                simple_match = re.match(r'^(\s*)(\d+)\.\s+(.+)$', line)
                if simple_match:
                    indent = simple_match.group(1)
                    number = simple_match.group(2)
                    content = simple_match.group(3)
                    # 使用反引号包裹
                    new_line = f"{indent}`{number}.` {content}"
                    processed_lines.append(new_line)
                else:
                    processed_lines.append(line)
        
        return processed_lines
    
