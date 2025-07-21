import re
import markdown
from typing import List, Dict, Any

class MarkdownParser:
    """Markdown解析器，专门用于公文格式转换"""
    
    def __init__(self):
        self.md = markdown.Markdown(extensions=[
            'markdown.extensions.tables',
            'markdown.extensions.fenced_code',
        ])
    
    def parse_file(self, file_path: str) -> Dict[str, Any]:
        """解析Markdown文件"""
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 获取文件名作为标题（去掉扩展名）
        import os
        filename = os.path.basename(file_path)
        title_from_filename = os.path.splitext(filename)[0]
        
        result = self.parse_content(content)
        
        # 始终使用文件名作为文档标题
        result['title'] = title_from_filename
        
        return result
    
    def parse_content(self, content: str) -> Dict[str, Any]:
        """解析Markdown内容"""
        lines = content.split('\n')
        
        # 过滤YAML front matter
        lines = self._filter_yaml_frontmatter(lines)
        
        # 过滤结尾的Date和标签
        lines = self._filter_ending_metadata(lines)
        
        # 去除加粗标记
        lines = self._remove_bold_formatting(lines)
        
        # 处理数学公式
        lines = self._process_math_formulas(lines)
        
        # 合并被意外分割的行
        lines = self._merge_broken_lines(lines)
        result = {
            'title': '',
            'subtitle': '',
            'sender': '',       # 主送机关
            'body': [],         # 正文段落
            'date': '',         # 成文日期
            'attachments': []   # 附件说明
        }
        
        current_section = 'body'
        current_paragraph = []
        
        for line in lines:
            line = line.strip()
            
            if not line:
                if current_paragraph:
                    result[current_section].append({
                        'type': 'paragraph',
                        'content': '\n'.join(current_paragraph),
                        'level': 0
                    })
                    current_paragraph = []
                continue
            
            # 检测标题
            if line.startswith('#'):
                if current_paragraph:
                    result[current_section].append({
                        'type': 'paragraph',
                        'content': '\n'.join(current_paragraph),
                        'level': 0
                    })
                    current_paragraph = []
                
                level = len(line) - len(line.lstrip('#'))
                title_text = line.lstrip('#').strip()
                
                if level == 1:
                    # #标题忽略，因为已经使用文件名作为文档标题
                    pass
                else:
                    # ##作为一级标题，###作为二级标题，####作为三级标题
                    result['body'].append({
                        'type': 'heading',
                        'content': title_text,
                        'level': level - 1  # 减1是因为##对应一级标题
                    })
            
            # 检测日期（年月日格式）
            elif re.match(r'.*\d{4}年.*\d{1,2}月.*\d{1,2}日.*', line):
                result['date'] = self._convert_date_to_chinese(line)
            
            # 检测附件说明
            elif line.startswith('附件') or '附件' in line:
                result['attachments'].append(line)
            
            # 普通文本
            else:
                current_paragraph.append(line)
        
        # 处理最后一段
        if current_paragraph:
            result[current_section].append({
                'type': 'paragraph',
                'content': '\n'.join(current_paragraph),
                'level': 0
            })
        
        return result
    
    def _convert_date_to_chinese(self, date_str: str) -> str:
        """将阿拉伯数字日期转换为汉字数字日期"""
        # 数字到汉字的映射
        num_map = {
            '0': '〇', '1': '一', '2': '二', '3': '三', '4': '四',
            '5': '五', '6': '六', '7': '七', '8': '八', '9': '九'
        }
        
        # 提取年月日
        year_match = re.search(r'(\d{4})年', date_str)
        month_match = re.search(r'(\d{1,2})月', date_str)
        day_match = re.search(r'(\d{1,2})日', date_str)
        
        if year_match and month_match and day_match:
            year = year_match.group(1)
            month = month_match.group(1)
            day = day_match.group(1)
            
            # 转换年份
            chinese_year = ''.join([num_map[d] for d in year])
            
            # 转换月份
            if len(month) == 1:
                chinese_month = num_map[month]
            else:
                if month[0] == '1':
                    chinese_month = '十' + (num_map[month[1]] if month[1] != '0' else '')
                else:
                    chinese_month = num_map[month[0]] + '十' + (num_map[month[1]] if month[1] != '0' else '')
            
            # 转换日期
            if len(day) == 1:
                chinese_day = num_map[day]
            else:
                if day[0] == '1':
                    chinese_day = '十' + (num_map[day[1]] if day[1] != '0' else '')
                elif day[0] == '2':
                    chinese_day = '二十' + (num_map[day[1]] if day[1] != '0' else '')
                elif day[0] == '3':
                    chinese_day = '三十' + (num_map[day[1]] if day[1] != '0' else '')
                else:
                    chinese_day = num_map[day[0]] + '十' + (num_map[day[1]] if day[1] != '0' else '')
            
            return f"{chinese_year}年{chinese_month}月{chinese_day}日"
        
        return date_str
    
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
    
    def _process_math_formulas(self, lines: List[str]) -> List[str]:
        """处理数学公式，将LaTeX格式转换为可读文本"""
        processed_lines = []
        
        for line in lines:
            # 处理行内数学公式 $formula$
            processed_line = re.sub(r'\$([^$]+)\$', self._convert_math_formula, line)
            processed_lines.append(processed_line)
        
        return processed_lines
    
    def _convert_math_formula(self, match) -> str:
        """转换数学公式为纯文本"""
        formula = match.group(1)
        
        # 处理上标 ^{number} 或 ^number
        formula = re.sub(r'\^\{([^}]+)\}', r'^\1', formula)
        formula = re.sub(r'\^([0-9]+)', r'^\1', formula)
        
        # 处理下标 _{number} 或 _number  
        formula = re.sub(r'_\{([^}]+)\}', r'_\1', formula)
        formula = re.sub(r'_([0-9]+)', r'_\1', formula)
        
        # 处理常见的希腊字母和符号
        replacements = {
            '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
            '\\epsilon': 'ε', '\\theta': 'θ', '\\lambda': 'λ', '\\mu': 'μ',
            '\\pi': 'π', '\\sigma': 'σ', '\\phi': 'φ', '\\omega': 'ω',
            '\\times': '×', '\\cdot': '·', '\\approx': '≈', '\\leq': '≤',
            '\\geq': '≥', '\\neq': '≠', '\\pm': '±', '\\infty': '∞'
        }
        
        for latex, unicode_char in replacements.items():
            formula = formula.replace(latex, unicode_char)
        
        return formula
    
    def _merge_broken_lines(self, lines: List[str]) -> List[str]:
        """合并被意外分割的行"""
        merged_lines = []
        i = 0
        
        while i < len(lines):
            current_line = lines[i].strip()
            
            # 如果当前行不为空
            if current_line:
                # 检查下一行是否是单词片段（通常被意外分割的行）
                if (i + 1 < len(lines) and 
                    lines[i + 1].strip() and
                    not lines[i + 1].strip().startswith('#') and
                    not lines[i + 1].strip().startswith('附件') and
                    len(lines[i + 1].strip()) < 20 and  # 短行更可能是分割的部分
                    not lines[i + 1].strip().endswith('：') and
                    not lines[i + 1].strip().endswith(':')):
                    
                    # 合并当前行和下一行
                    merged_line = current_line + lines[i + 1].strip()
                    merged_lines.append(merged_line)
                    i += 2  # 跳过下一行
                else:
                    merged_lines.append(current_line)
                    i += 1
            else:
                merged_lines.append(current_line)
                i += 1
        
        return merged_lines