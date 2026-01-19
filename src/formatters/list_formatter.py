"""
列表格式化器 - 负责有序和无序列表的格式化

缩进规则：
- 第一级：2字符缩进（32pt）
- 每嵌套一级：+2字符（+32pt）

Bullet规则：
- 默认使用实心圆（●）
- 只有当父级是实心圆时，子级使用空心圆（○）
"""
from docx import Document
from docx.shared import Pt
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from .base_formatter import BaseFormatter
from ..utils.constants import Patterns
from ..utils.xpath_cache import OptimizedXMLProcessor


# 每个字符的pt值（三号字16pt）
CHAR_WIDTH_PT = 16


class ListFormatter(BaseFormatter):
    """列表格式化器 - 负责有序和无序列表的格式化"""

    def __init__(self, config=None):
        super().__init__(config)
        self.xml_processor = OptimizedXMLProcessor()

    def format_lists(self, doc: Document):
        """格式化列表 - 统一缩进和字体"""
        # 先修复numbering定义中的bullet字符和缩进
        self._fix_numbering_definitions(doc)

        for paragraph in doc.paragraphs:
            # 跳过标题段落
            if paragraph.style.name.startswith('Heading'):
                continue

            # 处理Word内置列表
            if self._is_word_list_item(paragraph):
                self._format_list_item(paragraph)

    def _is_word_list_item(self, paragraph) -> bool:
        """判断是否为Word内置列表项"""
        if paragraph.style.name == 'Compact':
            return True

        if hasattr(paragraph, '_element'):
            numPr = self.xml_processor.cache.find_first(paragraph._element, './/w:numPr')
            return numPr is not None

        return False

    def _get_list_level(self, paragraph) -> int:
        """获取列表项的级别（0-based）"""
        if hasattr(paragraph, '_element'):
            ilvl_elem = self.xml_processor.cache.find_first(paragraph._element, './/w:ilvl')
            if ilvl_elem is not None:
                val = ilvl_elem.get(qn('w:val'))
                if val:
                    return int(val)
        return 0

    def _format_list_item(self, paragraph):
        """格式化列表项"""
        level = self._get_list_level(paragraph)

        # 检查是否有 [NESTED] 标记（表示原本是嵌套列表）
        is_nested = self._check_and_remove_nested_marker(paragraph)

        # 应用缩进：嵌套项4字符，普通项2字符
        self._apply_level_indent(paragraph, level, is_nested)

        # 设置字体格式
        for run in paragraph.runs:
            run.font.name = self.config.FONTS['fangsong']
            run.font.size = self.config.FONT_SIZES['body']
            self._set_chinese_font(run, self.config.FONTS['fangsong'])

        # 设置列表编号/bullet的字体
        self._format_list_number_font(paragraph)

    def _check_and_remove_nested_marker(self, paragraph) -> bool:
        """检查并移除 [NESTED] 标记，返回是否是嵌套项"""
        # 检查整个段落文本
        full_text = paragraph.text
        if '[NESTED]' not in full_text:
            return False

        # 遍历所有run移除标记
        for run in paragraph.runs:
            if '[NESTED]' in run.text:
                run.text = run.text.replace('[NESTED]', '')

        # 如果标记跨越多个run，需要更复杂的处理
        # 重新检查是否还有残留
        if '[NESTED]' in paragraph.text:
            # 标记可能被拆分到多个run中，尝试重建文本
            self._remove_marker_across_runs(paragraph)

        return True

    def _remove_marker_across_runs(self, paragraph):
        """处理标记被拆分到多个run的情况"""
        marker = '[NESTED]'
        runs = paragraph.runs
        if not runs:
            return

        # 收集所有run的文本和位置信息
        full_text = ''.join(run.text for run in runs)
        marker_start = full_text.find(marker)

        if marker_start == -1:
            return

        marker_end = marker_start + len(marker)

        # 找到标记所在的run范围并清除
        current_pos = 0
        for run in runs:
            run_start = current_pos
            run_end = current_pos + len(run.text)

            if run_start < marker_end and run_end > marker_start:
                # 这个run包含标记的一部分
                text = run.text
                # 计算需要移除的部分
                remove_start = max(0, marker_start - run_start)
                remove_end = min(len(text), marker_end - run_start)
                run.text = text[:remove_start] + text[remove_end:]

            current_pos = run_end

    def _apply_level_indent(self, paragraph, level: int, is_nested: bool = False):
        """应用基于级别的缩进"""
        paragraph_format = paragraph.paragraph_format

        if is_nested:
            # 原本嵌套的列表项：4字符缩进
            paragraph_format.left_indent = Pt(CHAR_WIDTH_PT * 4)
            paragraph_format.first_line_indent = Pt(-CHAR_WIDTH_PT)
        elif level == 0:
            # 普通第一级：2字符缩进
            paragraph_format.left_indent = Pt(CHAR_WIDTH_PT * 2)
            paragraph_format.first_line_indent = Pt(-CHAR_WIDTH_PT)
        else:
            # 其他嵌套级别：每级+2字符
            base_indent = (level + 1) * 2 * CHAR_WIDTH_PT
            paragraph_format.left_indent = Pt(base_indent)
            paragraph_format.first_line_indent = Pt(-CHAR_WIDTH_PT)

        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph.alignment = self.config.ALIGNMENTS['justify']

        # 启用文档网格对齐
        self._enable_snap_to_grid(paragraph)

    def _format_list_number_font(self, paragraph):
        """设置列表编号/bullet的字体为仿宋"""
        if not hasattr(paragraph, '_element'):
            return

        pPr = paragraph._element.find(qn('w:pPr'))
        if pPr is None:
            return

        numPr = pPr.find(qn('w:numPr'))
        if numPr is None:
            return

        # 查找或创建rPr元素
        rPr = numPr.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            numPr.append(rPr)

        # 设置字体
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)

        font_name = self.config.FONTS['fangsong']
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:eastAsia'), font_name)

        # 设置字号
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            rPr.append(sz)
        sz.set(qn('w:val'), str(int(self.config.FONT_SIZES['body'].pt * 2)))

        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = OxmlElement('w:szCs')
            rPr.append(szCs)
        szCs.set(qn('w:val'), str(int(self.config.FONT_SIZES['body'].pt * 2)))

    def _fix_numbering_definitions(self, doc: Document):
        """修复numbering定义：设置正确的bullet字符和缩进"""
        if not hasattr(doc, 'part') or not hasattr(doc.part, 'numbering_part'):
            return

        numbering_part = doc.part.numbering_part
        if numbering_part is None:
            return

        numbering_elm = numbering_part._element

        # 遍历所有abstractNum定义
        for abstractNum in numbering_elm.findall(qn('w:abstractNum')):
            self._fix_abstract_num(abstractNum)

    def _fix_abstract_num(self, abstractNum):
        """修复单个abstractNum定义"""
        levels = abstractNum.findall(qn('w:lvl'))

        # 先收集所有级别的类型和bullet状态
        level_info = {}  # ilvl -> {'type': 'bullet'/'decimal', 'is_solid_bullet': bool}

        for lvl in levels:
            ilvl = int(lvl.get(qn('w:ilvl'), 0))
            numFmt = lvl.find(qn('w:numFmt'))
            fmt_val = numFmt.get(qn('w:val')) if numFmt is not None else None
            level_info[ilvl] = {'type': fmt_val, 'is_solid_bullet': False}

        # 按级别顺序处理，确定每个bullet的样式
        for lvl in sorted(levels, key=lambda x: int(x.get(qn('w:ilvl'), 0))):
            ilvl = int(lvl.get(qn('w:ilvl'), 0))
            numFmt = lvl.find(qn('w:numFmt'))

            if numFmt is not None and numFmt.get(qn('w:val')) == 'bullet':
                lvlText = lvl.find(qn('w:lvlText'))
                if lvlText is not None:
                    # 查找最近的父级bullet是否是实心
                    parent_is_solid_bullet = False
                    for parent_lvl in range(ilvl - 1, -1, -1):
                        if parent_lvl in level_info:
                            parent_info = level_info[parent_lvl]
                            if parent_info['type'] == 'bullet' and parent_info['is_solid_bullet']:
                                parent_is_solid_bullet = True
                                break
                            elif parent_info['type'] == 'bullet':
                                # 父级是bullet但不是实心，继续向上查找
                                continue
                            else:
                                # 父级是数字，停止查找
                                break

                    # 设置bullet字符
                    if parent_is_solid_bullet:
                        # 父级是实心bullet，这级用空心
                        lvlText.set(qn('w:val'), '○')
                        level_info[ilvl]['is_solid_bullet'] = False
                    else:
                        # 父级不是实心bullet（可能是数字或空心），这级用实心
                        lvlText.set(qn('w:val'), '●')
                        level_info[ilvl]['is_solid_bullet'] = True

                # 设置bullet字符的字号为较小尺寸
                self._set_bullet_font_size(lvl)

            # 修复缩进
            self._fix_level_indent(lvl, ilvl)

    def _set_bullet_font_size(self, lvl):
        """设置bullet字符的字号为较小尺寸（10pt）"""
        # 查找或创建rPr元素
        rPr = lvl.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            lvl.append(rPr)

        # 设置字号为10pt（正文16pt的约60%）
        bullet_size_pt = 10
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            rPr.append(sz)
        sz.set(qn('w:val'), str(bullet_size_pt * 2))  # Word uses half-points

        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = OxmlElement('w:szCs')
            rPr.append(szCs)
        szCs.set(qn('w:val'), str(bullet_size_pt * 2))

    def _fix_level_indent(self, lvl, ilvl: int):
        """修复级别定义中的缩进设置"""
        if ilvl == 0:
            # 第一级：无缩进，只有悬挂缩进容纳编号
            indent_pt = CHAR_WIDTH_PT * 2
            hanging_pt = CHAR_WIDTH_PT * 2
        else:
            # 嵌套级别：level=1 -> 2字符, level=2 -> 4字符, etc.
            indent_pt = ilvl * 2 * CHAR_WIDTH_PT + CHAR_WIDTH_PT
            hanging_pt = CHAR_WIDTH_PT

        # 转换为twips (1pt = 20twips)
        indent_twips = int(indent_pt * 20)
        hanging_twips = int(hanging_pt * 20)

        # 查找或创建pPr元素
        pPr = lvl.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            lvl.append(pPr)

        # 查找或创建ind元素
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)

        # 设置缩进
        ind.set(qn('w:left'), str(indent_twips))
        ind.set(qn('w:hanging'), str(hanging_twips))
