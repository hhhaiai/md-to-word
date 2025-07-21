# Markdown到Word公文格式转换工具

这是一个Python工具，用于将Markdown文件转换为符合中国国家标准GB/T 9704-2012的Word公文格式。

## 功能特点

- ✅ 符合GB/T 9704-2012《党政机关公文格式》国家标准
- ✅ 自动设置正确的页边距、字体、字号和行距
- ✅ 支持标题层级转换
- ✅ 自动过滤YAML front matter和结尾元数据
- ✅ 去除加粗标记，保持格式统一
- ✅ 处理数学公式（LaTeX格式转纯文本）
- ✅ 自动识别成文日期并转换为汉字格式
- ✅ 支持附件说明
- ✅ 自动添加页码

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python3 md_to_word.py document.md
```

生成与输入文件同名的`.docx`文件。

### 指定输出文件

```bash
python3 md_to_word.py input.md -o output.docx
```

### 转换任意路径文件

```bash
python3 md_to_word.py /path/to/document.md
python3 md_to_word.py ~/Desktop/报告.md
```

### 查看帮助

```bash
python3 md_to_word.py --help
```

## Markdown格式要求

### 标题层级对应关系

- `#` → 被忽略（使用文件名作为文档标题）
- `##` → 一级标题（黑体，三号，不加粗）
- `###` → 二级标题（Kaiti TC，三号，不加粗）
- `####` → 三级标题（仿宋，三号，不加粗）

### 自动处理功能

1. **YAML Front Matter**：自动过滤`---`包围的元数据
2. **结尾元数据**：自动过滤Date行、标签（如`#work`）
3. **加粗格式**：去除`**文字**`和`__文字__`标记
4. **数学公式**：`$10^{18}$` → `10^18`
5. **成文日期**：`2025年7月21日` → `二〇二五年七月二十一日`

### 基本结构示例

```markdown
---
项目: 示例
类别: Reports
---

# 文档标题（会被忽略，使用文件名）

## 一、项目概述

正文内容...

### （一）项目背景

正文内容...

**加粗文字**会被去除格式。

数学公式如$10^{18}$会转换为纯文本。

附件：1. 附件名称
附件：2. 附件名称

2025年7月21日

---
Date: 2025-07-19
#work
```

## 格式标准

本工具严格按照GB/T 9704-2012标准设置格式：

- **纸张**：A4 (210mm × 297mm)
- **页边距**：上37mm、下35mm、左28mm、右26mm
- **字体**：
  - 文档标题：小标宋体、二号、居中
  - 正文：FangSong、三号
  - 一级标题：黑体、三号、不加粗
  - 二级标题：Kaiti TC、三号、不加粗
  - 三级标题：仿宋、三号、不加粗
- **行距**：固定值28.8磅
- **对齐**：
  - 文档标题：居中
  - 其他所有文字：首行缩进2字符，两端对齐
- **页码**：仿宋、四号、居中（格式：`- 1 -`）

## 项目结构

```
md-to-word/
├── md_to_word.py          # 主程序
├── requirements.txt       # 依赖列表
├── config.py             # 格式配置
├── markdown_parser.py    # Markdown解析器
├── word_generator.py     # Word生成器
├── README.md             # 使用说明
├── CLAUDE.md             # 项目状态和问题记录
└── examples/            # 示例文件
    ├── input.md         # 示例输入
    └── *.docx          # 转换结果
```

## 支持的输入格式

- 标准Markdown文件
- 带YAML front matter的Markdown
- Obsidian笔记格式
- 包含LaTeX数学公式的Markdown

## 注意事项

1. 确保系统中已安装相应的中文字体
2. 工具不会自动识别主送机关，如需要请手动处理
3. 生成的Word文档已设置为符合公文标准的格式