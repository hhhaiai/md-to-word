# Markdown到Word公文格式转换工具

这是一个基于Pandoc的Python工具，用于将Markdown文件转换为符合中国国家标准GB/T 9704-2012的Word公文格式。支持LaTeX数学公式、表格和多级列表。

## 功能特点

- ✅ 符合GB/T 9704-2012《党政机关公文格式》国家标准
- ✅ **基于Pandoc引擎**，转换效果更专业
- ✅ **完整支持LaTeX数学公式**（使用MathML渲染）
- ✅ **原生支持Markdown表格**转Word表格
- ✅ **多级列表支持**（有序列表、无序列表）
- ✅ **图片导入功能**（支持Obsidian附件目录，智能路径解析）
- ✅ **图片格式化**（文字环绕、水平居中、智能caption处理）
- ✅ 自动设置正确的页边距、字体、字号和行距
- ✅ 支持标题层级转换
- ✅ 自动过滤YAML front matter和结尾元数据
- ✅ 去除加粗标记，保持格式统一
- ✅ 支持附件说明
- ✅ 自动添加页码

## 系统要求

### 必需软件
1. **Python 3.6+**
2. **Pandoc**（需要单独安装）
   - 安装说明：https://pandoc.org/installing.html
   - macOS: `brew install pandoc`
   - Ubuntu: `sudo apt install pandoc`
   - Windows: 下载安装包安装

### Python依赖
```bash
pip install -r requirements.txt
```

依赖包：
- `python-docx==0.8.11` - Word文档操作
- `pypandoc==1.11` - Pandoc Python接口

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

## 转换流程

本工具采用**预处理 → Pandoc转换 → 后处理**三阶段架构：

1. **预处理阶段**（markdown_preprocessor.py）
   - 过滤YAML front matter
   - 去除加粗标记
   - 保持LaTeX公式格式
   - 智能处理换行（避免破坏表格和列表）
   - 过滤结尾元数据

2. **Pandoc转换阶段**（pandoc_processor.py）
   - 使用MathML渲染数学公式
   - 转换表格格式
   - 处理多级列表
   - 生成基础Word文档

3. **后处理阶段**（word_postprocessor.py）
   - 应用GB/T 9704-2012格式要求
   - 设置字体、字号、行距
   - 格式化标题和正文
   - 添加页码和附件信息

## Markdown格式要求

### 标题层级对应关系

- `#` → 被忽略（使用文件名作为文档标题）
- `##` → 一级标题（黑体，三号，不加粗）
- `###` → 二级标题（Kaiti TC，三号，不加粗）
- `####` → 三级标题（仿宋，三号，不加粗）

### 数学公式支持

**行内公式**：
```markdown
载流子浓度达 $1.78 \times 10^{18}$ cm⁻³
```

**块级公式**：
```markdown
$$BFOM = \frac{V_B^2}{R_{on,sp}} \propto \epsilon\mu E_c^3$$
```

### 表格支持

```markdown
| 材料 | 禁带宽度(eV) | 击穿场强(MV/cm) |
|------|-------------|----------------|
| Si   | 1.12        | 0.3            |
| GaN  | 3.4         | 4.9            |
| SiC  | 3.3         | 3.1            |
```

### 列表支持

**有序列表**：
```markdown
1. **第一个要点：技术突破**
2. **第二个要点：市场优势**
3. **第三个要点：团队实力**
```

**无序列表**：
```markdown
* 第一项内容
* 第二项内容
  * 二级项目1
  * 二级项目2
```

### 图片支持

**标准Markdown格式**：
```markdown
![图片描述](image.png)
![相对路径](./images/photo.jpg)
```

**Obsidian格式**（推荐）：
```markdown
![[attachment_file.png]]
![图 1：工艺流程](process_diagram.png)
```

**图片功能特性**：
- 智能路径解析（支持多个搜索目录）
- Obsidian附件目录集成
- 自动设置Top and Bottom文字环绕
- 图片水平居中对齐
- 智能caption处理（移除文件名，保留有意义标题）
- Caption格式：仿宋4号字体，居中显示

### 自动处理功能

1. **YAML Front Matter**：自动过滤`---`包围的元数据
2. **结尾元数据**：自动过滤Date行、标签（如`#work`）
3. **加粗格式**：去除`**文字**`和`__文字__`标记

### 完整示例

```markdown
---
项目: 示例
类别: Reports
---

# 文档标题（会被忽略，使用文件名）

## 一、项目概述

正文内容，支持数学公式 $E = mc^2$ 和表格。

### （一）技术参数

| 参数 | 数值 |
|------|------|
| 功率 | 100W |

1. **第一个要点**：详细说明
2. **第二个要点**：包含公式 $P = I^2R$

**加粗文字**会被去除格式。

附件：1. 技术规格书
附件：2. 测试报告

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
  - 表格：仿宋、四号
  - 图片caption：仿宋、四号、居中
- **行距**：固定值25.5磅
- **对齐**：
  - 文档标题：居中
  - 其他所有文字：首行缩进2字符，两端对齐
- **页码**：仿宋、四号、居中（格式：`- 1 -`）

## 项目结构

```
md-to-word/
├── md_to_word.py              # 主程序入口
├── requirements.txt           # Python依赖列表
├── config.py                 # 公文格式配置（包含Pandoc参数）
├── markdown_preprocessor.py  # Markdown预处理器
├── pandoc_processor.py       # Pandoc转换处理器
├── word_postprocessor.py     # Word后处理器（格式化）
├── chinese_filter.lua        # Pandoc中文处理过滤器
├── README.md                 # 使用说明（本文件）
├── CLAUDE.md                 # 项目开发记录
└── examples/                 # 示例和测试文件
    ├── test_formatting.md    # 格式化测试文档
    ├── header_test.md        # 标题测试文档
    ├── list_test.md          # 列表测试文档
    └── *.docx               # 转换结果示例
```

## 支持的输入格式

- 标准Markdown文件
- 带YAML front matter的Markdown
- 包含LaTeX数学公式的Markdown
- 包含表格的Markdown文件
- 包含多级列表的Markdown
- 包含图片的Markdown文件（支持Obsidian格式）
- Obsidian笔记格式

## 技术架构

- **转换引擎**：Pandoc（专业文档转换）
- **数学公式**：MathML渲染（原生Word公式）
- **表格处理**：Pandoc原生转换 + 格式美化
- **列表处理**：保持层级结构和缩进
- **图片处理**：智能路径解析 + 格式化 + 文字环绕
- **字体处理**：符合国标的中文字体设置

## 注意事项

1. **Pandoc依赖**：确保系统已安装Pandoc
2. **字体要求**：确保系统中已安装相应的中文字体
3. **格式兼容**：生成的Word文档已设置为符合公文标准
4. **公式渲染**：LaTeX公式转为原生Word公式对象
5. **表格样式**：自动应用符合国标的表格格式

## 版本信息

- **当前版本**：2.0.0
- **更新日期**：2025年7月21日
- **重大更新**：基于Pandoc架构重构，新增完整的LaTeX公式、表格和列表支持

## 问题反馈

如遇到转换问题，请检查：
1. Pandoc是否正确安装：`pandoc --version`
2. Python依赖是否完整：`pip list`
3. 输入文件编码是否为UTF-8
4. 参考`examples/`目录中的示例文件格式