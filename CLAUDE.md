# 项目状态记录 - CLAUDE.md

## 项目概述

Markdown到Word公文格式转换工具，符合GB/T 9704-2012《党政机关公文格式》国家标准。

**重要更新：**已完成架构重构和性能优化，现在基于Pandoc引擎，支持LaTeX公式、表格和列表，采用模块化设计和优化的处理流程。

**最新优化（2025-07-22）：**
- 完全修复了LaTeX数学公式转换问题，支持行内公式（`$...$`）和块级公式（`$$...$$`）
- 消除功能冗余，优化代码结构，每个处理环节职责更明确
- 重构caption处理逻辑，提高代码可读性和可维护性

## 已完成功能

### ✅ 核心转换功能（基于Pandoc）
- [x] 基于Pandoc的Markdown到Word转换
- [x] **LaTeX数学公式原生渲染**（使用MathML）
- [x] **完整表格支持**（Markdown表格→Word表格）
- [x] **多级列表支持**（有序列表、无序列表）
- [x] **图片导入支持**（智能路径解析，支持Obsidian附件目录）
- [x] 公文格式设置（页边距、字体、字号、行距）
- [x] 标题层级处理
- [x] 页码添加
- [x] 命令行接口

### ✅ 格式处理（预处理阶段）
- [x] YAML front matter过滤
- [x] 结尾元数据过滤（Date、标签等）
- [x] 加粗标记去除（**文字**、__文字__）
- [x] **数字列表空格去除**（`1. ` → `1.`）
- [x] **星号列表转换**（`* ` → `- `）
- [x] **Caption位置调整**（确保图表标题在图表后面）
- [x] **LaTeX数学公式保持原格式**（交给pandoc处理）
- [x] ~~成文日期汉字转换~~（已移除）
- [x] 附件说明格式化
- [x] 改进的换行处理（避免破坏表格和列表）

### ✅ 字体和格式设置（后处理阶段）
- [x] 文档标题：小标宋体、二号、居中
- [x] 正文：FangSong、三号、首行缩进2字符、两端对齐
- [x] 一级标题（##）：黑体、三号、不加粗、首行缩进2字符、两端对齐
- [x] 二级标题（###）：Kaiti TC、三号、不加粗、首行缩进2字符、两端对齐
- [x] 三级标题（####）：仿宋、三号、不加粗、首行缩进2字符、两端对齐
- [x] 附件说明：仿宋、三号、首行缩进2字符、两端对齐
- [x] **移除日期处理功能**（2025-07-21）
- [x] **表格格式化**：仿宋、三号、两端对齐
- [x] **列表格式化**：保持层级结构，应用正确字体
- [x] **图片格式化**：智能移除图片文件名，保留有意义标题并格式化为caption
- [x] **图片文字环绕**：设置为Top and Bottom环绕模式，文字在图片上下方显示
- [x] **图片水平居中**：自动设置图片在页面中水平居中对齐
- [x] **Caption字体设置**：图片标题使用仿宋4号字体，居中显示
- [x] **图片标题优化**：智能处理图片标题显示

### ✅ 智能处理
- [x] 使用文件名作为文档标题
- [x] 忽略Markdown中的#标题
- [x] 去除主送机关自动识别（避免误识别）
- [x] **智能图片路径解析**（支持多搜索路径，自动扩展名匹配）
- [x] Pandoc可用性检查

## ✅ 已解决问题

### 1. 异常字符间距问题 [已修复]
**状态**：已通过Pandoc架构解决
**解决方案**：使用Pandoc的原生处理避免了自定义解析导致的换行问题

### 2. LaTeX公式支持 [已完成] ✅
**状态**：完全修复，支持完整的LaTeX数学公式转换
**实现**：
- 使用Pandoc直接调用（subprocess）替代pypandoc
- 修复预处理器对数学公式块的错误合并
- 修复后处理器破坏MathML内容的问题
- 支持行内公式（`$E=mc^2$`）和块级公式（`$$...$$`）

### 3. 表格支持 [已完成]
**状态**：完整支持Markdown表格转Word表格
**实现**：Pandoc原生处理 + 后处理格式化

### 4. 列表支持 [已完成]
**状态**：支持有序和无序多级列表
**实现**：Pandoc原生处理 + 格式保持

### 5. 代码架构重构 [已完成]
**状态**：完成模块化重构和性能优化
**实现**：
- 拆分815行的God Object为6个专门的格式化器类
- 预编译正则表达式模式，减少重复编译
- 实现XPath查询缓存和批量处理
- 修复XML注入和路径遍历安全漏洞
- 改进错误处理和异常类型

### 6. 图片处理优化 [已完成] 
**状态**：完成图片格式化和显示优化
**实现**：
- 智能图片标题处理和格式优化
- 完善图片显示效果和格式一致性

### 7. 功能冗余消除 [已完成] ✅
**状态**：完成代码重构和优化
**实现**：
- 重构`_reposition_captions`方法，从111行拆分为5个小函数
- 保留必要的列表处理功能
- 删除未使用的图片处理和空方法
- 简化后处理器中caption处理逻辑

## 新架构说明

### 架构模式
**Markdown → 预处理 → Pandoc转换 → 后处理 → Word**

### 组件说明
1. **markdown_preprocessor.py**：Markdown预处理（重构优化）
   - 文本清理和过滤
   - Caption位置调整
   - 列表格式处理
   - 元数据提取
2. **pandoc_processor.py**：Pandoc集成和转换控制
3. **word_postprocessor.py**：模块化格式化控制器（重构）
4. **formatters.py**：专业格式化器类集合
   - PageFormatter：页面格式、页码设置
   - ParagraphFormatter：段落和标题格式化
   - DocumentTitleFormatter：文档标题处理
   - TableFormatter：表格格式优化
   - ListFormatter：列表格式和缩进
   - ImageFormatter：图片处理和标题清理
5. **xpath_cache.py**：XPath查询优化和缓存
6. **exceptions.py**：专门的异常类型定义
7. **config.py**：格式配置（包含Pandoc参数）
8. **chinese_filter.lua**：Pandoc中文处理过滤器（可选）

## 🖼️ 图片支持功能详解

### 功能特性
- **智能路径解析**：自动在多个配置的搜索路径中查找图片
- **Obsidian集成**：完整支持Obsidian `![[]]` 格式和附件目录
- **多格式图片语法**：支持标准Markdown `![]()`和Obsidian `![[]]`语法
- **格式支持**：PNG、JPG、JPEG、GIF、BMP、SVG、WEBP
- **自动扩展名匹配**：支持无扩展名文件的自动匹配
- **智能图片标题处理**：移除文件名，保留有意义标题并格式化为caption
- **标题优化处理**：智能处理图片标题，保持文档整洁
- **文字环绕设置**：支持Top and Bottom文字环绕
- **图片对齐**：自动设置图片水平居中对齐
- **Caption格式化**：图片标题使用仿宋4号字体，居中显示

### 配置说明
**图片路径配置**（`config.py`的`IMAGE_CONFIG`）：
- `obsidian_attachments_path`：Obsidian附件目录路径
- `search_paths`：图片搜索路径列表（按优先级排序）
- `supported_formats`：支持的图片格式列表
- `copy_images`：是否复制图片到输出目录
- `output_dir`：图片输出目录名称

**图片格式配置**（`config.py`的`PANDOC_CONFIG`）：
- `image_wrap_text`：是否启用图片文字环绕（True/False）
- `image_wrap_type`：环绕类型（topAndBottom/square/tight/through/none）

### 搜索优先级
1. 源Markdown文件所在目录
2. Obsidian附件目录
3. `./images`目录
4. `./assets`目录  
5. 当前目录

### 使用示例
```markdown
![图片描述](image.png)                    # 标准Markdown格式
![相对路径](./images/photo.jpg)           # 相对路径引用
![[attachment_file.png]]                 # Obsidian格式（推荐）
![Obsidian图片](attachment_file.png)      # 标准格式引用Obsidian附件
```

## 项目结构

```
md-to-word/
├── md_to_word.py              # 主程序入口
├── requirements.txt           # Python依赖包
├── config.py                 # 公文格式配置
├── markdown_preprocessor.py  # Markdown预处理器（重构优化）
├── pandoc_processor.py       # Pandoc转换处理器
├── word_postprocessor.py     # Word后处理控制器（重构）
├── formatters.py             # 专业格式化器类集合（新增）
├── xpath_cache.py           # XPath查询优化器（新增）
├── exceptions.py            # 专门异常类型（新增）
├── chinese_filter.lua        # Pandoc中文处理过滤器
├── README.md                 # 项目说明文档
├── CLAUDE.md                 # 项目状态记录（本文件）
└── examples/                # 示例文件
    ├── *.md                 # 示例Markdown文件
    └── *.docx               # 转换结果示例
```

## 技术栈

- **Python 3**: 主要编程语言
- **pypandoc**: Pandoc Python接口
- **Pandoc**: 文档转换引擎（需要系统安装）
- **python-docx**: Word文档操作库
- **re**: 正则表达式处理

## 使用流程

1. **输入**：Markdown文件（支持任意路径）
2. **预处理**（markdown_preprocessor.py）：
   - 过滤YAML front matter
   - 去除加粗标记
   - 列表格式处理（去除空格、星号转换）
   - Caption位置调整
   - 保持LaTeX公式格式
   - 智能合并分割行（避免破坏表格列表）
   - 过滤结尾元数据
   - 跳过一级标题
3. **Pandoc转换**（pandoc_processor.py）：
   - MathML数学公式渲染
   - 表格转换
   - 列表处理
   - 基本Word格式生成
4. **后处理格式化**（word_postprocessor.py + formatters.py）：
   - 页面格式设置和页码添加
   - 文档标题处理（使用文件名）
   - 段落和标题格式化（公文标准）
   - 表格格式优化和对齐
   - 列表格式化和层级处理
   - 图片格式化和显示优化
   - 文档格式整理和优化
5. **输出**：符合公文格式的Word文档

## 示例文件

### 实际文档
- `examples/*.md`：复杂实际文档，包含图片、表格、列表等完整功能演示
- `examples/*.docx`：转换后的Word文档示例

## 依赖要求

### Python包
```
python-docx==0.8.11
pypandoc==1.11
```

### 系统要求
- **Pandoc**：需要在系统中安装pandoc
- 安装说明：https://pandoc.org/installing.html

## 版本信息
- 当前版本：2.1.0
- 最新更新：2025-07-22
- 状态：**功能完整，架构优化，代码精简**

## 性能优化特性
- **模块化架构**：6个专门的格式化器类，职责分离
- **批量处理**：减少DOM遍历次数，提升处理效率
- **缓存机制**：XPath查询结果缓存，避免重复计算
- **预编译模式**：正则表达式预编译，减少运行时开销
- **安全增强**：路径遍历防护，XML注入防护，异常处理优化