# 项目状态记录 - CLAUDE.md

## 项目概述

Markdown到Word公文格式转换工具，符合GB/T 9704-2012《党政机关公文格式》国家标准。

**重要更新：**已完成架构重构，现在基于Pandoc引擎，支持LaTeX公式、表格和列表。

## 已完成功能

### ✅ 核心转换功能（基于Pandoc）
- [x] 基于Pandoc的Markdown到Word转换
- [x] **LaTeX数学公式原生渲染**（使用MathML）
- [x] **完整表格支持**（Markdown表格→Word表格）
- [x] **多级列表支持**（有序列表、无序列表）
- [x] 公文格式设置（页边距、字体、字号、行距）
- [x] 标题层级处理
- [x] 页码添加
- [x] 命令行接口

### ✅ 格式处理（预处理阶段）
- [x] YAML front matter过滤
- [x] 结尾元数据过滤（Date、标签等）
- [x] 加粗标记去除（**文字**、__文字__）
- [x] **LaTeX数学公式保持原格式**（交给pandoc处理）
- [x] 成文日期汉字转换
- [x] 附件说明格式化
- [x] 改进的换行处理（避免破坏表格和列表）

### ✅ 字体和格式设置（后处理阶段）
- [x] 文档标题：小标宋体、二号、居中
- [x] 正文：FangSong、三号、首行缩进2字符、两端对齐
- [x] 一级标题（##）：黑体、三号、不加粗、首行缩进2字符、两端对齐
- [x] 二级标题（###）：Kaiti TC、三号、不加粗、首行缩进2字符、两端对齐
- [x] 三级标题（####）：仿宋、三号、不加粗、首行缩进2字符、两端对齐
- [x] 附件说明：仿宋、三号、首行缩进2字符、两端对齐
- [x] **表格格式化**：仿宋、三号、两端对齐
- [x] **列表格式化**：保持层级结构，应用正确字体

### ✅ 智能处理
- [x] 使用文件名作为文档标题
- [x] 忽略Markdown中的#标题
- [x] 去除主送机关自动识别（避免误识别）
- [x] Pandoc可用性检查

## ✅ 已解决问题

### 1. 异常字符间距问题 [已修复]
**状态**：已通过Pandoc架构解决
**解决方案**：使用Pandoc的原生处理避免了自定义解析导致的换行问题

### 2. LaTeX公式支持 [已完成]
**状态**：现在支持完整的LaTeX数学公式
**实现**：使用Pandoc的MathML渲染

### 3. 表格支持 [已完成]
**状态**：完整支持Markdown表格转Word表格
**实现**：Pandoc原生处理 + 后处理格式化

### 4. 列表支持 [已完成]
**状态**：支持有序和无序多级列表
**实现**：Pandoc原生处理 + 格式保持

## 新架构说明

### 架构模式
**Markdown → 预处理 → Pandoc转换 → 后处理 → Word**

### 组件说明
1. **markdown_preprocessor.py**：保留所有原有过滤逻辑
2. **pandoc_processor.py**：Pandoc集成和转换控制
3. **word_postprocessor.py**：对Pandoc输出应用GB/T格式
4. **config.py**：格式配置（包含Pandoc参数）
5. **chinese_filter.lua**：Pandoc中文处理过滤器（可选）

## 项目结构

```
md-to-word/
├── md_to_word.py              # 主程序入口（已更新）
├── requirements.txt           # Python依赖包（新增pypandoc）
├── config.py                 # 公文格式配置（新增Pandoc配置）
├── markdown_preprocessor.py  # Markdown预处理器（重构）
├── pandoc_processor.py       # Pandoc转换处理器（新增）
├── word_postprocessor.py     # Word后处理器（重构）
├── chinese_filter.lua        # Pandoc中文过滤器（新增）
├── README.md                 # 项目说明文档
├── CLAUDE.md                 # 项目状态记录（本文件）
└── examples/                # 示例和测试文件
    ├── test_features.md      # 功能测试文档（新增）
    ├── 广州拓诺稀科技有限公司投资建议书.md
    └── *.docx
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
   - 保持LaTeX公式格式
   - 智能合并分割行（避免破坏表格列表）
   - 过滤结尾元数据
   - 跳过一级标题
3. **Pandoc转换**（pandoc_processor.py）：
   - MathML数学公式渲染
   - 表格转换
   - 列表处理
   - 基本Word格式生成
4. **后处理格式化**（word_postprocessor.py）：
   - 设置页面格式
   - 添加文档标题（使用文件名）
   - 应用公文字体格式
   - 设置标题和正文格式
   - 格式化表格
   - 添加附件说明和成文日期
   - 设置页码
5. **输出**：符合公文格式的Word文档

## 测试文件

### 功能测试
- `examples/test_features.md`：包含LaTeX公式、表格、列表的综合测试
- `examples/广州拓诺稀科技有限公司投资建议书.md`：复杂实际文档测试

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
- 当前版本：2.0.0
- 重大更新：2025-07-21
- 状态：**功能完整，支持LaTeX公式、表格、列表**