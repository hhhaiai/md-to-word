# Markdown到Word公文格式转换工具

将Markdown文件转换为符合GB/T 9704-2012《党政机关公文格式》标准的Word文档。

## 功能特点

- 符合GB/T 9704-2012国家标准
- 支持LaTeX数学公式、表格和列表
- 智能处理Obsidian格式图片
- 自动设置公文格式（页边距、字体、字号、页码）

## 安装

### 系统要求
- Python 3.6+
- Pandoc（必须）：[安装说明](https://pandoc.org/installing.html)

### 安装依赖
```bash
pip3 install -r requirements.txt
```

## 项目结构

```
md-to-word/
├── src/
│   ├── core/                    # 核心处理模块
│   ├── formatters/              # 格式化器模块
│   ├── utils/                   # 工具模块
│   └── config/
│       └── config.py            # 格式配置文件
├── examples/                    # 示例文件
├── md_to_word.py               # 主程序
└── requirements.txt            # 依赖列表
```

## 配置说明

可在 `src/config/config.py` 中调整以下设置：
- 字体和字号配置
- 页边距设置
- 图片处理选项
- Pandoc转换参数

## 使用方法

```bash
# 基本用法
python3 md_to_word.py document.md

# 指定输出文件
python3 md_to_word.py input.md -o output.docx
```

## 格式说明

### 标题层级
- `##` → 一级标题（黑体）
- `###` → 二级标题（楷体）
- 其他 → 正文（仿宋）

### 支持的Markdown元素
- 数学公式：`$E=mc^2$`、`$$E=mc^2$$`
- 表格：标准Markdown表格语法
- 列表：有序列表自动转为正文格式
- 图片：支持`![]()`和`![[]]`格式

## 版本信息

**v2.4.0**（2025-07-24）
- 修复有序列表和多级编号处理问题
- 确保字体格式统一

## 许可

MIT License