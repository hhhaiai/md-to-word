# 配置指南

## 更新说明 (v2.6.0)

- 移除了未使用的配置选项：`image_width`、`image_dpi`、`citation_style`
- 简化了表格配置：移除了 `row_height_rule` 选项
- 移除了 pandoc 参数中的 `--reference-links`

本文档详细介绍 md-to-word 工具的配置选项，特别是环境变量的使用方法。

## 环境变量配置

md-to-word 支持通过环境变量配置 Obsidian 相关路径，方便在不同环境下使用，无需修改代码。

### 支持的环境变量

#### OBSIDIAN_VAULT_NAME

**描述**：指定 Obsidian Vault（仓库）的名称。

**默认值**：`YT's Obsidian`

**用途**：
- 用于自动定位 Obsidian 仓库路径
- 配合自动检测功能查找图片和附件
- 支持多个 Vault 的用户切换使用

**示例**：
```bash
# Linux/macOS
export OBSIDIAN_VAULT_NAME="我的笔记"

# Windows (PowerShell)
$env:OBSIDIAN_VAULT_NAME="我的笔记"

# Windows (CMD)
set OBSIDIAN_VAULT_NAME=我的笔记
```

#### OBSIDIAN_ATTACHMENTS_FOLDER

**描述**：指定 Obsidian 中附件文件夹的名称。

**默认值**：`- Attachments`

**用途**：
- 定位 Obsidian 中存储图片和其他附件的文件夹
- 支持自定义附件管理策略
- 确保图片引用能正确解析

**示例**：
```bash
# Linux/macOS
export OBSIDIAN_ATTACHMENTS_FOLDER="附件"

# Windows (PowerShell)
$env:OBSIDIAN_ATTACHMENTS_FOLDER="附件"

# Windows (CMD)
set OBSIDIAN_ATTACHMENTS_FOLDER=附件
```

#### OBSIDIAN_VAULT_PATH

**描述**：直接指定 Obsidian Vault 的完整路径。

**默认值**：`None`（未设置时使用自动检测）

**用途**：
- 当自动检测无法找到 Vault 时使用
- 优先级高于自动检测
- 适用于非标准位置的 Vault

**示例**：
```bash
# Linux/macOS
export OBSIDIAN_VAULT_PATH="/Users/username/Documents/MyVault"

# Windows (PowerShell)
$env:OBSIDIAN_VAULT_PATH="C:\Users\username\Documents\MyVault"

# Windows (CMD)
set OBSIDIAN_VAULT_PATH=C:\Users\username\Documents\MyVault
```

## 使用场景

### 场景1：使用默认 iCloud 同步的 Obsidian

如果您的 Obsidian Vault 存储在 macOS 的 iCloud 默认位置，通常只需设置 Vault 名称：

```bash
export OBSIDIAN_VAULT_NAME="我的知识库"
```

工具会自动在以下路径查找：
- `~/Library/Mobile Documents/iCloud~md~obsidian/Documents/我的知识库`

### 场景2：自定义位置的 Obsidian Vault

如果您的 Vault 不在标准位置，可以直接指定完整路径：

```bash
export OBSIDIAN_VAULT_PATH="/home/user/my-notes"
export OBSIDIAN_ATTACHMENTS_FOLDER="images"
```

### 场景3：临时使用不同配置

可以在运行命令时临时设置环境变量：

```bash
# Linux/macOS
OBSIDIAN_VAULT_NAME="工作笔记" python3 md_to_word.py document.md

# Windows (PowerShell)
$env:OBSIDIAN_VAULT_NAME="工作笔记"; python md_to_word.py document.md
```

### 场景4：在脚本中使用

创建一个批处理脚本或 shell 脚本来固定配置：

**Linux/macOS (convert.sh)**：
```bash
#!/bin/bash
export OBSIDIAN_VAULT_NAME="我的笔记"
export OBSIDIAN_ATTACHMENTS_FOLDER="附件"
python3 /path/to/md_to_word.py "$@"
```

**Windows (convert.bat)**：
```batch
@echo off
set OBSIDIAN_VAULT_NAME=我的笔记
set OBSIDIAN_ATTACHMENTS_FOLDER=附件
python C:\path\to\md_to_word.py %*
```

## 配置优先级

1. **环境变量 OBSIDIAN_VAULT_PATH**（最高优先级）
   - 如果设置了此变量，直接使用指定路径

2. **自动检测 + OBSIDIAN_VAULT_NAME**
   - 在标准位置查找指定名称的 Vault
   - macOS: `~/Library/Mobile Documents/iCloud~md~obsidian/Documents/`
   - 备选位置: `~/Documents/`, `~/Desktop/`

3. **默认搜索路径**（最低优先级）
   - `./images`
   - `./assets`
   - `./`（当前目录）

## 故障排除

### 图片无法找到

1. 检查环境变量是否正确设置：
   ```bash
   echo $OBSIDIAN_VAULT_NAME
   echo $OBSIDIAN_VAULT_PATH
   ```

2. 确认 Vault 路径存在：
   ```bash
   ls -la "$OBSIDIAN_VAULT_PATH"
   ```

3. 验证附件文件夹名称：
   ```bash
   ls -la "$OBSIDIAN_VAULT_PATH/$OBSIDIAN_ATTACHMENTS_FOLDER"
   ```

### 自动检测失败

如果自动检测无法找到您的 Obsidian Vault，请：

1. 使用 `OBSIDIAN_VAULT_PATH` 直接指定路径
2. 确保 Vault 名称与实际文件夹名称完全一致（包括大小写）
3. 检查是否有权限访问相关目录

## 高级配置

除了环境变量，您还可以通过修改 `src/config/config.py` 文件来调整更多设置：

- **字体配置**：修改 `FONTS` 字典
- **页面设置**：调整 `PAGE_MARGINS` 
- **图片格式**：配置 `IMAGE_CONFIG`
- **Pandoc参数**：自定义 `PANDOC_CONFIG`

详细配置选项请参考源代码中的注释。