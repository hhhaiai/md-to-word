# Security Policy

## Security Measures

This project implements several security measures to protect against common vulnerabilities:

### 1. Command Injection Protection
- Pandoc 调用使用 `subprocess.run([...])` 列表参数方式传参，避免 shell 解析
- 所有用户输入参数不会通过 shell 直接执行
- 参数构造集中在 `src/core/pandoc_processor.py::_get_pandoc_args`

### 2. Path Traversal Protection
- Comprehensive path validation using `path_validator.py`
- Validates against:
  - Directory traversal attempts (`..`, `.`, `~`)
  - Symbolic links
  - Absolute paths (when not allowed)
  - Hidden files/directories
- Uses `Path.is_relative_to()` for containment checking

### 3. XML Injection Protection
- Uses proper XML APIs for element creation
- Avoids string concatenation for XML construction
- Sanitizes user input before XML processing

### 4. Non-interactive Safety
- CLI 提供 `--force` 标志用于非交互覆盖输出文件，避免阻塞式输入

## Reporting Security Vulnerabilities

If you discover a security vulnerability, please:
1. **DO NOT** create a public issue
2. Email the maintainers directly with details
3. Allow reasonable time for a fix before public disclosure

## Security Audit History

### v2.6.0 (2025-07-26)
- Fixed command injection vulnerability in `pandoc_processor.py`
- Fixed path traversal vulnerability with comprehensive validation
- Removed dead code and consolidated duplicate methods
- Simplified configuration to reduce attack surface