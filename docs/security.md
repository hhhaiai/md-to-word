# Security Policy

## Security Measures

This project implements several security measures to protect against common vulnerabilities:

### 1. Command Injection Protection
- All shell arguments are escaped using `shlex.quote()`
- Pandoc command arguments are properly sanitized
- No user input is directly passed to shell commands

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