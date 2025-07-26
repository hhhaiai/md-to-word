#!/usr/bin/env python3
import argparse
import os
import sys
import logging
from pathlib import Path

from src.core.markdown_preprocessor import MarkdownPreprocessor
from src.core.pandoc_processor import PandocProcessor
from src.core.word_postprocessor import WordPostprocessor
from src.utils.exceptions import FileProcessingError, PandocError, PathSecurityError

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stderr)
    ]
)
from src.utils.config_validator import ConfigValidator

def main():
    """主程序入口"""
    parser = argparse.ArgumentParser(
        description='将Markdown文件转换为符合GB/T 9704-2012标准的Word公文格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  %(prog)s input.md                    # 输出到 input.docx
  %(prog)s input.md -o output.docx     # 指定输出文件
  %(prog)s input.md --output output.docx
        """
    )
    
    parser.add_argument(
        'input',
        nargs='?',
        help='输入的Markdown文件路径'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='输出的Word文件路径（默认为输入文件名.docx）'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 2.0.0'
    )
    
    parser.add_argument(
        '--check-config',
        action='store_true',
        help='仅检查配置而不进行转换'
    )
    
    parser.add_argument(
        '--skip-validation',
        action='store_true',
        help='跳过配置验证（不推荐）'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='启用详细日志输出（用于调试）'
    )
    
    args = parser.parse_args()
    
    # 设置日志级别
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # 如果只是检查配置
    if args.check_config:
        print("检查配置...\n")
        validator = ConfigValidator()
        is_valid, results = validator.validate_all()
        validator.print_results(results)
        sys.exit(0 if is_valid else 1)
    
    # 运行配置验证（除非跳过）
    if not args.skip_validation:
        validator = ConfigValidator()
        is_valid, results = validator.validate_all()
        
        # 只在有错误时显示验证结果
        if not is_valid:
            print("配置验证失败：")
            validator.print_results(results)
            print("\n请修复以上错误后再运行转换，或使用 --skip-validation 跳过验证（不推荐）")
            sys.exit(1)
    
    # 检查是否提供了输入文件
    if not args.input:
        parser.error("必须提供输入文件路径")
    
    # 输入验证
    try:
        # 验证路径安全性
        input_path = Path(args.input).resolve()
        if ".." in args.input:
            raise PathSecurityError("输入路径包含不安全的目录遍历")
        
        # 检查文件是否存在
        if not input_path.exists():
            print(f"错误：文件不存在: {args.input}", file=sys.stderr)
            sys.exit(1)
        
        # 检查是否为文件（不是目录）
        if not input_path.is_file():
            print(f"错误：不是文件: {args.input}", file=sys.stderr)
            sys.exit(1)
        
        # 检查文件扩展名
        valid_extensions = ['.md', '.markdown', '.mdown', '.mkd', '.mdwn']
        if not input_path.suffix.lower() in valid_extensions:
            print(f"错误：需要Markdown文件 ({', '.join(valid_extensions)})", file=sys.stderr)
            sys.exit(1)
        
        # 检查文件是否可读
        if not os.access(input_path, os.R_OK):
            print(f"错误：无法读取文件: {args.input}", file=sys.stderr)
            sys.exit(1)
        
        # 检查文件大小（限制为100MB）
        max_size_mb = 100
        file_size_mb = input_path.stat().st_size / (1024 * 1024)
        if file_size_mb > max_size_mb:
            print(f"错误：文件太大 ({file_size_mb:.1f}MB > {max_size_mb}MB)", file=sys.stderr)
            sys.exit(1)
            
    except PathSecurityError as e:
        print(f"安全错误：{e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"输入验证错误：{e}", file=sys.stderr)
        sys.exit(1)
    
    # 验证和设置输出路径
    try:
        if args.output:
            # 验证用户提供的输出路径
            output_path = Path(args.output).resolve()
            if ".." in args.output:
                raise PathSecurityError("输出路径包含不安全的目录遍历")
            
            # 确保输出文件扩展名正确
            if not output_path.suffix.lower() in ['.docx', '.doc']:
                output_path = output_path.with_suffix('.docx')
        else:
            # 使用输入文件名生成输出文件名
            output_path = input_path.with_suffix('.docx')
        
        # 检查输出目录是否可写
        output_dir = output_path.parent
        if output_dir.exists() and not os.access(output_dir, os.W_OK):
            # 使用相对路径显示
            try:
                rel_output_dir = os.path.relpath(output_dir)
            except ValueError:
                rel_output_dir = str(output_dir)
            print(f"错误：无法写入目录: {rel_output_dir}", file=sys.stderr)
            sys.exit(1)
        
        # 确保输出目录存在
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 如果输出文件已存在，询问是否覆盖
        if output_path.exists():
            # 使用相对路径显示
            try:
                rel_output_path = os.path.relpath(output_path)
            except ValueError:
                rel_output_path = str(output_path)
            response = input(f"文件已存在: {rel_output_path}，覆盖？(y/N): ")
            if response.lower() != 'y':
                print("已取消操作")
                sys.exit(0)
                
    except PathSecurityError as e:
        print(f"安全错误：{e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"输出路径错误：{e}", file=sys.stderr)
        sys.exit(1)
    
    try:
        # 预处理Markdown文件
        preprocessor = MarkdownPreprocessor()
        preprocessed_data = preprocessor.preprocess_file(str(input_path))
        
        # 使用pandoc转换为Word文档
        pandoc_processor = PandocProcessor()
        
        # 检查pandoc是否可用
        if not pandoc_processor.check_pandoc_available():
            print("错误：需要安装Pandoc\n访问: https://pandoc.org/installing.html", file=sys.stderr)
            sys.exit(1)
        
        # 转换为Word文档
        temp_output = pandoc_processor.convert_markdown_to_docx(
            preprocessed_data['content'], 
            str(output_path),
            title=None  # 不在pandoc阶段添加标题，后处理时添加
        )
        
        # 应用公文格式
        postprocessor = WordPostprocessor()
        postprocessor.apply_formatting(temp_output, preprocessed_data, preprocessed_data['content'])
        
        # 格式化表格（如果有）
        postprocessor.format_tables()
        
        # 验证输出文件是否成功创建
        if not output_path.exists():
            raise FileProcessingError("输出文件未成功创建")
        
        # 显示文件大小信息
        output_size_mb = output_path.stat().st_size / (1024 * 1024)
        # 使用相对路径显示
        try:
            rel_output_path = os.path.relpath(output_path)
        except ValueError:
            rel_output_path = str(output_path)
        print(f"完成: {rel_output_path} ({output_size_mb:.2f}MB)")
        
    except FileProcessingError as e:
        print(f"错误: {e}", file=sys.stderr)
        import logging
        logging.error(f"FileProcessingError: {e}", exc_info=True)
        sys.exit(1)
    except PandocError as e:
        print(f"Pandoc错误: {e}", file=sys.stderr)
        import logging
        logging.error(f"PandocError: {e}", exc_info=True)
        sys.exit(1)
    except PathSecurityError as e:
        print(f"路径错误: {e}", file=sys.stderr)
        import logging
        logging.error(f"PathSecurityError: {e}", exc_info=True)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n中断", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"未知错误: {e}", file=sys.stderr)
        import logging
        logging.error(f"Unexpected error: {e}", exc_info=True)
        sys.exit(1)

if __name__ == '__main__':
    main()