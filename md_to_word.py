#!/usr/bin/env python3
import argparse
import os
import sys
from pathlib import Path

from markdown_preprocessor import MarkdownPreprocessor
from pandoc_processor import PandocProcessor
from word_postprocessor import WordPostprocessor
from exceptions import FileProcessingError, PandocError

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
    
    args = parser.parse_args()
    
    # 检查输入文件是否存在
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：输入文件 '{args.input}' 不存在", file=sys.stderr)
        sys.exit(1)
    
    if not input_path.suffix.lower() in ['.md', '.markdown']:
        print(f"警告：输入文件 '{args.input}' 可能不是Markdown文件", file=sys.stderr)
    
    # 确定输出文件路径
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.docx')
    
    # 确保输出目录存在
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        print(f"正在预处理Markdown文件: {input_path}")
        
        # 预处理Markdown文件
        preprocessor = MarkdownPreprocessor()
        preprocessed_data = preprocessor.preprocess_file(str(input_path))
        
        
        print(f"正在使用Pandoc转换为Word文档: {output_path}")
        
        # 使用pandoc转换为Word文档
        pandoc_processor = PandocProcessor()
        
        # 检查pandoc是否可用
        if not pandoc_processor.check_pandoc_available():
            print("错误：Pandoc未安装或不可用。请安装pandoc后再试。", file=sys.stderr)
            print("安装说明：https://pandoc.org/installing.html", file=sys.stderr)
            sys.exit(1)
        
        # 转换为Word文档
        temp_output = pandoc_processor.convert_markdown_to_docx(
            preprocessed_data['content'], 
            str(output_path),
            title=None  # 不在pandoc阶段添加标题，后处理时添加
        )
        
        print(f"正在应用公文格式...")
        
        # 应用公文格式
        postprocessor = WordPostprocessor()
        postprocessor.apply_formatting(temp_output, preprocessed_data, preprocessed_data['content'])
        
        # 格式化表格（如果有）
        postprocessor.format_tables()
        
        print(f"转换完成！输出文件: {output_path}")
        
    except FileProcessingError as e:
        print(f"文件处理错误: {e}", file=sys.stderr)
        sys.exit(1)
    except PandocError as e:
        print(f"Pandoc转换错误: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"转换过程中出现未知错误: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()