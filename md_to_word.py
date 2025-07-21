#!/usr/bin/env python3
import argparse
import os
import sys
from pathlib import Path

from markdown_parser import MarkdownParser
from word_generator import WordGenerator

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
        version='%(prog)s 1.0.0'
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
        print(f"正在解析Markdown文件: {input_path}")
        
        # 解析Markdown文件
        parser = MarkdownParser()
        parsed_data = parser.parse_file(str(input_path))
        
        print(f"正在生成Word文档: {output_path}")
        
        # 生成Word文档
        generator = WordGenerator()
        generator.create_document(parsed_data, str(output_path))
        
        print(f"转换完成！输出文件: {output_path}")
        
    except Exception as e:
        print(f"转换过程中出现错误: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()