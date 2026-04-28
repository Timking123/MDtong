#!/usr/bin/env python3
"""会议纪要 Markdown → DOCX 转换工具（主入口）"""

import os
import sys
import argparse

from converter import run_gui
from converter.services import convert_file

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    if len(sys.argv) > 1 and not sys.argv[1].startswith('--gui'):
        parser = argparse.ArgumentParser(description='将 Markdown 会议纪要转换为 .docx 文件')
        parser.add_argument('input_file', nargs='?', help='Markdown 文件路径（省略则打开 GUI）')
        parser.add_argument('-o', '--output', help='输出 .docx 文件路径')
        args = parser.parse_args()

        if args.input_file:
            if not os.path.isfile(args.input_file):
                print(f'错误：找不到文件 "{args.input_file}"', file=sys.stderr)
                sys.exit(1)
            print(f'已从文件读取：{args.input_file}')
            try:
                result = convert_file(args.input_file, args.output)
                print(f'转换完成：{result.docx_path}')
            except Exception as e:
                print(f'错误：{e}', file=sys.stderr)
                sys.exit(1)
        else:
            run_gui()
    else:
        run_gui()


if __name__ == '__main__':
    main()
