import argparse
import os
import sys

import yaml
from yaml import FullLoader

sys.path.append('..')

from src.parser.md_parser import md2html
import time

from src.provider.docx_processor import DocxProcessor

config: dict = {
    "version": "0.1.0"
}


def find_md_files(path):
    """递归查找目录下所有的 markdown 文件"""
    md_files = []
    if os.path.isfile(path):
        if path.lower().endswith('.md'):
            md_files.append(path)
    else:
        for root, _, files in os.walk(path):
            for file in files:
                if file.lower().endswith('.md'):
                    md_files.append(os.path.join(root, file))
    return md_files


# 生成资源文件目录访问路径
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):  # 是否Bundle Resource
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        sys.argv.append("-h")

    parser = argparse.ArgumentParser(description="markdocx - %s" % config["version"])
    parser.add_argument('input', help="Markdown file or directory path")
    parser.add_argument('-o', '--output', help="Optional. Directory to save docx files (required when input is directory)")
    parser.add_argument('-s', '--style',
                        help="Optional. YAML file with style configuration")
    parser.add_argument('-a', action="store_true",
                        help="Optional. Automatically open docx file when finished converting")
    args = parser.parse_args()

    # 获取所有需要处理的 markdown 文件
    md_files = find_md_files(args.input)
    if not md_files:
        print(f"[ERROR] No markdown files found in {args.input}")
        sys.exit(1)

    # 如果输入是目录，输出参数必须指定
    if os.path.isdir(args.input) and not args.output:
        print("[ERROR] Output directory (-o/--output) must be specified when input is a directory")
        sys.exit(1)

    # 载入样式配置
    if not args.style:
        args.style = resource_path(os.path.join("config", "default_style.yaml"))

    with open(args.style, "r", encoding="utf-8") as file:
        conf = yaml.load(file, FullLoader)

    start_time = time.time()  # 记录总转换耗时
    
    for md_file in md_files:
        file_start_time = time.time()
        
        # 确定输出路径
        if os.path.isdir(args.input):
            # 保持目录结构
            rel_path = os.path.relpath(md_file, args.input)
            docx_path = os.path.join(args.output, os.path.splitext(rel_path)[0] + '.docx')
            # 确保输出目录存在
            os.makedirs(os.path.dirname(docx_path), exist_ok=True)
        else:
            docx_path = args.output if args.output is not None else md_file + ".docx"

        # 转换过程
        html_path = md_file + ".html"
        md2html(md_file, html_path)
        
        DocxProcessor(style_conf=conf).html2docx(html_path, docx_path)
        
        file_done_time = time.time()
        print(f"[SUCCESS] Converted {md_file} in: %.4f sec(s)" % (file_done_time - file_start_time))
        print(f"[SUCCESS] Saved to: {os.path.abspath(docx_path)}")
        
        # 如果是单文件且指定了自动打开，则打开文件
        if args.a and not os.path.isdir(args.input):
            os.startfile(os.path.abspath(docx_path))

    done_time = time.time()
    print("\n[SUCCESS] All conversions finished in: %.4f sec(s)" % (done_time - start_time))
