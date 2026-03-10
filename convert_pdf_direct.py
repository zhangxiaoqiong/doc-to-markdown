#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接发送PDF文件给Claude Vision API处理
"""

import os
import sys
import base64
from pathlib import Path
from anthropic import Anthropic

if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def pdf_to_base64(file_path):
    """将PDF文件转换为base64"""
    with open(file_path, 'rb') as f:
        return base64.standard_b64encode(f.read()).decode("utf-8")


def convert_pdf_with_vision(file_name, file_path):
    """使用Claude Vision API解析PDF"""
    client = Anthropic(
        api_key=os.getenv("ANTHROPIC_AUTH_TOKEN"),
        base_url=os.getenv("ANTHROPIC_BASE_URL")
    )

    # 读取PDF文件
    pdf_base64 = pdf_to_base64(file_path)

    print(f"正在用Vision API转换 {file_name}（PDF大小: {len(pdf_base64)/1024:.1f}KB）...")

    # 构建消息
    content = [
        {
            "type": "text",
            "text": f"""请将以下PDF文档完整转换为Markdown格式。

文件名：{file_name}

请：
1. 保留完整的文档结构（所有标题、小节、章节等）
2. 转换所有文本内容，确保准确
3. 对流程图、表格、图表等可视化内容进行清晰的描述和转换
4. 保留所有数据、数字、日期、金额等关键信息
5. 对于流程图，用清晰的文字和列表描述每个步骤
6. 对于表格，转换为Markdown表格格式，保留所有行列
7. 对于特别复杂的流程图或图表，可以用缩进列表清晰地表示层级关系
8. 确保内容完整准确，不删除或省略任何重要信息

请直接输出完整的、可以直接保存为.md文件的Markdown内容，不需要任何前缀或后缀说明。"""
        },
        {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": pdf_base64
            }
        }
    ]

    # 调用Claude Vision API（支持PDF文档类型）
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=8192,
        messages=[
            {"role": "user", "content": content}
        ]
    )

    return message.content[0].text


def process_pdfs(input_dir=None, output_dir=None, file_list=None):
    """处理PDF文件

    Args:
        input_dir: 输入目录路径（默认为当前目录）
        output_dir: 输出目录路径（默认为当前目录）
        file_list: 要处理的文件名列表（如果为None，处理目录中的所有.pdf文件）
    """
    if input_dir is None:
        input_dir = Path.cwd()
    else:
        input_dir = Path(input_dir)

    if output_dir is None:
        output_dir = Path.cwd()
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    # 如果未指定文件列表，扫描输入目录的PDF文件
    if file_list is None:
        pdf_files = list(input_dir.glob('*.pdf'))
        file_list = [f.name for f in pdf_files]

    if not file_list:
        print("未找到需要处理的PDF文件")
        return

    results = []

    for file_name in file_list:
        file_path = input_dir / file_name

        if not file_path.exists():
            print(f"[失败] 文件不存在: {file_path}")
            results.append((file_name, "文件不存在"))
            continue

        try:
            markdown_content = convert_pdf_with_vision(file_name, file_path)
            if not markdown_content:
                results.append((file_name, "转换失败"))
                continue

            # 保存输出
            output_file_name = file_name.rsplit('.', 1)[0] + '.md'
            output_path = output_dir / output_file_name

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)

            print(f"[成功] {output_file_name}\n")
            results.append((file_name, "成功"))

        except Exception as e:
            print(f"[失败] {file_name}: {str(e)}\n")
            results.append((file_name, f"错误: {str(e)}"))

    # 输出总结
    print("="*60)
    print("PDF Vision转换完成")
    print("="*60)
    success_count = sum(1 for _, status in results if status == "成功")
    print(f"成功: {success_count}/{len(results)}")
    for file_name, status in results:
        print(f"  {file_name}: {status}")


if __name__ == "__main__":
    if not os.getenv("ANTHROPIC_BASE_URL") or not os.getenv("ANTHROPIC_AUTH_TOKEN"):
        print("错误: 未设置环境变量")
        sys.exit(1)

    import argparse
    parser = argparse.ArgumentParser(description='使用Claude Vision API将PDF文件转换为Markdown格式')
    parser.add_argument('--input', '-i', default=None, help='输入目录（默认为当前目录）')
    parser.add_argument('--output', '-o', default=None, help='输出目录（默认为当前目录）')
    parser.add_argument('files', nargs='*', help='要转换的PDF文件列表（如果未指定，处理输入目录中的所有.pdf文件）')

    args = parser.parse_args()

    process_pdfs(
        input_dir=args.input,
        output_dir=args.output,
        file_list=args.files if args.files else None
    )
