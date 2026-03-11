#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务文档转换程序
使用Claude API将DOCX和PDF文件转换为Markdown格式
"""

import os
import sys
import time
from pathlib import Path
from anthropic import Anthropic

# 修复Windows编码问题
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    from docx import Document
except ImportError:
    print("需要安装 python-docx: pip install python-docx")
    sys.exit(1)

try:
    import pdfplumber
except ImportError:
    print("需要安装 pdfplumber: pip install pdfplumber")
    sys.exit(1)


def extract_docx_content(file_path):
    """从DOCX文件提取文本内容"""
    doc = Document(file_path)
    content = []

    for para in doc.paragraphs:
        if para.text.strip():
            content.append(para.text)

    # 提取表格
    for table in doc.tables:
        content.append("\n[TABLE]")
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            content.append(" | ".join(row_data))
        content.append("[/TABLE]\n")

    return "\n".join(content)


def extract_pdf_content(file_path):
    """从PDF文件提取文本内容和表格"""
    content = []

    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            content.append(f"--- Page {page_num} ---\n")

            # 提取文本
            text = page.extract_text()
            if text:
                content.append(text)

            # 提取表格
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    content.append("\n[TABLE]")
                    for row in table:
                        row_data = [str(cell) if cell else "" for cell in row]
                        content.append(" | ".join(row_data))
                    content.append("[/TABLE]\n")

    return "\n".join(content)


def convert_with_claude(file_name, file_content, max_retries=3):
    """使用Claude API将文档内容转换为Markdown，支持重试"""
    # 使用环境变量初始化客户端
    base_url = os.getenv("ANTHROPIC_BASE_URL")
    auth_token = os.getenv("ANTHROPIC_AUTH_TOKEN")

    client = Anthropic(
        api_key=auth_token,
        base_url=base_url
    )

    system_prompt = """你是一个专业的文档转换专家。你的任务是将用户提供的文档内容转换为结构清晰的Markdown格式。

转换要求：
1. 保留原文档的层级结构（标题、小节、段落等）
2. 使用适当的Markdown语法（# 为一级标题，## 为二级标题，等等）
3. 将[TABLE]...[/TABLE]标记之间的内容转换为Markdown表格格式
4. 保留所有原始数据、日期、数字、金额等关键信息
5. 保留列表和编号列表的结构
6. 确保内容完整，不删除或篡改任何信息
7. 使用清晰的分段和格式，提高可读性

你的输出应该是完整的、可以直接保存为.md文件的Markdown内容。"""

    user_message = f"""请将以下文档内容转换为Markdown格式。

文件名：{file_name}

文档内容：
```
{file_content}
```

请直接输出转换后的Markdown内容，不需要添加任何说明或标记。"""

    print(f"正在转换 {file_name}...", end=" ", flush=True)

    # 重试逻辑，指数退避
    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=4096,
                messages=[
                    {"role": "user", "content": user_message}
                ],
                system=system_prompt
            )
            print("✓")
            return message.content[0].text
        except Exception as e:
            error_str = str(e)
            # 判断是否是可重试的错误（429限流、500服务器错误）
            if "429" in error_str or "500" in error_str or "503" in error_str:
                if attempt < max_retries - 1:
                    wait_time = 2 ** attempt  # 指数退避：1s, 2s, 4s
                    print(f"\n  [重试 {attempt+1}/{max_retries-1}] 等待 {wait_time}s... ", end="", flush=True)
                    time.sleep(wait_time)
                    continue
            # 其他错误或最后一次重试失败
            raise


def process_files(input_dir=None, output_dir=None, file_list=None):
    """处理所有需要转换的文件

    Args:
        input_dir: 输入目录路径（默认为当前目录）
        output_dir: 输出目录路径（默认为当前目录）
        file_list: 要处理的文件名列表（如果为None，处理目录中的所有.docx和.pdf文件）
    """
    if input_dir is None:
        input_dir = Path.cwd()
    else:
        input_dir = Path(input_dir)

    if output_dir is None:
        output_dir = Path.cwd()
    else:
        output_dir = Path(output_dir)

    # 确保输出目录存在
    output_dir.mkdir(parents=True, exist_ok=True)

    # 失败日志路径
    failed_log = output_dir / "failed.log"

    # 如果未指定文件列表，扫描输入目录
    if file_list is None:
        supported_files = []
        for ext in ['*.docx', '*.pdf']:
            supported_files.extend(input_dir.glob(ext))
        files_to_process = [f.name for f in supported_files]
    else:
        files_to_process = file_list

    if not files_to_process:
        print("未找到需要处理的文件")
        return

    results = []
    failed_files = []

    for file_name in files_to_process:
        file_path = input_dir / file_name

        if not file_path.exists():
            print(f"[失败] 文件不存在: {file_path}")
            results.append((file_name, "文件不存在"))
            failed_files.append(f"{file_name}|文件不存在")
            continue

        try:
            # 提取文件内容
            if file_name.endswith('.docx'):
                file_content = extract_docx_content(file_path)
            elif file_name.endswith('.pdf'):
                file_content = extract_pdf_content(file_path)
            else:
                print(f"[失败] 不支持的文件格式: {file_name}")
                results.append((file_name, "不支持的格式"))
                failed_files.append(f"{file_name}|不支持的格式")
                continue

            # 用Claude转换
            markdown_content = convert_with_claude(file_name, file_content)

            # 保存输出
            output_file_name = file_name.rsplit('.', 1)[0] + '.md'
            output_path = output_dir / output_file_name

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)

            print(f"[成功] {output_file_name}")
            results.append((file_name, "成功"))

        except Exception as e:
            error_msg = str(e)
            print(f"[失败] {file_name}: {error_msg[:100]}")
            results.append((file_name, f"错误: {error_msg[:100]}"))
            failed_files.append(f"{file_name}|{error_msg}")

    # 写入失败日志
    if failed_files:
        with open(failed_log, 'a', encoding='utf-8') as f:
            f.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            for failed_entry in failed_files:
                f.write(failed_entry + "\n")

    # 输出总结
    print("\n" + "="*60)
    print("转换完成总结")
    print("="*60)
    success_count = sum(1 for _, status in results if status == "成功")
    print(f"成功: {success_count}/{len(results)}")
    if failed_files:
        print(f"失败: {len(failed_files)}/{ len(results)}")
        print(f"详见: {failed_log}")
    for file_name, status in results:
        if status != "成功":
            print(f"  {file_name}: {status}")


if __name__ == "__main__":
    # 检查环境变量
    if not os.getenv("ANTHROPIC_BASE_URL") or not os.getenv("ANTHROPIC_AUTH_TOKEN"):
        print("错误: 缺少必要的环境变量")
        print("\n请设置以下环境变量:")
        print("  ANTHROPIC_BASE_URL     - API服务地址（如: https://api.anthropic.com）")
        print("  ANTHROPIC_AUTH_TOKEN   - API认证密钥")
        print("\n设置方法:")
        if sys.platform == "win32":
            print("  set ANTHROPIC_BASE_URL=<your-url>")
            print("  set ANTHROPIC_AUTH_TOKEN=<your-token>")
        else:
            print("  export ANTHROPIC_BASE_URL=<your-url>")
            print("  export ANTHROPIC_AUTH_TOKEN=<your-token>")
        sys.exit(1)

    # 解析命令行参数
    import argparse
    parser = argparse.ArgumentParser(description='将DOCX和PDF文件转换为Markdown格式')
    parser.add_argument('--input', '-i', default=None, help='输入目录（默认为当前目录）')
    parser.add_argument('--output', '-o', default=None, help='输出目录（默认为当前目录）')
    parser.add_argument('files', nargs='*', help='要转换的文件列表（如果未指定，处理输入目录中的所有.docx和.pdf文件）')

    args = parser.parse_args()

    process_files(
        input_dir=args.input,
        output_dir=args.output,
        file_list=args.files if args.files else None
    )
