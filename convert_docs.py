#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务文档转换程序
使用Claude API将DOCX和PDF文件转换为Markdown格式
"""

import os
import sys
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


def convert_with_claude(file_name, file_content):
    """使用Claude API将文档内容转换为Markdown"""
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

    print(f"正在转换 {file_name}...")

    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        messages=[
            {"role": "user", "content": user_message}
        ],
        system=system_prompt
    )

    return message.content[0].text


def process_files():
    """处理所有需要转换的文件"""
    base_dir = Path("知识库base")
    output_dir = Path("知识库md")

    # 确保输出目录存在
    output_dir.mkdir(exist_ok=True)

    # 定义需要处理的文件
    files_to_process = [
        "3-丰图科技：备用金及个人借款管理规范V4.0.docx",
        "6-应付结算例外事项管理办法【3.0】.docx",
        "7-丰图科技预付管理规定V3.0版20250307.docx",
        "丰图科技员工费用报销操作指引V3.0 (1).pdf",
        "丰图科技客户评级管理规则【1.0】.pdf",
        "关于项目投入及费用结算的管理要求及标准.pdf",
        "员工报销管理规定.pdf",
    ]

    results = []

    for file_name in files_to_process:
        file_path = base_dir / file_name

        if not file_path.exists():
            print(f"[失败] 文件不存在: {file_path}")
            results.append((file_name, "文件不存在"))
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
            print(f"[失败] {file_name}: {str(e)}")
            results.append((file_name, f"错误: {str(e)}"))

    # 输出总结
    print("\n" + "="*60)
    print("转换完成总结")
    print("="*60)
    success_count = sum(1 for _, status in results if status == "成功")
    print(f"成功: {success_count}/{len(results)}")
    for file_name, status in results:
        print(f"  {file_name}: {status}")


if __name__ == "__main__":
    # 检查环境变量
    if not os.getenv("ANTHROPIC_BASE_URL") or not os.getenv("ANTHROPIC_AUTH_TOKEN"):
        print("错误: 未设置ANTHROPIC_BASE_URL或ANTHROPIC_AUTH_TOKEN环境变量")
        sys.exit(1)

    process_files()
