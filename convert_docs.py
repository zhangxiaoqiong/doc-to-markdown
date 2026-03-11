#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务文档转换程序
使用Claude API将DOCX和PDF文件转换为Markdown格式（保持图文表绝对顺序）
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
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.text.paragraph import Paragraph
    from docx.table import Table
except ImportError:
    print("需要安装 python-docx: pip install python-docx")
    sys.exit(1)

try:
    import pdfplumber
except ImportError:
    print("需要安装 pdfplumber: pip install pdfplumber")
    sys.exit(1)


def extract_docx_content(file_path, image_dir):
    """从DOCX文件提取内容，严格保持文字、图片、表格的原有顺序"""
    doc = Document(file_path)
    content = []
    image_counter = 1

    # 遍历文档的每一个块级元素（段落和表格），保证从上到下的顺序
    for child in doc.element.body.iterchildren():
        
        # 1. 如果是段落（包含文字和行内图片）
        if isinstance(child, CT_P):
            para = Paragraph(child, doc)
            
            # 优先提取该段落中嵌入的图片
            for blip in child.xpath('.//a:blip'):
                # 获取图片关联ID
                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if rId and rId in doc.part.rels:
                    try:
                        image_part = doc.part.rels[rId].target_part
                        ext = image_part.content_type.split('/')[-1]
                        if ext == 'jpeg': ext = 'jpg'
                        
                        img_name = f"image_{image_counter}.{ext}"
                        img_path = os.path.join(image_dir, img_name)
                        
                        # 保存图片
                        with open(img_path, "wb") as f:
                            f.write(image_part.blob)
                            
                        # 在当前位置插入图片占位符
                        content.append(f"\n[IMAGE: {img_name}]\n")
                        image_counter += 1
                    except Exception as e:
                        print(f"\n  [警告] DOCX图片提取失败: {e}")
            
            # 提取段落文本
            if para.text.strip():
                content.append(para.text)

        # 2. 如果是表格
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            content.append("\n[TABLE]")
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                content.append(" | ".join(row_data))
            content.append("[/TABLE]\n")

    return "\n".join(content)


def extract_pdf_content(file_path, image_dir):
    """从PDF文件提取文本内容、表格和图片"""
    content = []
    image_counter = 1

    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            content.append(f"--- Page {page_num} ---\n")

            # PDF受限于格式，依然按 文本 -> 图片 -> 表格 的顺序输出每页内容
            # 1. 提取文本
            text = page.extract_text()
            if text:
                content.append(text)

            # 2. 提取图片
            if page.images:
                for img in page.images:
                    try:
                        if img["width"] < 50 or img["height"] < 50:
                            continue
                            
                        bbox = (img["x0"], img["top"], img["x1"], img["bottom"])
                        img_name = f"page{page_num}_img{image_counter}.png"
                        img_path = os.path.join(image_dir, img_name)
                        
                        cropped = page.crop(bbox)
                        img_obj = cropped.to_image(resolution=400)
                        img_obj.save(img_path, format="PNG")
                        
                        content.append(f"\n[IMAGE: {img_name}]\n")
                        image_counter += 1
                    except Exception:
                        pass 

            # 3. 提取表格
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    content.append("\n[TABLE]")
                    for row in table:
                        row_data = [str(cell) if cell else "" for cell in row]
                        content.append(" | ".join(row_data))
                    content.append("[/TABLE]\n")

    return "\n".join(content)


def convert_with_claude(file_name, file_content, file_base_name, max_retries=3):
    """使用Claude API将文档内容转换为Markdown"""
    base_url = os.getenv("ANTHROPIC_BASE_URL")
    auth_token = os.getenv("ANTHROPIC_AUTH_TOKEN")

    client = Anthropic(
        api_key=auth_token,
        base_url=base_url
    )

    system_prompt = f"""你是一个专业的文档转换专家。你的任务是将用户提供的文档内容转换为结构清晰的Markdown格式。

转换要求：
1. 保留原文档的层级结构（标题、小节、段落等）
2. 使用适当的Markdown语法（# 为一级标题，## 为二级标题，等等）
3. 将[TABLE]...[/TABLE]标记之间的内容转换为Markdown表格格式
4. 遇到 [IMAGE: filename.png] 占位符时，请原封不动地在它所在的位置将其转换为Markdown图片语法：
![图片描述](<assets/{file_base_name}/filename.png>))
5. 保留所有原始数据、日期、数字、金额等关键信息，保证图文表的顺序完全对应原文档。
6. 确保内容完整，不删除或篡改任何信息
7. 使用清晰的分段和格式，提高可读性

你的输出应该是完整的、可以直接保存为.md文件的Markdown内容。"""

    user_message = f"""请将以下文档内容转换为Markdown格式。

文件名：{file_name}

文档内容：{file_content}

请直接输出转换后的Markdown内容，不需要添加任何说明或标记。"""

    print(f"正在转换 {file_name}...", end=" ", flush=True)

    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=20000,
                messages=[
                    {"role": "user", "content": user_message}
                ],
                system=system_prompt
            )
            print("✓")
            return message.content[0].text
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "500" in error_str or "503" in error_str:
                if attempt < max_retries - 1:
                    wait_time = 2 ** attempt
                    print(f"\n  [重试 {attempt+1}/{max_retries-1}] 等待 {wait_time}s... ", end="", flush=True)
                    time.sleep(wait_time)
                    continue
            raise


def process_files(input_dir=None, output_dir=None, file_list=None):
    if input_dir is None:
        input_dir = Path.cwd()
    else:
        input_dir = Path(input_dir)

    if output_dir is None:
        output_dir = Path.cwd()
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)
    failed_log = output_dir / "failed.log"

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
        file_base_name = file_name.rsplit('.', 1)[0]
        
        image_dir = output_dir / "assets" / file_base_name

        if not file_path.exists():
            print(f"[失败] 文件不存在: {file_path}")
            results.append((file_name, "文件不存在"))
            continue

        try:
            if file_name.endswith('.docx'):
                image_dir.mkdir(parents=True, exist_ok=True)
                file_content = extract_docx_content(file_path, image_dir)
            elif file_name.endswith('.pdf'):
                image_dir.mkdir(parents=True, exist_ok=True)
                file_content = extract_pdf_content(file_path, image_dir)
            else:
                print(f"[失败] 不支持的文件格式: {file_name}")
                continue

            if image_dir.exists() and not any(image_dir.iterdir()):
                image_dir.rmdir()

            markdown_content = convert_with_claude(file_name, file_content, file_base_name)

            output_file_name = f"{file_base_name}.md"
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

    print("\n" + "="*60)
    print("转换完成总结")
    print("="*60)
    success_count = sum(1 for _, status in results if status == "成功")
    print(f"成功: {success_count}/{len(results)}")
    for file_name, status in results:
        if status != "成功":
            print(f"  {file_name}: {status}")


if __name__ == "__main__":
    if not os.getenv("ANTHROPIC_BASE_URL") or not os.getenv("ANTHROPIC_AUTH_TOKEN"):
        print("错误: 缺少必要的环境变量")
        sys.exit(1)

    import argparse
    parser = argparse.ArgumentParser(description='将DOCX和PDF文件转换为Markdown格式')
    parser.add_argument('--input', '-i', default=None, help='输入目录')
    parser.add_argument('--output', '-o', default=None, help='输出目录')
    parser.add_argument('files', nargs='*', help='要转换的文件列表')
    args = parser.parse_args()

    process_files(
        input_dir=args.input,
        output_dir=args.output,
        file_list=args.files if args.files else None
    )