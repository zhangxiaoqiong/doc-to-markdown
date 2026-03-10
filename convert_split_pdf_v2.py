#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分割处理超大PDF文件 - 更小的块
"""

import os
import sys
import base64
import time
from pathlib import Path
from anthropic import Anthropic

if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("需要安装 pypdf: pip install pypdf")
    sys.exit(1)


def split_pdf(file_path, pages_per_chunk=10):
    """将PDF分割成更小的部分"""
    try:
        reader = PdfReader(file_path)
        total_pages = len(reader.pages)
        print(f"总页数: {total_pages}\n")
        
        chunks = []
        for start_page in range(0, total_pages, pages_per_chunk):
            end_page = min(start_page + pages_per_chunk, total_pages)
            
            writer = PdfWriter()
            for page_num in range(start_page, end_page):
                writer.add_page(reader.pages[page_num])
            
            # 保存临时PDF
            temp_dir = Path("/tmp")
            temp_dir.mkdir(exist_ok=True)
            temp_path = temp_dir / f"chunk_{start_page:03d}_{end_page:03d}.pdf"
            
            with open(temp_path, 'wb') as f:
                writer.write(f)
            
            chunks.append({
                'path': str(temp_path),
                'pages': f"{start_page+1}-{end_page}",
                'start': start_page + 1,
                'end': end_page
            })
            print(f"  已分割: 第 {start_page+1:3d}-{end_page:3d} 页 ({(end_page-start_page)} 页)")
        
        return chunks
    except Exception as e:
        print(f"分割失败: {str(e)}")
        return []


def pdf_to_base64(file_path):
    """将PDF文件转换为base64"""
    with open(file_path, 'rb') as f:
        return base64.standard_b64encode(f.read()).decode("utf-8")


def convert_pdf_chunk(file_name, chunk_path, chunk_info, chunk_num, total_chunks):
    """转换单个PDF块"""
    client = Anthropic(
        api_key=os.getenv("ANTHROPIC_AUTH_TOKEN"),
        base_url=os.getenv("ANTHROPIC_BASE_URL")
    )
    
    pdf_base64 = pdf_to_base64(chunk_path)
    file_size_mb = len(pdf_base64) / (1024 * 1024)
    
    print(f"[{chunk_num}/{total_chunks}] 第 {chunk_info['pages']} 页 ({file_size_mb:.2f}MB)... ", end='', flush=True)
    
    content = [
        {
            "type": "text",
            "text": f"""请将以下PDF片段（第 {chunk_info['pages']} 页）完整转换为Markdown。

文件名：{file_name}

这是员工报销管理规定的第 {chunk_info['pages']} 部分。请：
1. 提取该部分所有文本内容
2. 保留所有标题、表格、列表等结构
3. 保留数据和关键信息
4. 如有图表，用文字描述

请直接输出Markdown，无需其他说明。"""
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
    
    try:
        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4096,
            messages=[
                {"role": "user", "content": content}
            ]
        )
        result = message.content[0].text
        print("✓")
        return result
    except Exception as e:
        error_msg = str(e)[:100]
        print(f"✗ ({error_msg})")
        return None


def process_large_pdf():
    """处理超大PDF"""
    base_dir = Path("知识库base")
    output_dir = Path("知识库md")
    output_dir.mkdir(exist_ok=True)
    
    file_name = "员工报销管理规定.pdf"
    file_path = base_dir / file_name
    
    if not file_path.exists():
        print(f"文件不存在: {file_path}")
        return
    
    print(f"处理: {file_name}\n")
    
    # 分割PDF
    print("第1步: 分割PDF...")
    chunks = split_pdf(file_path, pages_per_chunk=5)
    if not chunks:
        print("分割失败")
        return
    
    print(f"\n已分割为 {len(chunks)} 个部分\n")
    
    # 转换每个块
    print("第2步: 逐块转换...")
    all_content = []
    success_count = 0
    
    for idx, chunk in enumerate(chunks, 1):
        markdown_content = convert_pdf_chunk(file_name, chunk['path'], chunk, idx, len(chunks))
        
        if markdown_content:
            all_content.append(markdown_content)
            all_content.append("\n\n")
            success_count += 1
        
        # 避免API限流，稍微延迟
        if idx < len(chunks):
            time.sleep(1)
    
    print(f"\n已成功转换: {success_count}/{len(chunks)} 个部分\n")
    
    # 合并内容
    print("第3步: 合并内容...")
    final_content = "# 员工报销管理规定\n\n"
    final_content += "".join(all_content)
    
    # 保存文件
    output_name = file_name.rsplit('.', 1)[0] + '.md'
    output_path = output_dir / output_name
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_content)
    
    print(f"已保存: {output_name}")
    print(f"文件大小: {len(final_content)/1024:.1f}KB")
    
    # 清理临时文件
    for chunk in chunks:
        try:
            os.remove(chunk['path'])
        except:
            pass


if __name__ == "__main__":
    if not os.getenv("ANTHROPIC_BASE_URL") or not os.getenv("ANTHROPIC_AUTH_TOKEN"):
        print("错误: 未设置环境变量")
        sys.exit(1)
    
    process_large_pdf()
