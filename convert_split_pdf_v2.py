#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分割处理超大PDF文件 - 更小的块
"""

import os
import sys
import io
import base64
import time
import random
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

try:
    from PIL import Image
except ImportError:
    print("需要安装 Pillow: pip install Pillow")
    sys.exit(1)


def split_pdf(file_path, pages_per_chunk=10, temp_dir=None):
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
            if temp_dir is None:
                temp_dir = Path(os.getenv("TEMP", Path.cwd() / "tmp"))
            temp_dir = Path(temp_dir)
            temp_dir.mkdir(parents=True, exist_ok=True)
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


def compress_pdf_page(page_data: bytes, max_size_mb: float = 4.5) -> bytes:
    """压缩PDF页面图片以满足API大小限制"""
    try:
        img = Image.open(io.BytesIO(page_data))

        # 转换为RGB
        if img.mode in ('RGBA', 'LA', 'P'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if 'A' in img.mode else None)
            img = background
        else:
            img = img.convert('RGB')

        # 压缩
        quality = 90
        max_size_bytes = int(max_size_mb * 1024 * 1024)

        while quality > 10:
            output = io.BytesIO()
            img.save(output, format='JPEG', quality=quality, optimize=True)
            compressed_data = output.getvalue()

            if len(compressed_data) <= max_size_bytes:
                return compressed_data

            quality -= 10

        # 二次缩放
        width, height = img.size
        while width > 1000 or height > 1000:
            scale = min(1000 / width, 1000 / height, 1.0)
            width = int(width * scale)
            height = int(height * scale)
            resized = img.resize((width, height), Image.Resampling.LANCZOS)
            output = io.BytesIO()
            resized.save(output, format='JPEG', quality=80, optimize=True)
            compressed_data = output.getvalue()
            if len(compressed_data) <= max_size_bytes:
                return compressed_data

        return page_data

    except Exception as e:
        print(f"      图片压缩异常: {str(e)}")
        return page_data


def pdf_to_base64(file_path):
    """将PDF文件转换为base64"""
    with open(file_path, 'rb') as f:
        return base64.standard_b64encode(f.read()).decode("utf-8")


def convert_pdf_pages_to_images(chunk_path, dpi=150):
    """将PDF块的每一页转换为图片，返回图片数据列表"""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(chunk_path)
        image_data_list = []
        for page in doc:
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat)
            buf = io.BytesIO()
            pix.pil_save(buf, format='JPEG', quality=90)
            image_data_list.append(buf.getvalue())
        doc.close()
        return image_data_list
    except Exception as e:
        print(f"PDF转图片失败: {str(e)}")
        return []


def convert_pdf_chunk(file_name, chunk_path, chunk_info, chunk_num, total_chunks, max_retries=3):
    """转换单个PDF块，支持重试 - 使用image类型代替document类型"""
    client = Anthropic(
        api_key=os.getenv("ANTHROPIC_AUTH_TOKEN"),
        base_url=os.getenv("ANTHROPIC_BASE_URL")
    )

    # 将PDF转为图片列表
    image_data_list = convert_pdf_pages_to_images(chunk_path)
    if not image_data_list:
        print("x (PDF转图片失败)")
        return None

    # 压缩每张图片
    compressed_images = []
    for img_data in image_data_list:
        compressed = compress_pdf_page(img_data)
        compressed_images.append(compressed)

    total_size_mb = sum(len(d) for d in compressed_images) / (1024 * 1024)
    print(f"[{chunk_num}/{total_chunks}] 第 {chunk_info['pages']} 页 ({len(compressed_images)}张图片, {total_size_mb:.2f}MB)... ", end='', flush=True)

    # 构建内容：文本prompt + 多张图片
    content = [
        {
            "type": "text",
            "text": f"""请将以下PDF页面（第 {chunk_info['pages']} 页）完整转换为Markdown。

文件名：{file_name}

这是文件的第 {chunk_info['pages']} 部分。请严格遵守以下要求：
1. 【忽略水印】：忽略页面边缘的OA系统打印痕迹（如日期、网址链接、页码等）。
2. 【提取结构】：保留所有标题、表格、列表等结构，图表用文字清晰描述。
3. 【⚠️ 防范错别字】：本文档是企业管理/财务制度，请仔细辨认中文字体，严禁将形近字认错！
   - 注意职务：如识别出类似"查事长"，须纠正为"董事长"或"副总裁"。
   - 注意词汇：是"综合"不是"统合"，是"统一"不是"核一"，是"管理"不是"管外"。
   - 注意业务词：是"单线/全线"而不是"眼线"。
4. 【模糊处理】：遇到手写签名或极其模糊的字，请结合上下文推测；若无法辨别，请直接写 `[字迹不清]`，绝不生造生僻词。

请直接输出Markdown，无需其他说明。"""
        }
    ]

    # 添加所有图片
    for img_data in compressed_images:
        img_base64 = base64.standard_b64encode(img_data).decode("utf-8")
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": img_base64
            }
        })

    # 重试逻辑
    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=20000,
                messages=[
                    {"role": "user", "content": content}
                ]
            )
            # 提取文本内容（兼容ThinkingBlock）
            result = ""
            for block in message.content:
                if hasattr(block, 'text'):
                    result = block.text
                    break
            print("OK")
            return result
        except Exception as e:
            error_str = str(e)
            # 判断是否可重试的错误
            if "429" in error_str or "500" in error_str or "503" in error_str:
                if attempt < max_retries - 1:
                    wait_time = (2 ** attempt) + random.uniform(0, 1)
                    print(f"\n  [重试 {attempt+1}/{max_retries-1}] 等待 {wait_time:.1f}s... ", end='', flush=True)
                    time.sleep(wait_time)
                    continue
            # 其他错误或最后一次重试失败
            print(f"x ({error_str[:80]})")
            return None


def clean_placeholder_content(text):
    """清理占位符内容和重复部分"""
    import re

    # 删除各种形式的占位符
    # 删除 [此部分内容在第X页] 形式
    text = re.sub(r'\n\s*\[\s*此部分内容在第\d+页\s*\]\s*\n', '\n', text)
    text = re.sub(r'^\s*\[\s*此部分内容在第\d+页\s*\]\s*$', '', text, flags=re.MULTILINE)

    # 删除 （本部分在原文档第X页） 形式
    text = re.sub(r'\n\s*\(\s*本部分在原文档第\d+页\s*\)\s*\n', '\n', text)
    text = re.sub(r'^\s*\(\s*本部分在原文档第\d+页\s*\)\s*$', '', text, flags=re.MULTILINE)

    # 删除 本部分内容在第X-Y页 形式（最新发现的）
    text = re.sub(r'\n\s*本部分内容在第\d+[\-～]\d+页\s*\n', '\n', text)
    text = re.sub(r'^\s*本部分内容在第\d+[\-～]\d+页\s*$', '', text, flags=re.MULTILINE)

    # 删除 本部分内容在第X页 形式
    text = re.sub(r'\n\s*本部分内容在第\d+页\s*\n', '\n', text)
    text = re.sub(r'^\s*本部分内容在第\d+页\s*$', '', text, flags=re.MULTILINE)

    # 清理空的章节（标题后面只有占位符或空行）
    text = re.sub(r'^### [^\n]*\n\s*\n(?=###|##|# )', '', text, flags=re.MULTILINE)
    text = re.sub(r'^## [^\n]*\n\s*\n(?=##|# )', '', text, flags=re.MULTILINE)

    # 处理重复的一级标题 - 保留第一个，删除之后出现的相同标题及其后续内容直到下一个不同的一级标题
    lines = text.split('\n')
    seen_h1 = {}
    cleaned_lines = []
    skip_mode = False
    skip_title = None

    for i, line in enumerate(lines):
        if line.startswith('# ') and not line.startswith('## '):
            title = line.strip()

            # 检查是否之前见过这个标题
            if title in seen_h1:
                # 进入跳过模式，跳过直到遇到新的一级标题
                skip_mode = True
                skip_title = title
                continue
            else:
                # 新的一级标题，记录并结束跳过模式
                seen_h1[title] = True
                skip_mode = False
                skip_title = None
                cleaned_lines.append(line)
        elif skip_mode and line.startswith('# ') and not line.startswith('## '):
            # 遇到新的一级标题，结束跳过
            skip_mode = False
            skip_title = None
            title = line.strip()
            if title not in seen_h1:
                seen_h1[title] = True
                cleaned_lines.append(line)
            else:
                skip_mode = True
                skip_title = title
        elif not skip_mode:
            cleaned_lines.append(line)

    text = '\n'.join(cleaned_lines)

    # 清理多余的空行（超过2个连续空行改为2个）
    text = re.sub(r'\n\n\n+', '\n\n', text)

    return text


def process_large_pdf(input_file, input_dir=None, output_dir=None, pages_per_chunk=5, min_size_mb=15, temp_dir=None):
    """处理大型PDF文件

    Args:
        input_file: 输入文件名
        input_dir: 输入目录路径（默认为当前目录）
        output_dir: 输出目录路径（默认为当前目录）
        pages_per_chunk: 每个块包含的页数（默认为5）
        min_size_mb: 仅当文件大于此大小（MB）时才进行分割（默认为15MB）
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

    file_path = input_dir / input_file

    if not file_path.exists():
        print(f"文件不存在: {file_path}")
        return

    print(f"处理: {input_file}\n")

    # 检查文件大小
    file_size_mb = file_path.stat().st_size / (1024 * 1024)

    if file_size_mb < min_size_mb:
        print(f"WARNING: 文件大小仅 {file_size_mb:.1f}MB，小于 {min_size_mb}MB 的分割阈值")
        print(f"跳过分割，直接用Vision API处理整个文件\n")

        # 将PDF转为图片后用image类型处理
        image_data_list = convert_pdf_pages_to_images(file_path)
        if not image_data_list:
            print("PDF转图片失败")
            return

        # 压缩每张图片
        compressed_images = []
        for img_data in image_data_list:
            compressed_images.append(compress_pdf_page(img_data))

        total_size_mb = sum(len(d) for d in compressed_images) / (1024 * 1024)
        print(f"共 {len(compressed_images)} 张图片 ({total_size_mb:.2f}MB)\n")

        content = [
            {
                "type": "text",
                "text": f"""请将以下PDF文档完整转换为Markdown格式。

文件名：{input_file}

请严格遵守以下要求：
1. 【忽略水印】：忽略页面边缘的OA系统打印痕迹（如日期、系统链接、页码等）。
2. 【提取结构】：保留完整的文档结构，所有表格转换为Markdown表格。
3. 【⚠️ 防范错别字】：本文档涉及企业管理制度，请仔细辨认扫描件中的中文，严禁写错形近字！
   - 注意职务：类似"查事长"大概率是"董事长"。
   - 注意词汇："综合"易被看成"统合"，"统一"易被看成"核一"，"管理"易看成"管外"。
   - 注意常识：类似"眼线"等不符合业务语境的词，须纠正为"单线/主线"等。
4. 【模糊处理】：遇到无法辨认的模糊手写字迹，请用 `[字迹不清]` 标记，不要生造词。

请直接输出完整的Markdown内容，无需多余解释。"""
            }
        ]

        # 添加所有图片
        for img_data in compressed_images:
            img_base64 = base64.standard_b64encode(img_data).decode("utf-8")
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/jpeg",
                    "data": img_base64
                }
            })

        client = Anthropic(
            api_key=os.getenv("ANTHROPIC_AUTH_TOKEN"),
            base_url=os.getenv("ANTHROPIC_BASE_URL")
        )

        print("使用Vision API处理整个文件...")
        for attempt in range(3):
            try:
                message = client.messages.create(
                    model="claude-opus-4-6",
                    max_tokens=20000,
                    messages=[
                        {"role": "user", "content": content}
                    ]
                )
                # 提取文本内容（兼容ThinkingBlock）
                markdown_content = ""
                for block in message.content:
                    if hasattr(block, 'text'):
                        markdown_content = block.text
                        break
                break
            except Exception as e:
                error_str = str(e)
                if "429" in error_str or "500" in error_str or "503" in error_str:
                    if attempt < 2:
                        wait_time = 2 ** attempt
                        print(f"遇到 {error_str.split()[0]} 错误，等待 {wait_time}s 后重试...")
                        time.sleep(wait_time)
                        continue
                raise

        output_name = input_file.rsplit('.', 1)[0] + '.md'
        output_path = output_dir / output_name

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        print(f"✓ 已保存: {output_name}")
        print(f"文件大小: {len(markdown_content)/1024:.1f}KB")
        return

    # 文件大于阈值，进行分割处理
    print(f"✓ 文件大小 {file_size_mb:.1f}MB，开始分割处理\n")
    print("第1步: 分割PDF...")
    chunks = split_pdf(file_path, pages_per_chunk=pages_per_chunk, temp_dir=temp_dir)
    if not chunks:
        print("分割失败")
        return

    print(f"\n已分割为 {len(chunks)} 个部分\n")

    # 转换每个块
    print("第2步: 逐块转换...")
    all_content = []
    success_count = 0
    failed_chunks = []

    for idx, chunk in enumerate(chunks, 1):
        markdown_content = convert_pdf_chunk(input_file, chunk['path'], chunk, idx, len(chunks))

        if markdown_content:
            all_content.append(markdown_content)
            all_content.append("\n\n")
            success_count += 1
        else:
            failed_chunks.append(f"第 {chunk['pages']} 页")

        # 避免API限流，稍微延迟
        if idx < len(chunks):
            time.sleep(1)

    print(f"\n已成功转换: {success_count}/{len(chunks)} 个部分\n")

    if failed_chunks:
        print(f"⚠️  以下块转换失败: {', '.join(failed_chunks)}")

    # 合并内容
    print("第3步: 合并内容...")
    final_content = "".join(all_content)

    # 清理占位符
    print("第4步: 清理占位符...")
    final_content = clean_placeholder_content(final_content)

    # 保存文件
    output_name = input_file.rsplit('.', 1)[0] + '.md'
    output_path = output_dir / output_name

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_content)

    print(f"已保存: {output_name}")
    print(f"文件大小: {len(final_content)/1024:.1f}KB")

    # 清理临时文件
    print("第5步: 清理临时文件...")
    for chunk in chunks:
        try:
            os.remove(chunk['path'])
        except:
            pass

    # 记录失败块信息到日志
    if failed_chunks:
        failed_log = output_dir / "failed.log"
        with open(failed_log, 'a', encoding='utf-8') as f:
            f.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            f.write(f"文件: {input_file}\n")
            for chunk_info in failed_chunks:
                f.write(f"  失败: {chunk_info}\n")


if __name__ == "__main__":
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

    import argparse
    parser = argparse.ArgumentParser(description='将大型PDF文件分割后转换为Markdown格式')
    parser.add_argument('input_file', help='输入PDF文件名')
    parser.add_argument('--input', '-i', default=None, help='输入目录（默认为当前目录）')
    parser.add_argument('--output', '-o', default=None, help='输出目录（默认为当前目录）')
    parser.add_argument('--pages-per-chunk', '-p', type=int, default=5, help='每个块包含的页数（默认为5）')
    parser.add_argument('--min-size', '-m', type=float, default=15, help='仅当文件大于此大小(MB)时才分割（默认为15）')
    parser.add_argument('--temp-dir', '-t', default=None, help='临时文件目录（默认使用系统TEMP或当前目录/tmp）')

    args = parser.parse_args()

    process_large_pdf(
        input_file=args.input_file,
        input_dir=args.input,
        output_dir=args.output,
        pages_per_chunk=args.pages_per_chunk,
        min_size_mb=args.min_size,
        temp_dir=args.temp_dir
    )
