#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合文档转换程序（文件级串行处理）
每个文件完整走完4步流程，然后处理下一个文件
- 步骤1：提取内容 + Claude转换
- 步骤2：检测质量，判断是否需要Vision
- 步骤3：Vision重处理（如需）
- 步骤4：Claude校对OCR错误
"""

import os
import sys
import subprocess
from pathlib import Path
import re
from datetime import datetime

# Windows编码
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def get_failed_log_path(output_dir):
    """获取失败日志文件路径"""
    return Path(output_dir) / "failed.log"


def is_in_failed_log(output_dir, file_name):
    """检查文件是否在failed.log中"""
    failed_log = get_failed_log_path(output_dir)
    if not failed_log.exists():
        return False

    try:
        with open(failed_log, 'r', encoding='utf-8') as f:
            content = f.read()
            return file_name in content
    except:
        return False


def is_file_complete(output_dir, source_file_path):
    """
    判断文件是否已完整处理过
    标准：输出文件存在 + 不在failed.log中
    """
    file_stem = source_file_path.stem
    output_file = Path(output_dir) / (file_stem + '.md')

    # 输出文件存在 且 不在failed.log中 → 认为已完成
    if output_file.exists() and not is_in_failed_log(output_dir, source_file_path.name):
        return True
    return False


def log_failed(output_dir, file_name, error_msg):
    """记录失败的文件到failed.log"""
    failed_log = get_failed_log_path(output_dir)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(failed_log, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {file_name} | {error_msg}\n")


def remove_from_failed_log(output_dir, file_name):
    """从failed.log中移除文件（处理成功后）"""
    failed_log = get_failed_log_path(output_dir)
    if not failed_log.exists():
        return

    try:
        with open(failed_log, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        # 过滤出不包含该文件名的行
        filtered_lines = [line for line in lines if file_name not in line]

        with open(failed_log, 'w', encoding='utf-8') as f:
            f.writelines(filtered_lines)
    except:
        pass


def run_convert_docs_single_file(input_dir, output_dir, file_path):
    """处理单个文件：提取内容 + Claude转换

    returns: (success: bool, error_msg: str or None)
    """
    # 运行convert_docs.py处理该文件
    cmd = [
        sys.executable,
        "convert_docs.py",
        "--input", str(input_dir),
        "--output", str(output_dir)
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    # 检查输出中是否包含成功标记
    if result.returncode == 0:
        return True, None
    else:
        # 提取错误信息
        error_msg = result.stderr if result.stderr else result.stdout
        return False, error_msg


def check_needs_vision_single_file(input_dir, output_dir, file_path):
    """检查单个文件是否需要用Vision API重新处理

    returns: (needs_vision: bool, reason: str or None)
    """
    file_stem = file_path.stem
    md_file = Path(output_dir) / (file_stem + '.md')

    if not md_file.exists():
        return False, None

    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 判断1：PDF文件一定需要用Vision重处理（保留图片和排版）
    if file_path.suffix.lower() == '.pdf':
        return True, 'PDF文件需要Vision保留图片和排版'

    # 判断2：DOCX中是否有流程/图片内容但缺少解析
    if file_path.suffix.lower() == '.docx':
        # 检查是否提到流程、图、图表等
        has_visual_mention = bool(re.search(r'(流程|图|图表|图片|示意|二维码)', content))

        # 检查是否有流程图的解析迹象
        has_flow_parsed = bool(re.search(r'(→|↓|├|┌|流程图详解|步骤\d+|操作步骤)', content))

        if has_visual_mention and not has_flow_parsed:
            return True, '含有流程/图片内容但缺少Vision解析'

    return False, None


def run_convert_pdf_vision_single_file(input_dir, output_dir, file_path):
    """对单个PDF或DOCX文件用Vision API重新处理

    returns: (success: bool, error_msg: str or None)
    """
    if file_path.suffix.lower() == '.pdf':
        # 用convert_split_pdf_v2处理PDF
        file_name = file_path.stem

        cmd = [
            sys.executable,
            "convert_split_pdf_v2.py",
            file_path.name,
            "--input", str(input_dir),
            "--output", str(output_dir)
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            return True, None
        else:
            error_msg = result.stderr if result.stderr else result.stdout
            return False, error_msg

    elif file_path.suffix.lower() == '.docx':
        # DOCX文件无法自动处理，提示用户
        return False, "DOCX文件含有流程/图片，无法自动处理。建议用Word打开手工检查。"

    return False, "未知文件类型"


def run_fix_markdown_single_file(output_dir, file_path):
    """校对单个Markdown文件的OCR错误

    returns: (success: bool, error_msg: str or None)
    """
    file_stem = file_path.stem
    md_file = Path(output_dir) / (file_stem + '.md')

    if not md_file.exists():
        return True, None  # 文件不存在，跳过

    cmd = [
        sys.executable,
        "fix_markdown_with_claude.py",
        "--file", str(md_file)
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode == 0:
        return True, None
    else:
        error_msg = result.stderr if result.stderr else result.stdout
        return False, error_msg


def process_single_file(input_dir, output_dir, file_path):
    """处理单个文件的完整流程（4步）

    returns: (success: bool, error_msg: str or None)
    """
    file_name = file_path.name

    print(f"\n  [步骤1] 转换: {file_name}")
    success, error = run_convert_docs_single_file(input_dir, output_dir, file_path)
    if not success:
        return False, f"步骤1转换失败: {error}"

    print(f"  [步骤2] 检测: {file_name}")
    needs_vision, reason = check_needs_vision_single_file(input_dir, output_dir, file_path)

    if needs_vision:
        print(f"  [步骤3] Vision重处理: {reason}")
        success, error = run_convert_pdf_vision_single_file(input_dir, output_dir, file_path)
        if not success:
            return False, f"步骤3Vision处理失败: {error}"

    print(f"  [步骤4] 校对OCR: {file_name}")
    success, error = run_fix_markdown_single_file(output_dir, file_path)
    if not success:
        return False, f"步骤4校对失败: {error}"

    return True, None


def main():
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

    import argparse
    parser = argparse.ArgumentParser(
        description='综合文档转换程序：文件级串行处理，每个文件完整走完4步流程'
    )
    parser.add_argument('--input', '-i', default='知识库base', help='输入目录')
    parser.add_argument('--output', '-o', default='知识库md_v1.0', help='输出目录')

    args = parser.parse_args()

    input_dir = Path(args.input)
    output_dir = Path(args.output)

    if not input_dir.exists():
        print(f"错误：输入目录不存在 {input_dir}")
        sys.exit(1)

    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)

    # 获取所有源文件（DOCX和PDF）
    source_files = sorted([
        f for f in input_dir.glob('*')
        if f.suffix.lower() in ['.pdf', '.docx']
    ])

    if not source_files:
        print(f"错误：没有找到PDF或DOCX文件在 {input_dir}")
        sys.exit(1)

    print("="*60)
    print(f"开始处理 {len(source_files)} 个文件（文件级串行）")
    print("="*60)

    completed = 0
    skipped = 0
    failed = 0

    for file_path in source_files:
        file_name = file_path.name

        # 检查文件是否已完整处理
        if is_file_complete(output_dir, file_path):
            print(f"\n[跳过] {file_name} 已处理")
            skipped += 1
            continue

        # 处理文件
        print(f"\n[处理] {file_name}")
        success, error_msg = process_single_file(input_dir, output_dir, file_path)

        if success:
            print(f"[完成] {file_name}")
            remove_from_failed_log(output_dir, file_name)
            completed += 1
        else:
            print(f"[失败] {file_name}: {error_msg}")
            log_failed(output_dir, file_name, error_msg)
            failed += 1

    # 统计结果
    print("\n" + "="*60)
    print("✓ 转换完成！")
    print(f"  完成: {completed}")
    print(f"  跳过: {skipped}")
    print(f"  失败: {failed}")
    if failed > 0:
        print(f"\n✗ {failed} 个文件处理失败，详见 {get_failed_log_path(output_dir)}")
        sys.exit(1)
    print("="*60)


if __name__ == "__main__":
    main()
