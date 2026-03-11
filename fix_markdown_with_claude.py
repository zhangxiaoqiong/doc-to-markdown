#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用Claude作为校对员，检查和修复Markdown中的OCR错误
基于财务领域知识进行纠正
"""

import os
import sys
import time
from pathlib import Path
from anthropic import Anthropic

# Windows编码
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def fix_markdown_content(file_path, max_retries=3):
    """用Claude检查并修复Markdown中的OCR错误，支持重试"""

    # 读取Markdown文件
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 初始化Claude客户端
    client = Anthropic(
        api_key=os.getenv("ANTHROPIC_AUTH_TOKEN"),
        base_url=os.getenv("ANTHROPIC_BASE_URL")
    )

    # 构建修复提示词
    system_prompt = """这是一份企业财务相关的规范文档，已转换为Markdown。
检查并修复明显的OCR识别错误。
不确定的地方保持原样。
输出修复后的完整Markdown，无需解释。"""

    user_message = f"""请检查并修复以下Markdown文档中的OCR识别错误：

```markdown
{content}
```

输出修复后的完整Markdown内容。"""

    print(f"正在用Claude校对: {Path(file_path).name}...", end=" ", flush=True)

    # 重试逻辑
    for attempt in range(max_retries):
        try:
            # 调用Claude API
            message = client.messages.create(
                model="claude-opus-4-6",
                max_tokens=8192,
                messages=[
                    {"role": "user", "content": user_message}
                ],
                system=system_prompt
            )

            fixed_content = message.content[0].text

            # 保存修复后的内容
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(fixed_content)

            print("✓")
            return fixed_content
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "500" in error_str or "503" in error_str:
                if attempt < max_retries - 1:
                    wait_time = 2 ** attempt
                    print(f"\n  [重试 {attempt+1}/{max_retries-1}] 等待 {wait_time}s... ", end="", flush=True)
                    time.sleep(wait_time)
                    continue
            # 其他错误或最后一次重试失败
            raise


def fix_all_markdown_files(output_dir):
    """修复输出目录中的所有Markdown文件"""
    print("="*60)
    print("用Claude校对所有Markdown文件")
    print("="*60)

    output_dir = Path(output_dir)
    md_files = list(output_dir.glob('*.md'))

    if not md_files:
        print("未找到Markdown文件")
        return

    print(f"\n共找到 {len(md_files)} 个Markdown文件\n")

    failed_files = []

    for idx, md_file in enumerate(md_files, 1):
        try:
            fix_markdown_content(md_file)
            if idx < len(md_files):
                # 避免API限流
                time.sleep(2)
        except Exception as e:
            print(f"✗")
            print(f"  错误: {str(e)[:100]}\n")
            failed_files.append((md_file.name, str(e)))

    print("\n" + "="*60)
    print(f"✓ 完成校对 {len(md_files)} 个文件")
    if failed_files:
        print(f"⚠️  失败: {len(failed_files)} 个")
        # 写入失败日志
        failed_log = output_dir / "failed.log"
        with open(failed_log, 'a', encoding='utf-8') as f:
            f.write(f"\n=== {time.strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            for file_name, error in failed_files:
                f.write(f"{file_name}|{error}\n")
        print(f"详见: {failed_log}")
    print("="*60)


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

    import argparse
    parser = argparse.ArgumentParser(description='用Claude校对Markdown文档中的OCR错误')
    parser.add_argument('--dir', '-d', default='知识库md_v1.0', help='Markdown文件所在目录')

    args = parser.parse_args()

    fix_all_markdown_files(args.dir)
