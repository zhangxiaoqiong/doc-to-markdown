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

这是一份丰图科技的财务管理文档。请：
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


def process_pdfs():
    """处理所有PDF文件"""
    base_dir = Path("知识库base")
    output_dir = Path("知识库md")

    output_dir.mkdir(exist_ok=True)

    pdf_files = [
        "丰图科技员工费用报销操作指引V3.0 (1).pdf",
        "丰图科技客户评级管理规则【1.0】.pdf",
        "关于项目投入及费用结算的管理要求及标准.pdf",
        "员工报销管理规定.pdf",
    ]

    results = []

    for file_name in pdf_files:
        file_path = base_dir / file_name

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
    
    process_pdfs()
