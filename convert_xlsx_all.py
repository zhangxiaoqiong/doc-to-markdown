#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel to Markdown converter - optimized for Dify knowledge base segmentation
Wrapper around convert_xlsx.py for single file processing
Supports generating separate files for each sheet
"""

import sys
import os
from pathlib import Path
from typing import Tuple, Optional

if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Import the main conversion function
from convert_xlsx import excel_to_markdown_by_sheets


def convert_xlsx_single_file(input_dir: Path, output_dir: Path, file_path: Path, enable_row_descriptions: bool = False) -> Tuple[bool, Optional[str]]:
    """Convert a single XLSX file to markdown, generating separate files for each sheet if needed."""
    try:
        if not file_path.exists():
            return False, f"文件不存在: {file_path}"

        sheet_markdowns = excel_to_markdown_by_sheets(str(file_path), enable_row_descriptions=enable_row_descriptions)

        if not sheet_markdowns:
            return False, "转换结果为空"

        # If only one sheet, use simple filename
        if len(sheet_markdowns) == 1:
            sheet_name, markdown_content = list(sheet_markdowns.items())[0]
            output_file = output_dir / (file_path.stem + '.md')
            output_file.write_text(markdown_content, encoding='utf-8')
        else:
            # Multiple sheets: generate separate files with sheet names
            for sheet_name, markdown_content in sheet_markdowns.items():
                # Sanitize sheet name for filename
                safe_sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace(':', '_')
                output_filename = f"{file_path.stem}_{safe_sheet_name}.md"
                output_file = output_dir / output_filename
                output_file.write_text(markdown_content, encoding='utf-8')

        return True, None

    except Exception as e:
        return False, f"转换失败: {str(e)}"


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert XLSX file to Markdown with optional row descriptions")
    parser.add_argument("input_dir", help="Input directory")
    parser.add_argument("output_dir", help="Output directory")
    parser.add_argument("file_path", help="Path to XLSX file (relative or absolute)")
    parser.add_argument(
        "--enable-row-descriptions",
        action="store_true",
        default=False,
        help="Enable LLM-generated descriptions for each row"
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)

    # Handle file_path (same as before)
    file_arg = args.file_path
    file_path = Path(file_arg)
    if not file_path.is_absolute():
        # Check if it's relative to input_dir
        full_path = input_dir / file_arg
        if full_path.exists():
            file_path = full_path

    success, error = convert_xlsx_single_file(
        input_dir,
        output_dir,
        file_path,
        enable_row_descriptions=args.enable_row_descriptions
    )

    if success:
        print(f"✓ Successfully converted: {file_arg}")
        sys.exit(0)
    else:
        print(f"✗ Error: {error}")
        sys.exit(1)
