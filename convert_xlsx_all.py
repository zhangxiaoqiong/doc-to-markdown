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


def convert_xlsx_single_file(input_dir: Path, output_dir: Path, file_path: Path) -> Tuple[bool, Optional[str]]:
    """Convert a single XLSX file to markdown, generating separate files for each sheet if needed."""
    try:
        if not file_path.exists():
            return False, f"文件不存在: {file_path}"

        sheet_markdowns = excel_to_markdown_by_sheets(str(file_path))

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
    if len(sys.argv) < 3:
        print("Usage: python convert_xlsx_all.py <input_dir> <output_dir> [file_path]")
        sys.exit(1)

    input_dir = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])

    # If file_path provided, use it directly (it may be relative path or absolute)
    if len(sys.argv) >= 4:
        file_arg = sys.argv[3]
        # If it's a relative path within input_dir, construct full path
        file_path = Path(file_arg)
        if not file_path.is_absolute():
            # Check if it's relative to input_dir
            full_path = input_dir / file_arg
            if full_path.exists():
                file_path = full_path
    else:
        print("Error: file_path is required")
        sys.exit(1)

    success, error = convert_xlsx_single_file(input_dir, output_dir, file_path)

    if success:
        print(f"✓ Successfully converted: {file_arg}")
        sys.exit(0)
    else:
        print(f"✗ Error: {error}")
        sys.exit(1)
