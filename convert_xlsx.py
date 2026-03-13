"""Simple Excel to Markdown converter - core logic only."""

from pathlib import Path
import openpyxl


def excel_to_markdown(file_path: str) -> str:
    """Convert Excel to simple markdown.

    Reads all rows and formats them hierarchically.
    """
    workbook = openpyxl.load_workbook(file_path)
    markdown_lines = []

    try:
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            rows = list(sheet.iter_rows(values_only=True))

            if not rows:
                continue

            # Skip header row
            data_rows = rows[1:]
            current_category = None

            for row in data_rows:
                if not row or all(cell is None for cell in row):
                    continue

                # Extract columns - handle None values properly
                col0 = str(row[0]).strip() if row[0] else ""
                col1 = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                col2 = str(row[2]).strip() if len(row) > 2 and row[2] else ""

                if not col1 or not col2:  # Need at least 2 columns
                    continue

                # Add category heading
                if col0 and col0 != current_category:
                    markdown_lines.append(f"# {col0}")
                    current_category = col0

                # Add as question/item
                markdown_lines.append(f"## {col1}")
                markdown_lines.append(col2)
                markdown_lines.append("")
                markdown_lines.append("---")

    finally:
        workbook.close()

    return '\n'.join(markdown_lines)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python simple_converter.py <input.xlsx> <output.md>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    md = excel_to_markdown(input_file)
    Path(output_file).write_text(md, encoding='utf-8')
    print(f"✓ Generated: {output_file}")
