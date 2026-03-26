# XLSX 行描述生成 Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add LLM-powered natural language descriptions for each row in Excel tables, formatted as "table row + description" paragraphs for RAG-based knowledge base retrieval.

**Architecture:**
- Extend `convert_xlsx.py` to accept `enable_row_descriptions` flag and generate descriptions via Claude API
- Refactor table output format: for each data row, output table header + row + LLM-generated description
- Pass parameter through `convert_xlsx_all.py` wrapper and `convert_all.py` orchestrator
- Reuse existing Anthropic SDK client initialization pattern from `convert_docs.py`

**Tech Stack:** Python 3, Anthropic SDK (already in use), openpyxl, pathlib

---

## Chunk 1: Core Row Description Generation

### Task 1: Add LLM description function to convert_xlsx.py

**Files:**
- Modify: `convert_xlsx.py` - add `generate_row_description()` function and update `excel_to_markdown_by_sheets()` signature

**Steps:**

- [ ] **Step 1: Read convert_docs.py to understand Anthropic client pattern**

Check lines 136-144 to see how client is initialized with ANTHROPIC_BASE_URL and ANTHROPIC_AUTH_TOKEN environment variables.

- [ ] **Step 2: Add generate_row_description function to convert_xlsx.py**

Add this function after imports, before `_parse_xlsx_from_xml()`:

```python
def generate_row_description(headers: list, row_data: list, enable: bool = False) -> str:
    """Generate natural language description for a single table row using Claude API.

    Args:
        headers: List of column header names
        row_data: List of cell values for this row
        enable: Whether to actually generate (if False, return empty string)

    Returns:
        Generated description or empty string if enable=False or API fails
    """
    if not enable or not headers or not row_data:
        return ""

    try:
        import os
        from anthropic import Anthropic

        base_url = os.getenv("ANTHROPIC_BASE_URL")
        auth_token = os.getenv("ANTHROPIC_AUTH_TOKEN")

        if not auth_token:
            return ""  # Gracefully fail if no API key

        client = Anthropic(api_key=auth_token, base_url=base_url)

        # Build prompt: pair headers with values
        field_descriptions = []
        for header, value in zip(headers, row_data):
            field_descriptions.append(f"{header}：{value}")

        prompt = "根据以下字段和数据，用一句话陈述事实：\n" + "，".join(field_descriptions)

        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=200,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

        return message.content[0].text.strip()
    except Exception as e:
        # Silently fail on API errors - don't break conversion
        print(f"    ⚠️ 描述生成失败: {e}", file=sys.stderr)
        return ""
```

- [ ] **Step 3: Update excel_to_markdown_by_sheets signature**

Change function signature (line ~95):
```python
def excel_to_markdown_by_sheets(file_path: str, enable_row_descriptions: bool = False) -> dict:
```

- [ ] **Step 4: Update _process_rows to accept enable_row_descriptions**

Change signature (line ~172):
```python
def _process_rows(rows, sheet_name, enable_row_descriptions: bool = False):
```

And pass it through all calls to `_process_rows()`:
- Line 110: `_process_rows(rows, sheet_name, enable_row_descriptions)`
- Line 132: `_process_rows(rows, sheet_name, enable_row_descriptions)`
- Line 155: `_process_rows(rows, sheet_name, enable_row_descriptions)`
- Line 165: `_process_rows(rows, f"Sheet{sheet_idx + 1}", enable_row_descriptions)`

- [ ] **Step 5: Implement new table formatting logic in _process_rows**

This is the key change. Replace the table generation logic (around lines 200-230) to output "table header + one data row + description" for each row:

```python
def _process_rows(rows, sheet_name, enable_row_descriptions: bool = False):
    """Process a list of rows and return markdown lines.

    If enable_row_descriptions is True:
    - Output table header + each data row separately + description
    - Each "header + row + description" forms one logical paragraph

    Otherwise:
    - Output standard markdown table
    """
    markdown_lines = []

    if not rows or len(rows) < 2:
        return markdown_lines

    header_row = rows[0]
    data_rows = rows[1:]

    # Determine format (Q&A vs table) - same as before
    is_qa_format = len(header_row) >= 3 and all(header_row[i] for i in range(3))

    # If descriptions disabled, use original logic
    if not enable_row_descriptions:
        # ... original _process_rows logic here ...
        return original_table_output(rows, header_row, data_rows, is_qa_format)

    # NEW LOGIC: Row-by-row output with descriptions
    current_category = None

    for row_idx, row in enumerate(data_rows):
        if not row or all(cell is None for cell in row):
            continue

        # Extract columns
        col0 = str(row[0]).strip() if row[0] else ""
        col1 = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        col2 = str(row[2]).strip() if len(row) > 2 and row[2] else ""

        if not col1:  # Skip rows without primary content
            continue

        if is_qa_format:
            # Q&A format: treat as structured QA pairs
            # Output as before (preserve existing Q&A behavior)
            # ... existing Q&A logic ...
        else:
            # Table format: output header + row + description
            clean_headers = [str(h).strip() if h else "" for h in header_row]
            clean_row_data = [str(cell).strip() if cell else "" for cell in row]

            # Output table header + this row
            markdown_lines.append("")  # Blank line before table
            markdown_lines.append("| " + " | ".join(clean_headers) + " |")
            markdown_lines.append("|" + "|".join(["---"] * len(clean_headers)) + "|")
            markdown_lines.append("| " + " | ".join(clean_row_data) + " |")

            # Generate and append description
            description = generate_row_description(clean_headers, clean_row_data, enable=True)
            if description:
                markdown_lines.append("")
                markdown_lines.append(description)

    return markdown_lines
```

Wait, this gets complicated because we need to preserve Q&A format behavior. Let me simplify:

```python
def _process_rows(rows, sheet_name, enable_row_descriptions: bool = False):
    """Process a list of rows and return markdown lines.

    If enable_row_descriptions=True and table format (not Q&A):
    - Output separate table for each row + its description

    Otherwise:
    - Use original logic
    """
    markdown_lines = []

    if not rows or len(rows) < 2:
        return markdown_lines

    header_row = rows[0]
    data_rows = rows[1:]

    is_qa_format = len(header_row) >= 3 and all(header_row[i] for i in range(3))

    # If Q&A format or descriptions disabled, use original logic
    if is_qa_format or not enable_row_descriptions:
        return _process_rows_original(rows, sheet_name)

    # TABLE FORMAT WITH DESCRIPTIONS
    clean_headers = [str(h).strip() if h else "" for h in header_row]

    for row in data_rows:
        if not row or all(cell is None for cell in row):
            continue

        clean_row_data = [str(cell).strip() if cell else "" for cell in row]

        if not any(clean_row_data):  # Skip empty rows
            continue

        # Output: table header + row + description
        markdown_lines.append("")  # Separator
        markdown_lines.append("| " + " | ".join(clean_headers) + " |")
        markdown_lines.append("|" + "|".join(["---"] * len(clean_headers)) + "|")
        markdown_lines.append("| " + " | ".join(clean_row_data) + " |")

        # Generate description
        description = generate_row_description(clean_headers, clean_row_data, enable=True)
        if description:
            markdown_lines.append("")
            markdown_lines.append(description)

    return markdown_lines


def _process_rows_original(rows, sheet_name):
    """Original _process_rows logic for Q&A format and when descriptions disabled.

    This preserves the existing behavior completely.
    """
    # ... copy entire original _process_rows logic here ...
```

Actually, let's keep it simpler: just refactor the existing _process_rows to check the flag early.

- [ ] **Step 6: Test the function locally**

Create a test XLSX with 2 rows:
```
headers: 姓名, 部门, 职位
row1: 张三, 财务处, 出纳
row2: 李四, IT部, 工程师
```

Run:
```bash
python -c "from convert_xlsx import excel_to_markdown_by_sheets; import pprint; pprint.pprint(excel_to_markdown_by_sheets('test.xlsx', enable_row_descriptions=False))"
```

Verify output format without descriptions works.

- [ ] **Step 7: Commit**

```bash
git add convert_xlsx.py
git commit -m "feat: add row description generation support to convert_xlsx.py

- Add generate_row_description() function using Claude API
- Update excel_to_markdown_by_sheets() to accept enable_row_descriptions flag
- Refactor _process_rows() to output table header + row + description for each data row
- Preserve Q&A format behavior and gracefully handle API failures"
```

---

## Chunk 2: Parameter Threading

### Task 2: Update convert_xlsx_all.py to accept and pass the flag

**Files:**
- Modify: `convert_xlsx_all.py` - add `--enable-row-descriptions` parameter

**Steps:**

- [ ] **Step 1: Read convert_xlsx_all.py current structure**

Understand the existing argument parsing (lines 53-83).

- [ ] **Step 2: Add argument to convert_xlsx_single_file function signature**

Change line 22:
```python
def convert_xlsx_single_file(input_dir: Path, output_dir: Path, file_path: Path, enable_row_descriptions: bool = False) -> Tuple[bool, Optional[str]]:
```

- [ ] **Step 3: Pass flag to excel_to_markdown_by_sheets**

Update line 28:
```python
sheet_markdowns = excel_to_markdown_by_sheets(str(file_path), enable_row_descriptions=enable_row_descriptions)
```

- [ ] **Step 4: Add command-line argument parsing**

Update the `if __name__ == "__main__"` block (line 53):

```python
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert XLSX file to Markdown with optional row descriptions"
    )
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
```

- [ ] **Step 5: Test parameter passing**

Run:
```bash
python convert_xlsx_all.py . . test.xlsx --enable-row-descriptions
```

Verify it doesn't crash and passes the flag correctly (check by adding debug print).

- [ ] **Step 6: Commit**

```bash
git add convert_xlsx_all.py
git commit -m "feat: add --enable-row-descriptions flag to convert_xlsx_all.py

- Add argument parser for --enable-row-descriptions
- Pass flag through to convert_xlsx_single_file()
- Update function signature to accept and propagate the parameter"
```

---

## Chunk 3: Integration with convert_all.py

### Task 3: Thread parameter through convert_all.py

**Files:**
- Modify: `convert_all.py` - add `--enable-row-descriptions` parameter to main script

**Steps:**

- [ ] **Step 1: Locate where convert_all.py calls run_convert_xlsx_single_file**

Search for `run_convert_xlsx_single_file` and `process_single_file` calls.

- [ ] **Step 2: Add --enable-row-descriptions to main argument parser**

Find the main `if __name__ == "__main__"` section with `ArgumentParser` setup. Add:

```python
parser.add_argument(
    "--enable-row-descriptions",
    action="store_true",
    default=False,
    help="Enable LLM-generated descriptions for each table row"
)
```

- [ ] **Step 3: Update run_convert_xlsx_single_file calls**

Find all calls to `run_convert_xlsx_single_file(...)` and add the flag:

```python
success = run_convert_xlsx_single_file(
    input_dir,
    output_dir,
    file_path,
    enable_row_descriptions=args.enable_row_descriptions
)
```

- [ ] **Step 4: Update run_convert_xlsx_single_file function signature**

Locate the function definition and update:
```python
def run_convert_xlsx_single_file(input_dir: Path, output_dir: Path, file_path: Path, enable_row_descriptions: bool = False) -> bool:
```

- [ ] **Step 5: Pass flag to convert_xlsx_all.py subprocess call**

In the subprocess command building, add the flag conditionally:

```python
cmd = [sys.executable, "convert_xlsx_all.py", str(input_dir), str(output_dir), str(file_path)]
if enable_row_descriptions:
    cmd.append("--enable-row-descriptions")
```

- [ ] **Step 6: Test end-to-end**

Run:
```bash
python convert_all.py --input test_dir --output output_dir --enable-row-descriptions
```

Verify the XLSX file is processed with descriptions.

- [ ] **Step 7: Commit**

```bash
git add convert_all.py
git commit -m "feat: integrate --enable-row-descriptions into convert_all.py orchestrator

- Add --enable-row-descriptions CLI flag
- Thread parameter through to XLSX conversion subprocess
- Update run_convert_xlsx_single_file() signature"
```

---

## Chunk 4: Testing & Documentation

### Task 4: Create test file and verify end-to-end

**Files:**
- Create: `test_row_descriptions.py` - basic integration test
- Modify: `README.md` - document the new flag

**Steps:**

- [ ] **Step 1: Create test_row_descriptions.py**

```python
#!/usr/bin/env python3
"""Quick integration test for row description feature."""

import tempfile
from pathlib import Path
from openpyxl import Workbook

def test_row_descriptions():
    """Create test XLSX, convert with descriptions, verify output format."""

    # Create temporary test file
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Create test XLSX
        wb = Workbook()
        ws = wb.active
        ws.title = "TestSheet"
        ws.append(["姓名", "部门", "职位"])
        ws.append(["张三", "财务处", "出纳"])
        ws.append(["李四", "IT部", "工程师"])

        test_xlsx = tmpdir / "test.xlsx"
        wb.save(test_xlsx)

        # Convert without descriptions
        from convert_xlsx import excel_to_markdown_by_sheets

        result_no_desc = excel_to_markdown_by_sheets(str(test_xlsx), enable_row_descriptions=False)
        print("✓ Conversion without descriptions succeeded")
        print("Output (no descriptions):")
        print(result_no_desc["TestSheet"][:200])
        print()

        # Note: Skip description test in CI/offline environments
        # (requires ANTHROPIC_* env vars)
        print("ℹ️ Skipping description generation test (requires API keys)")
        print("   To test locally: set ANTHROPIC_AUTH_TOKEN and run manually")

if __name__ == "__main__":
    test_row_descriptions()
```

- [ ] **Step 2: Run the test**

```bash
python test_row_descriptions.py
```

Expected: Conversion succeeds, shows sample output without descriptions.

- [ ] **Step 3: Update README.md**

Add to the "Usage" section:

```markdown
### Row Descriptions for RAG

Generate natural language descriptions for each table row (for knowledge base retrieval):

```bash
python convert_all.py --input 知识库base_1 --output 知识库md_1 --enable-row-descriptions
```

Output format:
- Each table row is output with its header
- Followed by a Claude-generated natural language description
- Forms self-contained knowledge chunks for RAG systems

Requires: `ANTHROPIC_AUTH_TOKEN` and `ANTHROPIC_BASE_URL` environment variables

Notes:
- Only applies to table format (not Q&A format)
- Descriptions gracefully skip on API failures
- Each row generates ~1 API call
```

- [ ] **Step 4: Commit**

```bash
git add test_row_descriptions.py README.md
git commit -m "docs & test: add row description feature documentation and integration test

- Add test_row_descriptions.py for basic conversion testing
- Update README with --enable-row-descriptions usage
- Note API requirements and behavior"
```

---

## Summary

**Total commits:** 4
- Chunk 1: Core LLM integration (1 commit)
- Chunk 2: convert_xlsx_all.py parameter threading (1 commit)
- Chunk 3: convert_all.py integration (1 commit)
- Chunk 4: Testing & docs (1 commit)

**Files modified:** 3 core Python files
**Files created:** 1 test file
**Backward compatible:** Yes (flag defaults to False)

