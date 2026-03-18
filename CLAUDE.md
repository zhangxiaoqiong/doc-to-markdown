# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

This is a **universal document-to-Markdown conversion tool** for 丰图科技 (Fengtu Technology). It converts company documents from multiple formats (DOCX, PDF, XLSX) into structured Markdown for use in knowledge base systems, Q&A platforms, and internal documentation management.

**Primary Use Cases:**
- 财务知识问答 (Financial Knowledge Q&A) - Financial policies and procedures
- 企业内部文档库 (Internal Documentation) - Department communications, guidelines, rules
- 杂项文档管理 (Miscellaneous Documentation) - Spreadsheets, forms, reference materials

## Repository Structure

```
.
├── 知识库base/              # Financial documents (PDFs, Word docs)
│   ├── Employee expense reimbursement policies
│   ├── Prepayment management rules
│   ├── Employee loan and advance management
│   ├── Customer rating management
│   └── Cost settlement standards
├── 知识库base_1/            # Internal corporate documents (XLSX, DOCX, PDF)
│   ├── Department communication guides
│   ├── Process guidelines
│   ├── Robot corpus
│   └── Various department materials
├── 知识库base_2/            # Additional documents (XLSX)
│   └── Supplementary materials
├── 知识库md/                # Converted financial documents (Markdown)
├── 知识库md_1/              # Converted internal documents (Markdown)
└── 知识库md_2/              # Converted additional documents (Markdown)
```

## Primary Tasks

1. **Document Conversion**: Convert documents from PDF/DOCX/XLSX to Markdown format
   - Supports batch processing with file-level serialization
   - Intelligent quality detection for PDFs (skips unnecessary Vision processing)
   - Automatic handling of different file types with appropriate tools

2. **Knowledge Base Enrichment**: Structure and organize documents for Q&A systems and knowledge bases

3. **Quality Assurance**:
   - Automatic detection of scanned/image-only PDFs
   - OCR error correction via Claude
   - Progress tracking via `inventory.xlsx` spreadsheet

## File Organization Guidelines

- Store markdown versions in appropriate output directories (`知识库md/`, `知识库md_1/`, `知识库md_2/`)
- Preserve document hierarchy and sections for easy navigation
- Include dates/versions in filenames when available
- Use consistent markdown formatting across all documents
- Financial documents: Cross-reference related policies (expense categories, approval limits)

## Supported File Formats

| Format | Tool | Use Case |
|--------|------|----------|
| DOCX | `convert_docs.py` | Structured documents, policies |
| PDF | `convert_docs.py` + `convert_split_pdf_v2.py` + Vision API | Mixed content (text + images) |
| XLSX | `convert_xlsx.py` | Spreadsheets, data tables, guidelines |

## Key Conversion Scripts

- **convert_docs.py** - DOCX and small PDF (<15MB) conversion
- **convert_split_pdf_v2.py** - Large PDF (>15MB) splitting and batch conversion
- **convert_xlsx.py** - XLSX parsing with XML fallback for corrupted files
- **convert_all.py** - Orchestrates all conversion steps with quality detection
- **fix_markdown_with_claude.py** - OCR error correction

## Notes for Future Work

- Documents contain policy-specific information that may require periodic updates
- Financial knowledge documents: maintain consistency in cross-references
- Consider automating periodic updates for policy documents
- Monitor `failed.log` for recurring failure patterns
- Review `inventory.xlsx` for quality metrics and anomalies

