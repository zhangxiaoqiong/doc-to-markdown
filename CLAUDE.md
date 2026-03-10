# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

This is a **financial knowledge base repository** for 丰图科技 (Fengtu Technology). It serves as a centralized source for company financial policies and management procedures. The repository supports a "财务知识问答" (Financial Knowledge Q&A) system.

## Repository Structure

```
.
├── 知识库base/          # Original source documents (PDFs, Word docs)
│   ├── Employee expense reimbursement policies
│   ├── Prepayment management rules
│   ├── Employee loan and advance management
│   ├── Customer rating management
│   └── Cost settlement standards
└── 知识库md/            # Markdown-formatted versions of documents (to be populated)
```

## Primary Tasks

1. **Document Conversion**: Convert financial policy documents from original formats (PDF, DOCX) to markdown files in `知识库md/`
2. **Knowledge Base Enrichment**: Organize and structure documents for use in Q&A systems
3. **Policy Documentation**: Maintain accurate, searchable versions of company financial policies

## File Organization Guidelines

- Store markdown versions in `知识库md/` with clear, descriptive filenames matching source documents
- Preserve document hierarchy and sections for easy navigation
- Include dates/versions in filenames when available (e.g., `员工费用报销规定_v3.0.md`)
- Use consistent markdown formatting across all documents

## Key Documents

- **员工费用报销规定** (Employee Expense Reimbursement Policy)
- **丰图科技预付管理规定** (Prepayment Management Rules)
- **备用金及个人借款管理规范** (Advance and Personal Loan Management Standards)
- **应付结算例外事项管理办法** (Payable Settlement Exception Management)
- **丰图科技客户评级管理规则** (Customer Rating Management Rules)
- **项目投入及费用结算管理** (Project Investment and Cost Settlement Management)

## Notes for Future Work

- Documents contain policy-specific information (dates, rates, approval thresholds) that may require periodic updates
- Cross-references between documents exist (e.g., expense categories, approval limits) - maintain consistency
- Consider creating an index document linking all knowledge base items
