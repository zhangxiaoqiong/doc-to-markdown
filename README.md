# 财务知识库文档转换工具

使用Claude AI将财务知识库文件（DOCX、PDF）转换为高质量Markdown格式。支持自动重试、失败日志和断点续传。

## 功能概览

| 脚本 | 用途 | 输入 | 输出 |
|------|------|------|------|
| `convert_docs.py` | 转换DOCX和小型PDF（<15MB）到Markdown | DOCX/PDF | Markdown |
| `convert_split_pdf_v2.py` | 分拆处理大型PDF（>15MB） | 大PDF | Markdown |
| `convert_all.py` | 一键完整转换（文件级串行处理，含4步流程和失败日志） | 目录 | Markdown + 失败日志 |
| `fix_markdown_with_claude.py` | 用Claude校对OCR错误 | Markdown目录 | 修复后的Markdown |

## 安装依赖

### 必需依赖

\`\`\`bash
pip install anthropic python-docx pdfplumber pypdf
\`\`\`

### 各库的用途

| 库 | 用途 | 使用脚本 |
|----|------|---------|
| \`anthropic\` | Claude API客户端 | 所有脚本 |
| \`python-docx\` | 提取DOCX内容 | convert_docs.py |
| \`pdfplumber\` | 提取PDF文本和表格 | convert_docs.py |
| \`pypdf\` | 分拆大型PDF | convert_split_pdf_v2.py |

**注意**：\`pdfplumber\` 用于纯文本提取（快速），\`pypdf\` 用于分割PDF（精确）。两者不可互换。

## 环境配置

### 1. 设置API认证

你需要设置两个环境变量来与Claude API通信：

#### Windows（命令行）

\`\`\`bash
set ANTHROPIC_BASE_URL=https://api.anthropic.com
set ANTHROPIC_AUTH_TOKEN=sk-ant-...
\`\`\`

#### Linux/Mac（Bash）

\`\`\`bash
export ANTHROPIC_BASE_URL=https://api.anthropic.com
export ANTHROPIC_AUTH_TOKEN=sk-ant-...
\`\`\`

#### PowerShell（Windows）

\`\`\`powershell
\$env:ANTHROPIC_BASE_URL="https://api.anthropic.com"
\$env:ANTHROPIC_AUTH_TOKEN="sk-ant-..."
\`\`\`

### 2. 验证环境变量

\`\`\`bash
# Windows
echo %ANTHROPIC_AUTH_TOKEN%

# Linux/Mac
echo $ANTHROPIC_AUTH_TOKEN
\`\`\`

## 使用指南

### 快速开始：一键转换所有文件

\`\`\`bash
python convert_all.py --input 知识库base --output 知识库md_v1.0
\`\`\`

这会对每个文件按以下步骤自动处理（**文件级串行**）：
1. **步骤1** - 用 `convert_docs.py` 提取内容并用Claude转换
2. **步骤2** - 检测转换质量，判断是否需要Vision重处理
3. **步骤3** - 对PDF用Vision API重新处理（保留图片和排版）；DOCX提示手工检查
4. **步骤4** - 用 `fix_markdown_with_claude.py` 校对OCR错误

**特点**：
- 每个文件完整走完4步后再处理下一个
- 失败文件记录到 `failed.log`，支持断点续传
- 已完成的文件自动跳过，增量处理

### 仅转换DOCX和小型PDF

\`\`\`bash
python convert_docs.py --input 知识库base --output 知识库md_v1.0
\`\`\`

**适用场景**：
- DOCX文件（自动提取结构）
- PDF < 15MB（直接转换）
- 快速处理，成本低

### 分拆处理大型PDF

\`\`\`bash
# 处理单个大PDF（>15MB）
python convert_split_pdf_v2.py 员工报销管理规定.pdf --input 知识库base --output 知识库md_v1.0

# 自定义分拆大小（每个块10页）
python convert_split_pdf_v2.py 大文件.pdf --input 知识库base --output 知识库md_v1.0 --pages-per-chunk 10
\`\`\`

**参数说明**：
- \`--pages-per-chunk\` (默认5)：每个块的页数。减少可以减少单个块的成本和失败率
- \`--min-size\` (默认15)：仅当文件大于此大小(MB)时才分拆

### 校对OCR错误

\`\`\`bash
python fix_markdown_with_claude.py --dir 知识库md_v1.0
\`\`\`

生成的Markdown会被检查并修复明显的OCR错误。

## 重试和失败处理

### 自动重试机制

所有脚本都内置了智能重试，处理API限流和临时错误：

- **重试条件**：429（限流）、500、503（服务不可用）错误
- **退避策略**：指数退避 (1s, 2s, 4s, ...)
- **最大重试**：3次

### 失败日志和断点续传

失败的文件会被记录到 \`failed.log\`，便于后续重跑：

\`\`\`
=== 2026-03-11 15:30:45 ===
文件名.pdf|429: Rate limit exceeded
另一个文件.docx|Connection timeout
\`\`\`

**断点续传**：删除 \`知识库md_v1.0/\` 中已成功转换的文件，然后重新运行脚本。脚本只会处理：
1. 源目录中有但输出目录中没有的文件
2. \`failed.log\` 中列出的失败文件

## 成本优化建议

| 方案 | 成本 | 质量 | 速度 |
|-----|------|------|------|
| 仅 \`convert_docs.py\` | 低 | 中 | 快 |
| \`convert_all.py\` | 高 | 高 | 中 |
| 手动分拆PDF（大块） | 中 | 中 | 快 |
| 手动分拆PDF（小块） | 高 | 高 | 慢 |

**建议工作流**：
1. 先用 \`convert_docs.py\` 快速转换，成本最低
2. 检查输出质量，如有问题再用 \`convert_all.py\` 重新处理
3. 对于关键PDF，增加 \`--pages-per-chunk\` 参数（更少分割 = 更低成本）

## 常见问题

### Q: 如何只重新转换失败的文件？

A: 脚本会自动跳过已存在的输出文件。要重新转换：
\`\`\`bash
# 删除想要重新转换的 .md 文件
rm 知识库md_v1.0/文件名.md
# 重新运行脚本
python convert_all.py
\`\`\`

### Q: 环境变量设置后仍然报错？

A: 验证设置：
\`\`\`bash
# 确认变量已设置
echo %ANTHROPIC_AUTH_TOKEN%  # Windows
echo $ANTHROPIC_AUTH_TOKEN   # Linux/Mac

# 如果为空，重新启动终端或使用管理员权限
\`\`\`

### Q: PDF转换后有OCR错误怎么办？

A: 运行校对脚本：
\`\`\`bash
python fix_markdown_with_claude.py --dir 知识库md_v1.0
\`\`\`

### Q: 如何降低成本？

A:
1. 使用 \`convert_docs.py\` 而不是Vision API（DOCX速度快，PDF只提取文本）
2. 对于大PDF，增加 \`--pages-per-chunk\`（更少API调用）
3. 批量转换后再做一次性校对（而不是逐个校对）

### Q: DOCX中的流程图为什么缺失？

A: \`python-docx\` 无法直接提取图片。对于包含流程图的DOCX：
1. 运行 \`convert_docs.py\` 获取文本内容
2. 手工打开Word文件，手动添加流程图的Markdown描述
3. 或将DOCX导出为PDF，再用Vision API处理

## 文件结构

\`\`\`
.
├── convert_docs.py                    # 主转换脚本（DOCX + 小PDF）
├── convert_split_pdf_v2.py            # 大PDF分拆转换
├── convert_all.py                     # 一键完整流程
├── fix_markdown_with_claude.py         # OCR错误校对
├── README.md                          # 本文件
├── 知识库base/                        # 源文件目录（DOCX/PDF）
└── 知识库md_v1.0/                     # 输出目录
    ├── file1.md                       # 转换后的Markdown
    ├── file2.md
    └── failed.log                     # 失败日志（自动生成）
\`\`\`

## 故障排查

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| "未设置环境变量" | API认证信息缺失 | 见[环境配置](#环境配置)章节 |
| "429 Too Many Requests" | API限流 | 脚本会自动重试，或减少 \`--pages-per-chunk\` |
| "文件格式不支持" | 仅支持DOCX和PDF | 将其他格式转换为PDF后重试 |
| OCR错误过多 | Vision API识别差 | 运行 \`fix_markdown_with_claude.py\` 校对 |

## 更新日志

### v1.1 (2026-03-11)
- ✅ 添加指数退避重试机制（处理429/500错误）
- ✅ 生成失败日志，支持断点续传
- ✅ 改进环境变量提示（Windows和Unix）
- ✅ 提高max_tokens从8192/12000到16000，减少截断
- ✅ 更新README文档

### v1.0 (2026-03-06)
- ✅ 初始版本，支持DOCX和PDF转换
