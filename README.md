# 企业文档转换工具（Document-to-Markdown Converter）

使用Claude AI将企业文档（DOCX、PDF、XLSX）批量转换为高质量Markdown格式。支持智能质量检测、自动重试、失败日志、断点续传和进度追踪。

## 功能概览

| 脚本 | 用途 | 输入 | 输出 |
|------|------|------|------|
| `convert_docs.py` | 转换DOCX和小型PDF（<15MB）到Markdown | DOCX/PDF | Markdown |
| `convert_split_pdf_v2.py` | 分拆处理大型PDF（>15MB） | 大PDF | Markdown |
| `convert_xlsx.py` | XLSX核心转换逻辑（含XML备用方案） | XLSX | Markdown |
| `convert_xlsx_all.py` | 一键处理单个XLSX文件或目录 | XLSX | Markdown |
| `convert_all.py` | 一键完整转换（文件级串行处理，含4步流程、失败日志、智能质量检测、清单追踪） | 目录 | Markdown + 失败日志 + inventory.xlsx |
| `fix_markdown_with_claude.py` | 用Claude校对OCR错误 | Markdown目录 | 修复后的Markdown |

## 安装依赖

### 必需依赖

\`\`\`bash
pip install anthropic python-docx pdfplumber pypdf openpyxl openpyxl.drawing.image openpyxl-image-loader
\`\`\`

### 各库的用途

| 库 | 用途 | 使用脚本 |
|----|------|---------|
| \`anthropic\` | Claude API客户端 | 所有脚本 |
| \`python-docx\` | 提取DOCX内容 | convert_docs.py |
| \`pdfplumber\` | 提取PDF文本和表格、检测扫描件 | convert_docs.py、convert_all.py |
| \`pypdf\` | 分拆大型PDF | convert_split_pdf_v2.py |
| \`openpyxl\` | 解析XLSX文件 | convert_xlsx.py、convert_xlsx_all.py |
| \`openpyxl.drawing.image\` | XLSX中的图片处理 | convert_xlsx.py、convert_xlsx_all.py |

**注意**：
- \`pdfplumber\` 用于纯文本提取（快速）和扫描件检测，\`pypdf\` 用于分割PDF（精确）
- \`openpyxl\` 用于XLSX解析，含XML备用方案处理损坏的样式

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

这会对每个文件按以下步骤自动处理（**文件级串行**，每个文件完整走完4步再处理下一个）：

#### PDF文件处理流程
1. **步骤1** - 用 \`convert_docs.py\` 提取内容并用Claude转换
2. **步骤2** - 智能质量检测，判断处理策略：
   - PDF >15MB → 分拆处理（\`convert_split_pdf_v2.py\`）
   - PDF质量异常 → Vision API重处理
   - PDF质量良好 → 直接完成（跳过不必要的Vision处理）
   - 纯图片/扫描件 → 标记为需要人工检查
3. **步骤3** - 对符合条件的PDF用Vision API重新处理（保留图片和排版）
4. **步骤4** - 用 \`fix_markdown_with_claude.py\` 校对OCR错误

#### DOCX文件处理流程
1. **步骤1** - 用 \`convert_docs.py\` 提取内容并用Claude转换
2-4. **步骤2-4** - 跳过（DOCX无需Vision处理）

#### XLSX文件处理流程
1. **步骤1** - 用 \`convert_xlsx_all.py\` 转换
2-4. **步骤2-4** - 跳过（XLSX无需多步处理）

**特点**：
- ✅ 智能质量检测，避免画蛇添足（好的PDF直接跳过Vision）
- ✅ 扫描件前置检测，及时预警需要人工检查的文件
- ✅ 失败文件记录到 \`failed.log\`，支持断点续传
- ✅ 生成 \`inventory.xlsx\` 清单，实时追踪转换进度和质量
- ✅ 已完成的文件自动跳过，增量处理

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

### 处理XLSX文件

\`\`\`bash
# 处理单个XLSX文件
python convert_xlsx_all.py --input 知识库base_1 --output 知识库md_v1.0 --file 语料库.xlsx

# 处理目录中的所有XLSX文件
python convert_xlsx_all.py --input 知识库base_1 --output 知识库md_v1.0
\`\`\`

**特点**：
- 直接从Excel提取表格数据、图片和元数据
- 自动处理损坏的样式（XML备用方案）
- 保留表格、合并单元格和图片
- 跳过冗余的Vision处理，快速高效

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

脚本会自动处理：
1. 源目录中有但输出目录中没有的文件（新增文件）
2. \`failed.log\` 中列出的失败文件（重试）
3. 自动跳过已完成的文件（增量处理）

**清单追踪**：每次运行时生成或更新 \`inventory.xlsx\`，记录：
- 文件名、源文件大小、输出文件大小
- 使用的处理方法、是否检测到图片、异常标记
- 处理状态（完成/失败/警告）、备注

## 成本优化建议

| 方案 | 成本 | 质量 | 速度 |
|-----|------|------|------|
| 仅 \`convert_docs.py\` | 低 | 中 | 快 |
| \`convert_all.py\`（启用智能检测） | 中 | 高 | 中 |
| \`convert_all.py\`（无智能检测） | 高 | 高 | 中 |
| XLSX处理（\`convert_xlsx_all.py\`） | 最低 | 中 | 快 |
| 手动分拆PDF（大块） | 中 | 中 | 快 |
| 手动分拆PDF（小块） | 高 | 高 | 慢 |

**建议工作流**：
1. 先用 \`convert_docs.py\` 快速转换，成本最低
2. 检查输出质量，如有问题再用 \`convert_all.py\` 重新处理（智能检测会跳过质量好的PDF）
3. XLSX文件直接用 \`convert_xlsx_all.py\` 处理，无需多步
4. 对于关键PDF，增加 \`--pages-per-chunk\` 参数（更少分割 = 更低成本）

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
2. 用 \`convert_all.py\` 的智能质量检测（好的PDF会自动跳过Vision，不必要的重处理会被跳过）
3. XLSX文件直接用 \`convert_xlsx_all.py\` 处理，无需Vision
4. 对于大PDF，增加 \`--pages-per-chunk\`（更少API调用）
5. 批量转换后再做一次性校对（而不是逐个校对）

### Q: 如何处理扫描件或纯图片PDF？

A: \`convert_all.py\` 会自动检测扫描件并标记为"⚠️异常"，提醒需要人工处理：
- 检查 \`inventory.xlsx\` 的"异常标记"列
- 对于纯图片PDF，需要用其他OCR工具或Vision API手动处理
- 转换后检查Markdown内容确认质量

### Q: XLSX转换出错？

A: \`convert_xlsx.py\` 包含XML备用方案处理损坏的样式：
1. 首先尝试用openpyxl直接解析
2. 如果失败，自动尝试XML + SharedStrings方案
3. 如果仍然失败，记入failed.log并返回错误

对于特殊格式的XLSX，检查 \`inventory.xlsx\` 的"备注"列获取详细错误信息。

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
├── convert_xlsx.py                    # XLSX核心逻辑（含XML备用方案）
├── convert_xlsx_all.py                # XLSX可执行脚本
├── convert_all.py                     # 一键完整流程（文件级串行 + 智能质量检测 + 清单追踪）
├── fix_markdown_with_claude.py         # OCR错误校对
├── README.md                          # 本文件
├── USAGE_GUIDE.md                     # 详细使用指南
├── 知识库base/                        # 源文件目录（DOCX/PDF）- 财务知识库
├── 知识库base_1/                      # 源文件目录（XLSX/DOCX/PDF）- 企业内部文档库
├── 知识库base_2/                      # 源文件目录（XLSX）- 企业内部新增文档
├── 知识库md/                          # 输出目录（财务知识库）
├── 知识库md_1/                        # 输出目录（企业内部文档库）
├── 知识库md_2/                        # 输出目录（企业内部新增文档）
│   ├── file1.md                       # 转换后的Markdown
│   ├── file2.md
│   ├── failed.log                     # 失败日志（自动生成）
│   └── inventory.xlsx                 # 清单追踪表（自动生成）
└── docs/                              # 技术文档目录
    ├── 技术文档_文档转换方案总结.md
    └── 失败经验总结_4个废弃脚本.md
\`\`\`

## 故障排查

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| "未设置环境变量" | API认证信息缺失 | 见[环境配置](#环境配置)章节 |
| "429 Too Many Requests" | API限流 | 脚本会自动重试，或减少 \`--pages-per-chunk\` |
| "文件格式不支持" | 仅支持DOCX、PDF和XLSX | 将其他格式转换为相应格式后重试 |
| "⚠️异常"标记 | 文件质量异常（扫描件、纯图片） | 检查 \`inventory.xlsx\` 的"备注"列，可能需要人工处理 |
| OCR错误过多 | Vision API识别差 | 运行 \`fix_markdown_with_claude.py\` 校对 |
| XLSX解析失败 | 样式损坏或特殊格式 | 检查 \`failed.log\`，脚本会自动尝试XML备用方案 |

## 更新日志

### v1.3 (2026-03-18)
- ✅ 添加XLSX支持（openpyxl + XML备用方案）
- ✅ 智能质量检测，避免无必要的Vision处理
- ✅ 扫描件前置检测，及时预警
- ✅ 生成 \`inventory.xlsx\` 清单追踪
- ✅ 更新README文档（完整功能说明）

### v1.2 (2026-03-12)
- ✅ 第五阶段：XLSX核心转换逻辑和可执行脚本
- ✅ 集成XLSX到convert_all.py（自动识别文件类型）

### v1.1 (2026-03-11)
- ✅ 添加指数退避重试机制（处理429/500错误）
- ✅ 生成失败日志，支持断点续传
- ✅ 改进环境变量提示（Windows和Unix）
- ✅ 提高max_tokens从8192/12000到16000，减少截断
- ✅ 更新README文档

### v1.0 (2026-03-06)
- ✅ 初始版本，支持DOCX和PDF转换
