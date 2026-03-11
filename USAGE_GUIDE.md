# convert_all.py 使用指南

这是一份详细的使用指南，帮助你快速上手文档转换流程。`convert_all.py` 实现了**文件级串行处理**，支持断点续传，是处理大量文档的最佳方式。

## 快速开始

### 最简单的运行方式

```bash
python convert_all.py
```

这个命令会：
1. 读取 `知识库base/` 目录下的所有 PDF 和 DOCX 文件
2. 逐个处理每个文件（4步完整流程）
3. 输出转换后的 Markdown 到 `知识库md_v1.0/`
4. 自动记录失败文件到 `failed.log`

---

## 参数说明

### 可用参数

| 参数 | 简写 | 默认值 | 说明 |
|------|------|--------|------|
| `--input` | `-i` | `知识库base` | 输入目录，存放原始 PDF/DOCX 文件 |
| `--output` | `-o` | `知识库md_v1.0` | 输出目录，存放转换后的 Markdown 文件 |

### 参数示例

```bash
# 使用默认目录
python convert_all.py

# 指定自定义输入/输出目录
python convert_all.py --input 原始文件 --output markdown文件

# 只指定输出目录
python convert_all.py -o 知识库md_v2.0

# 只指定输入目录
python convert_all.py -i 新知识库
```

---

## 工作流程说明

### 4步转换流程（文件级串行）

对于每个文件，`convert_all.py` 执行以下步骤：

```
文件1 → [步骤1] → [步骤2] → [步骤3] → [步骤4] → 完成
文件2 → [步骤1] → [步骤2] → [步骤3] → [步骤4] → 完成
文件3 → ...
```

**步骤1：提取并初步转换**
- 使用 `convert_docs.py` 处理 DOCX 和纯文本 PDF
- 提取表格、列表、标题等结构化内容
- 输出：`output_dir/*.md`

**步骤2：智能质量检测** ⭐ **2026-03-11 优化版本**
- 检查输出质量，智能判断是否需要进一步处理
- 返回三种状态：
  - `'ok'` - 质量正常，无需额外处理
  - `'split'` - PDF >15MB，需要分割处理
  - `'vision'` - 输出质量异常，需要Vision处理
- 检查项（按优先级）：
  1. **大文件判断**：PDF > 15MB? → 标记为'split'
  2. **提取失败迹象**：输出太小(<源文件10%)? → 标记为'vision'
  3. **占位符检查**：是否有[此部分在第X页]? → 标记为'vision'
  4. **内容完整性**：输出太短(<500字符)? → 标记为'vision'
  5. **DOCX特殊处理**：DOCX有图片但无法解析? → 标记为'ok'（Vision不支持DOCX）

**步骤3：按需处理** ⭐ **2026-03-11 优化版本，三条处理路径**
- 根据步骤2的结果选择处理方式：

  **路径A (status='split')**: 分割处理
  - 用 `convert_split_pdf_v2.py` 处理 >15MB 的大PDF
  - 自动分割为5页/块，分别处理，然后合并

  **路径B (status='vision')**: Vision处理
  - 使用 `convert_pdf_vision.py` 利用 Claude Vision API
  - 重新识别内容，保留图片和排版
  - **新增**：自动清理Vision返回的```markdown污染标记

  **路径C (status='ok')**: 无需处理
  - 质量正常，直接跳过此步

**步骤4：最终校对**
- 使用 `fix_markdown_with_claude.py` 检查和修复 Markdown 格式
- 修复链接、代码块、表格等格式问题
- 最终输出：完成的 Markdown 文件

### 状态标记

运行时输出会显示每个文件的状态：

```
✅ [完成] file1.pdf - 4步全部处理成功
✅ [完成] file2.docx - 4步全部处理成功
⏭️  [跳过] file3.pdf - 已处理过，状态完整，跳过
❌ [失败] file4.pdf - 步骤3失败：超过限流

==================================================
统计：完成 2, 跳过 1, 失败 1
```

---

## 常见使用场景

### 场景1：首次转换所有文件（推荐）

**情况**：项目刚开始，需要处理所有源文件

```bash
python convert_all.py
```

**预期结果**：
- 所有文件都显示 `[完成]` 或 `[失败]`
- 失败文件记录在 `知识库md_v1.0/failed.log`
- 成功的 Markdown 文件存储在 `知识库md_v1.0/`

**耗时估计**：
- 小文件（< 5MB）：每个 1-2 分钟
- 大文件（> 15MB）：每个 5-10 分钟（自动分拆处理）

---

### 场景2：新增文件后继续处理

**情况**：已有文件处理完成，现在向 `知识库base/` 添加了新的文件

```bash
python convert_all.py
```

**行为**：
- 已完成的旧文件显示 `[跳过]`
- 新文件自动检测并处理，显示 `[完成]`
- 之前失败的文件会重新尝试

**示例输出**：
```
⏭️  [跳过] employee_policy.pdf - 已处理
⏭️  [跳过] prepayment_rules.docx - 已处理
✅ [完成] new_customer_rating.pdf - 新处理
⏭️  [跳过] settlement_rules.docx - 已处理
❌ [失败] technical_rules.pdf - 重试仍然失败
```

---

### 场景3：处理失败文件（等待后重试）

**情况**：某个文件处理失败（如 API 限流），想重新尝试

```bash
# 方案1：直接重新运行（推荐）
# 等待几分钟后运行，自动处理失败文件
python convert_all.py

# 方案2：查看失败原因（可选）
cat 知识库md_v1.0/failed.log
```

**failed.log 中的记录格式**：
```
[2026-03-11 15:30:45] file1.pdf | 步骤3失败: 429 Rate limit exceeded
[2026-03-11 15:31:20] file2.docx | 步骤1失败: Timeout after 3 retries
```

**如何恢复**：
1. 查看 `failed.log` 确认失败原因
2. 根据失败类型处理：
   - **429 (限流)**：等待 5-10 分钟后重新运行
   - **Timeout**：检查网络，可能需要增加重试等待时间
   - **Invalid format**：检查源文件是否损坏
3. 重新运行 `python convert_all.py`，自动处理失败的文件
4. 成功后自动从 `failed.log` 移除该文件记录

---

### 场景4：使用不同的输出目录（版本管理）

**情况**：想保留多个转换版本，比如 v1.0 和 v2.0

```bash
# 创建新版本
python convert_all.py --input 知识库base --output 知识库md_v2.0
```

**好处**：
- 保留历史版本，便于对比
- 不会覆盖已有的结果
- 便于测试新的转换策略

**查看版本**：
```bash
ls -la 知识库md_v1.0/    # v1.0 结果
ls -la 知识库md_v2.0/    # v2.0 结果
```

---

### 场景5：彻底清理并重新处理所有文件

**情况**：想放弃之前的处理结果，从零开始（如转换策略变更）

```bash
# Windows 用户
rmdir /s 知识库md_v1.0

# Linux/Mac 用户
rm -rf 知识库md_v1.0

# 重新转换
python convert_all.py
```

**谨慎**：这会删除所有已生成的 Markdown 文件和 failed.log

**更安全的做法**：
```bash
# 方案1：创建新版本而不是删除
python convert_all.py --output 知识库md_reprocess

# 方案2：只删除失败文件的记录，重新处理
# 手动编辑 知识库md_v1.0/failed.log，删除需要重新处理的文件行
# 然后运行：
python convert_all.py
```

---

## 环境要求检查

运行 `convert_all.py` 前，需要检查以下环境：

### 检查清单

```bash
# 1. Python 版本（需要 3.7+）
python --version

# 2. 必需的包
pip list | grep -E "anthropic|python-docx|pdf2image|Pillow"

# 3. API Key 配置
echo $ANTHROPIC_API_KEY              # Linux/Mac
echo %ANTHROPIC_API_KEY%             # Windows cmd
```

### 环境变量设置

如果提示缺少 API Key，需要设置环境变量：

**Windows（PowerShell）**：
```powershell
$env:ANTHROPIC_API_KEY = "your-api-key-here"
```

**Windows（命令提示符）**：
```cmd
set ANTHROPIC_API_KEY=your-api-key-here
```

**Linux/Mac**：
```bash
export ANTHROPIC_API_KEY=your-api-key-here
```

---

## 预期输出示例

### 完整运行示例

```
$ python convert_all.py
正在检查环境...
✓ Python 版本: 3.9.0
✓ 依赖包检查完成
✓ API Key 已配置

正在处理 7 个文件...

==================================================
✅ [完成] employee_expense_policy.pdf
   步骤1: ✓ 提取内容成功
   步骤2: ✓ 质量检测：无需额外处理（status='ok'）
   步骤3: ⏭️  跳过（质量正常）
   步骤4: ✓ 校对完成
   输出: 知识库md_v1.0/employee_expense_policy.md (45KB)

✅ [完成] prepayment_rules.docx
   步骤1: ✓ DOCX解析成功
   步骤2: ⏭️  质量检测：DOCX文件无需Vision
   步骤3: ⏭️  跳过
   步骤4: ✓ 校对完成
   输出: 知识库md_v1.0/prepayment_rules.md (38KB)

⏭️  [跳过] loan_management.docx - 已处理

✅ [完成] customer_rating_rules.pdf
   步骤1: ✓ PDF提取成功
   步骤2: ✓ 质量检测：无需额外处理（status='ok'）
   步骤3: ⏭️  跳过（质量正常）
   步骤4: ✓ 校对完成
   输出: 知识库md_v1.0/customer_rating_rules.md (52KB)

✅ [完成] large_document.pdf
   步骤1: ✓ PDF提取成功 (20MB)
   步骤2: ✓ 质量检测：需要分割处理（status='split'）
   步骤3: ✓ 分割处理成功（分为4块，依次处理并合并）
   步骤4: ✓ 校对完成
   输出: 知识库md_v1.0/large_document.md (95KB)

❌ [失败] settlement_exceptions.pdf
   步骤1: ✓ PDF提取成功
   步骤2: ✓ 质量检测：检测到异常（status='vision'）
   步骤3: ✗ Vision处理失败 (429 Rate limit exceeded)
   操作: 已记录到 failed.log，将在下次运行时重试

⏭️  [跳过] technical_guide.pdf - 已处理
⏭️  [跳过] cost_settlement.docx - 已处理

==================================================
总结：
  ✅ 完成: 4 个文件
  ⏭️  跳过: 3 个文件
  ❌ 失败: 1 个文件

下一步建议：
  1. 5分钟后重新运行处理失败的文件
  2. 检查 知识库md_v1.0/failed.log 查看失败详情
```

---

## 最佳实践

### 1. 定期检查 failed.log

```bash
# 查看失败日志
cat 知识库md_v1.0/failed.log

# 如果想清空日志（表示已全部处理）
rm 知识库md_v1.0/failed.log
```

### 2. 处理 API 限流

如果看到 `429 Rate limit exceeded` 错误：

```bash
# 等待至少 5-10 分钟
# 然后重新运行
python convert_all.py

# 或者使用自定义输出目录避免跳过
python convert_all.py --output 知识库md_v2.0
```

### 3. 验证转换质量

```bash
# 查看输出文件列表
ls -lh 知识库md_v1.0/*.md

# 查看 DOCX 文件中是否有未提取的图片/流程图
# （需要手工检查或用Word打开源文件确认）

# 随机抽查一个 Markdown 文件
cat 知识库md_v1.0/employee_expense_policy.md | head -50
```

### 4. 版本管理建议

```bash
# 保留多个版本便于对比
python convert_all.py --output 知识库md_v1.0  # 首次
python convert_all.py --output 知识库md_v2.0  # 优化后重新处理

# 比较两个版本的文件大小
ls -lh 知识库md_v1.0/*.md | awk '{print $5, $9}'
ls -lh 知识库md_v2.0/*.md | awk '{print $5, $9}'
```

### 5. 批量处理多个目录

如果有多个知识库需要处理：

```bash
# 处理目录1
python convert_all.py --input 知识库base --output 知识库md_v1.0

# 处理目录2
python convert_all.py --input 新知识库base --output 新知识库md
```

---

## 故障排查

### 问题1：提示"输入目录不存在"

```
错误：输入目录不存在 知识库base
```

**检查**：
```bash
# 确认目录是否存在
ls 知识库base/

# 如果目录不存在，指定正确的路径
python convert_all.py --input /path/to/your/documents
```

### 问题2：缺少 API Key

```
错误：缺少 ANTHROPIC_API_KEY 环境变量
请设置：export ANTHROPIC_API_KEY=<your-key>
```

**解决**：
```bash
# 设置 API Key（见前面的"环境变量设置"章节）
export ANTHROPIC_API_KEY=sk-ant-xxxx...

# 验证设置
echo $ANTHROPIC_API_KEY

# 重新运行
python convert_all.py
```

### 问题3：某个文件持续失败

```
❌ [失败] large_document.pdf
   步骤3: Vision处理失败 (429 Rate limit exceeded)
```

**解决方案**（按优先级）：

1. **如果是限流（429）**：
   - 等待 10-15 分钟
   - 重新运行 `python convert_all.py`

2. **如果是超时或网络错误**：
   - 检查网络连接
   - 尝试使用代理（如需要）
   - 重新运行

3. **如果是格式错误**：
   - 检查源文件是否损坏
   - 尝试用 Word/PDF 阅读器打开确认
   - 可考虑手动重新保存文件

4. **如果多次重试仍失败**：
   ```bash
   # 手动指定该文件输出路径重新处理
   python convert_docs.py large_document.pdf --output 知识库md_v1.0/large_document.md
   ```

### 问题4：输出文件不完整或格式错误

```bash
# 检查文件大小是否合理
ls -lh 知识库md_v1.0/*.md

# 查看文件开头和结尾
head -20 知识库md_v1.0/document.md
tail -20 知识库md_v1.0/document.md
```

**可能原因和解决**：

| 现象 | 原因 | 解决方案 |
|------|------|---------|
| 文件很小（<5KB） | 可能提取失败或文件为空 | 检查源文件，看是否真的是文档 |
| 出现占位符如"[此部分内容在第X页]" | PDF分拆后的残留 | 这通常已被清理，如仍存在请报告 |
| 表格/列表格式混乱 | Markdown格式问题 | 手动检查并修复，或用 fix_markdown_with_claude.py 单独处理 |
| 中文显示为乱码 | 编码问题（较少见） | 尝试 `iconv -f GBK -t UTF-8 file.md -o file_utf8.md` |

### 问题5：硬盘空间不足

```
错误：磁盘空间不足
```

**检查和清理**：
```bash
# 查看硬盘使用情况
df -h

# 查看输出目录大小
du -sh 知识库md_v1.0/

# 删除旧版本释放空间（确保已备份）
rm -rf 知识库md_v0.9/
```

**预计空间需求**：
- 输入文件总大小 × 0.5-1.5 倍（Markdown 通常更小）
- 缓存文件（临时）：取决于文件大小
- 推荐保留至少 1GB 自由空间

---

## 常见问题 (FAQ)

**Q：为什么有些文件标记 [跳过]？**
A：文件已经完整处理过（输出文件存在且不在 failed.log 中），没有必要重复处理。如要强制重新处理，删除对应的 .md 文件即可。

**Q：failed.log 该什么时候删除？**
A：不需要手动删除。当失败的文件成功处理后，会自动从 failed.log 中移除。如果 failed.log 中全是已处理的文件，可以安心删除。

**Q：能否并行处理多个文件加快速度？**
A：当前实现是串行处理（一个文件完整走完4步，再处理下一个）。这是有意设计，便于故障隔离和日志追踪。如需更快处理，可考虑提交多个 convert_all.py 进程。

**Q：处理 DOCX 文件中的图片会丢失吗？**
A：是的，convert_docs.py 无法自动提取 DOCX 中的图片。如果文档有重要图片/流程图，建议：
1. 用 Vision API 处理（需手动运行 convert_pdf_vision.py）
2. 或手动将图片转换为 PDF 后用 Vision 处理
3. 或手工补充到 Markdown 中

**Q：如何处理特别大的 PDF 文件？**
A：convert_all.py 自动检测文件大小：
- < 15MB：直接处理
- ≥ 15MB：自动分拆为 5 页一块，并行处理后合并

无需手动干预。

**Q：输出的 Markdown 有目录吗？**
A：有。大多数文档的 Markdown 会保留原有的标题层级，便于生成目录。如需生成目录 TOC，可用工具如 `markdown-toc` 或在 Markdown 编辑器中自动生成。

---

## 下一步

1. **首次运行**：按照"场景1"执行 `python convert_all.py`
2. **检查结果**：查看 `知识库md_v1.0/` 中的 Markdown 文件
3. **处理失败**：根据 failed.log 内容，按"故障排查"中的方案处理
4. **集成应用**：将生成的 Markdown 文件集成到 Q&A 系统或其他应用中

---

## 相关文档

- [README.md](README.md) - 项目概览和安装指南
- [IMPLEMENTATION_NOTES.md](IMPLEMENTATION_NOTES.md) - 文件级串行处理的技术细节
- convert_docs.py - DOCX 和纯文本 PDF 处理
- convert_pdf_vision.py - Vision API 处理
- fix_markdown_with_claude.py - Markdown 校对

---

## 📋 更新日志

### v2.1（2026-03-11 新增PDF处理逻辑优化）✨

**新特性**：
- ⭐ 智能三态检测：从单纯的True/False改为'ok'|'split'|'vision'三态返回
- ⭐ 自动污染清理：Vision处理后自动清理```markdown包装标记
- ⭐ 大文件智能分割：>15MB的PDF自动标记为'split'而不是无条件Vision
- ⭐ 质量异常检测：输出太小/含占位符/内容太短时自动触发Vision

**性能改进**：
- Vision调用减少75%（从4个文件→1个文件）
- 处理时间显著减少（Vision处理较慢）
- API配额节省（减少不必要Vision调用）

**质量改进**：
- 避免Vision重新理解导致表格列数减少
- 输出Markdown干净无污染
- 保留convert_docs的原始表格结构完整性

**详细说明**：
- 查看 [PDF_OPTIMIZATION_SUMMARY.md](PDF_OPTIMIZATION_SUMMARY.md) - 完整的优化分析
- 查看 [CODE_CHANGES_DETAIL.md](CODE_CHANGES_DETAIL.md) - 详细的代码改动对比

### v2.0（2026-03-11）
- 文件级串行处理（每个文件完整走完4步，再处理下一个）
- 断点续传支持
- Excel处理记录生成

### v1.0
- 初版实现

---

**最后更新**: 2026-03-11
**适用版本**: convert_all.py v2.1+（PDF处理智能优化版本）
