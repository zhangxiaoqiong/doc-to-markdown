# 文件级串行处理改造 - 实现说明

**实现日期**: 2026-03-11
**改造范围**: `convert_all.py` 文件
**兼容性**: 完全向后兼容

## 改造内容速览

### 从分阶段批量 → 文件级串行

**原先流程**（批量阶段式）:
```
阶段1: 处理所有文件 (convert_docs.py)
阶段2: 检测所有文件 (质量检测)
阶段3: Vision重处理所有PDF
阶段4: 校对所有文件 (fix_markdown)
```

**现在流程**（文件级串行）:
```
FOR EACH 源文件:
  IF 已完成 → 跳过
  ELSE:
    步骤1: 转换
    步骤2: 检测
    步骤3: Vision重处理（如需）
    步骤4: 校对
    记录结果
```

## 核心新增函数

### 1. 状态判断
```python
is_file_complete(output_dir, source_file_path)
  返回: True if (输出.md存在 && 不在failed.log中)
```

### 2. 失败记录
```python
log_failed(output_dir, file_name, error_msg)
  向failed.log追加: [时间] 文件名 | 错误原因

remove_from_failed_log(output_dir, file_name)
  处理成功后，从failed.log中移除该文件记录
```

### 3. 单文件处理
```python
process_single_file(input_dir, output_dir, file_path)
  串联4个步骤处理单个文件
  返回: (success: bool, error_msg: str)
```

## 使用示例

### 基础用法（无改变）
```bash
python convert_all.py --input 知识库base --output 知识库md_v1.0
```

### 输出示例
```
============================================================
开始处理 7 个文件（文件级串行）
============================================================

[处理] file1.pdf
  [步骤1] 转换: file1.pdf
  [步骤2] 检测: file1.pdf
  [步骤3] Vision重处理: PDF文件需要Vision保留图片和排版
  [步骤4] 校对OCR: file1.pdf
[完成] file1.pdf

[跳过] file2.docx 已处理

[处理] file3.pdf
  [步骤1] 转换: file3.pdf
  [步骤2] 检测: file3.pdf
  [失败] file3.pdf: 步骤1转换失败: 429 Rate limit exceeded

============================================================
✓ 转换完成！
  完成: 5
  跳过: 1
  失败: 1

✗ 1 个文件处理失败，详见 知识库md_v1.0/failed.log
============================================================
```

## 断点续传工作流

### 场景1: 新增文件
```bash
# 第1次：处理 file1, file2
python convert_all.py
# 输出: [完成] file1, [完成] file2

# 第2次：新增 file3
python convert_all.py
# 输出: [跳过] file1, [跳过] file2, [完成] file3
```

### 场景2: 失败后重试
```bash
# 第1次：file2转换失败
python convert_all.py
# 输出: [完成] file1, [失败] file2, [完成] file3
# failed.log: [时间] file2.pdf | 429 Rate limit exceeded

# 第2次：重新运行（file2会从failed.log中自动重新处理）
python convert_all.py
# 输出: [跳过] file1, [完成] file2, [跳过] file3
```

### 场景3: 手动重新处理
```bash
# 删除想要重新处理的.md文件
rm 知识库md_v1.0/problematic_file.md

# 重新运行脚本
python convert_all.py
# problematic_file会被重新处理
```

## Failed.log 格式

```
[2026-03-11 15:30:45] file1.pdf | 429: Rate limit exceeded
[2026-03-11 15:31:20] file2.docx | 步骤1转换失败: Connection timeout
[2026-03-11 15:32:10] file3.pdf | 步骤3Vision处理失败: Invalid PDF format
```

## 关键改进

| 方面 | 原先 | 现在 |
|------|------|------|
| 处理粒度 | 批量阶段式 | 文件级串行 |
| 失败定位 | 不清楚卡在哪一步 | 明确显示卡在哪一步 |
| 中途失败 | 需要手工跳过已成功的文件 | 自动跳过，继续处理新文件 |
| 增量处理 | 需要手工删除.md | 自动识别已完成的文件 |
| 新增文件 | 从头开始处理所有文件 | 只处理新增 + 失败文件 |

## 兼容性声明

✅ 完全向后兼容
- 保持 `--input` 和 `--output` 参数
- 保持输出目录结构不变
- 保持环境变量要求不变
- 其他脚本 (convert_docs.py, convert_split_pdf_v2.py, fix_markdown_with_claude.py) 无改动

## 技术实现细节

### 文件完成状态判断
```python
is_complete = (
    输出.md文件存在 AND
    文件名不在failed.log中
)
```

### 错误处理
- 每步都有 try-catch
- 失败时记录具体错误信息
- 提供有意义的错误提示给用户

### 日志管理
- failed.log 自动创建和追加
- 处理成功后自动清理记录
- 时间戳便于追踪

## 常见问题

### Q: 如何知道文件处理到了哪一步？
A: 输出会显示 `[步骤1/2/3/4]` 标记，失败时会明确说明卡在哪一步。

### Q: 失败的文件会自动重新处理吗？
A: 会的。下次运行脚本时，failed.log中的文件会被自动重新处理。

### Q: 如何强制重新处理某个文件？
A: 删除对应的 .md 文件再运行脚本即可。

### Q: failed.log 什么时候清除？
A: 当文件被成功处理后，会自动从 failed.log 中移除对应记录。

### Q: 可以中途停止吗？
A: 可以。下次运行时会从未处理的文件开始继续。

## 后续维护

定期运行来保持知识库更新：
```bash
python convert_all.py
```

脚本会自动：
1. 跳过已完成的文件
2. 处理新增的文件
3. 重新尝试之前失败的文件
