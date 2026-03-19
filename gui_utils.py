#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI辅助模块：配置管理、数据模型、转换调用接口
"""

import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

# Windows编码修复
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


# ============================================================================
# 数据模型
# ============================================================================

@dataclass
class ConversionResult:
    """单个文件转换结果"""
    file_name: str
    success: bool = False
    output_file: Optional[Path] = None
    error_message: str = ""
    quality_score: float = 0.0
    status: str = "待处理"  # 完成/失败/警告


@dataclass
class FileStatistics:
    """批处理统计"""
    total_files: int = 0
    completed_files: int = 0
    failed_files: int = 0
    review_files: int = 0  # 待审核（quality < threshold）
    total_quality_score: float = 0.0

    @property
    def avg_quality(self) -> float:
        if self.completed_files > 0:
            return self.total_quality_score / self.completed_files
        return 0.0


# ============================================================================
# 配置管理
# ============================================================================

class GuiSettings:
    """GUI配置持久化（QSettings包装）"""

    def __init__(self):
        try:
            from PyQt6.QtCore import QSettings
            self.settings = QSettings("Fengtu", "DocumentConverter")
            self._use_qsettings = True
        except ImportError:
            # 如果PyQt6不可用，使用文件系统备选方案
            self._use_qsettings = False
            self._settings_file = Path.home() / ".fengtu_converter.ini"
            self._settings_cache = {}
            self._load_cache()

    def _load_cache(self):
        """从文件加载缓存（备选方案）"""
        if not self._use_qsettings and self._settings_file.exists():
            try:
                with open(self._settings_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        if '=' in line:
                            key, value = line.strip().split('=', 1)
                            self._settings_cache[key] = value
            except Exception:
                pass

    def _save_cache(self):
        """保存缓存到文件（备选方案）"""
        if not self._use_qsettings:
            try:
                with open(self._settings_file, 'w', encoding='utf-8') as f:
                    for key, value in self._settings_cache.items():
                        f.write(f"{key}={value}\n")
            except Exception:
                pass

    def get_output_dir(self) -> str:
        default = str(Path.home() / "转换结果")
        if self._use_qsettings:
            return self.settings.value("output_dir", default)
        else:
            return self._settings_cache.get("output_dir", default)

    def set_output_dir(self, path: str):
        if self._use_qsettings:
            self.settings.setValue("output_dir", path)
        else:
            self._settings_cache["output_dir"] = path
            self._save_cache()

    def get_quality_threshold(self) -> int:
        if self._use_qsettings:
            return int(self.settings.value("quality_threshold", 75))
        else:
            return int(self._settings_cache.get("quality_threshold", "75"))

    def set_quality_threshold(self, value: int):
        if self._use_qsettings:
            self.settings.setValue("quality_threshold", value)
        else:
            self._settings_cache["quality_threshold"] = str(value)
            self._save_cache()

    def get_task_name(self) -> str:
        default = "新建任务"
        if self._use_qsettings:
            return self.settings.value("task_name", default)
        else:
            return self._settings_cache.get("task_name", default)

    def set_task_name(self, name: str):
        if self._use_qsettings:
            self.settings.setValue("task_name", name)
        else:
            self._settings_cache["task_name"] = name
            self._save_cache()

    def get_auto_retry(self) -> bool:
        if self._use_qsettings:
            return self.settings.value("auto_retry", True, type=bool)
        else:
            value = self._settings_cache.get("auto_retry", "True")
            return value.lower() == "true"

    def set_auto_retry(self, value: bool):
        if self._use_qsettings:
            self.settings.setValue("auto_retry", value)
        else:
            self._settings_cache["auto_retry"] = str(value)
            self._save_cache()

    def get_checkpoint(self) -> bool:
        if self._use_qsettings:
            return self.settings.value("checkpoint", True, type=bool)
        else:
            value = self._settings_cache.get("checkpoint", "True")
            return value.lower() == "true"

    def set_checkpoint(self, value: bool):
        if self._use_qsettings:
            self.settings.setValue("checkpoint", value)
        else:
            self._settings_cache["checkpoint"] = str(value)
            self._save_cache()


# ============================================================================
# 转换调用接口
# ============================================================================

def convert_single_file(
    input_file: Path,
    output_dir: Path,
    quality_threshold: int = 75
) -> ConversionResult:
    """
    调用convert_all.py处理单个文件所在目录

    Args:
        input_file: 输入文件路径
        output_dir: 输出目录
        quality_threshold: 质量阈值（暂未使用，预留接口）

    Returns:
        ConversionResult: 转换结果
    """
    result = ConversionResult(file_name=input_file.name)

    try:
        # 验证输入文件存在
        if not input_file.exists():
            result.success = False
            result.error_message = f"输入文件不存在: {input_file}"
            result.status = "失败"
            return result

        # 创建输出目录
        output_dir.mkdir(parents=True, exist_ok=True)

        # 构建命令
        convert_script = Path(__file__).parent / "convert_all.py"
        if not convert_script.exists():
            result.success = False
            result.error_message = f"转换脚本不存在: {convert_script}"
            result.status = "失败"
            return result

        # convert_all.py 处理整个目录，所以传入文件所在目录
        cmd = [
            sys.executable,
            str(convert_script),
            "--input",
            str(input_file.parent),
            "--output",
            str(output_dir)
        ]

        # 添加模拟处理时间（增加2秒延迟让用户能看到进度），然后运行转换（超时30分钟）
        import time
        time.sleep(2)

        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=1800,
            cwd=Path(__file__).parent
        )

        if proc.returncode == 0:
            # 检查输出文件是否存在
            output_file = output_dir / (input_file.stem + ".md")
            if output_file.exists():
                result.success = True
                result.output_file = output_file
                result.status = "完成"
                result.quality_score = 85.0  # 默认评分
            else:
                result.success = False
                result.error_message = "转换完成但输出文件不存在"
                result.status = "失败"
        else:
            result.success = False
            # 优先使用stderr，如果为空使用stdout
            error_output = proc.stderr.strip() if proc.stderr else proc.stdout.strip()
            result.error_message = error_output or "未知错误"
            result.status = "失败"

    except subprocess.TimeoutExpired:
        result.success = False
        result.error_message = "转换超时（超过30分钟）"
        result.status = "失败"
    except Exception as e:
        result.success = False
        result.error_message = str(e)
        result.status = "失败"

    return result
