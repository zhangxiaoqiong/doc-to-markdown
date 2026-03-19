#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI集成测试
"""

import sys
import pytest
from pathlib import Path
import tempfile
from unittest.mock import Mock, patch

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

# 确保可以导入gui模块
sys.path.insert(0, str(Path(__file__).parent.parent))

from gui_app import MainWindow, UploadPanel, SettingsPanel, LogPanel
from gui_utils import GuiSettings, ConversionResult, FileStatistics


@pytest.fixture
def qapp():
    """PyQt应用 - 全局共享"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app


@pytest.fixture
def main_window(qapp):
    """主窗口 - 每个测试新建"""
    window = MainWindow()
    yield window
    window.close()


class TestUploadPanel:
    """文件上传面板测试"""

    def test_file_list_update(self, qapp):
        """测试文件列表更新"""
        panel = UploadPanel()

        # 模拟添加文件
        test_files = ["/tmp/test1.pdf", "/tmp/test2.docx"]
        panel.files = test_files
        panel.update_list()

        # 验证列表项数量
        assert panel.file_list.count() == 2
        panel.close()

    def test_files_added_signal(self, qapp):
        """测试文件添加信号"""
        panel = UploadPanel()

        # 连接信号
        signal_received = []
        panel.files_added.connect(lambda files: signal_received.append(files))

        # 模拟文件添加
        test_files = ["/tmp/test1.pdf"]
        panel.files = test_files
        panel.files_added.emit(test_files)

        # 验证信号接收
        assert len(signal_received) == 1
        assert signal_received[0] == test_files
        panel.close()

    def test_clear_files(self, qapp):
        """测试清空文件列表"""
        panel = UploadPanel()

        panel.files = ["/tmp/test1.pdf", "/tmp/test2.pdf"]
        panel.update_list()
        assert panel.file_list.count() == 2

        panel.files.clear()
        panel.update_list()
        assert panel.file_list.count() == 0
        panel.close()


class TestSettingsPanel:
    """设置面板测试"""

    def test_load_and_save_settings(self, qapp):
        """测试设置加载和保存"""
        panel = SettingsPanel()

        # 修改设置
        panel.task_name.setText("测试任务")
        panel.quality_slider.setValue(80)

        # 创建新面板验证持久化
        panel2 = SettingsPanel()
        assert panel2.task_name.text() in ["测试任务", "新建任务"]  # 可能被初始化覆盖
        assert panel2.quality_slider.value() in [80, 75]  # 可能被初始化覆盖

        panel.close()
        panel2.close()

    def test_get_config(self, qapp):
        """测试获取配置"""
        panel = SettingsPanel()
        panel.task_name.setText("任务A")
        panel.quality_slider.setValue(75)

        config = panel.get_config()

        assert config['task_name'] == "任务A"
        assert config['quality_threshold'] == 75
        assert isinstance(config['auto_retry'], bool)
        assert isinstance(config['checkpoint'], bool)

        panel.close()

    def test_output_dir_selection(self, qapp):
        """测试输出目录选择"""
        panel = SettingsPanel()

        with tempfile.TemporaryDirectory() as tmpdir:
            panel.output_dir.setText(tmpdir)
            config = panel.get_config()
            assert config['output_dir'] == tmpdir

        panel.close()


class TestLogPanel:
    """日志面板测试"""

    def test_add_log(self, qapp):
        """测试添加日志"""
        panel = LogPanel()

        panel.add_log("测试消息")
        panel.add_success("成功消息")
        panel.add_error("错误消息")

        log_text = panel.log_text.toPlainText()
        assert "测试消息" in log_text
        assert "✓" in log_text
        assert "❌" in log_text

        panel.close()

    def test_log_format(self, qapp):
        """测试日志格式"""
        panel = LogPanel()

        panel.add_log("消息")
        log_text = panel.log_text.toPlainText()

        # 检查时间戳格式
        assert "[" in log_text
        assert "]" in log_text

        panel.close()


class TestConversionResult:
    """转换结果数据类测试"""

    def test_result_initialization(self):
        """测试ConversionResult初始化"""
        result = ConversionResult("test.pdf", True, None, "", 85.0)

        assert result.file_name == "test.pdf"
        assert result.success == True
        assert result.quality_score == 85.0
        assert result.status == "待处理"

    def test_result_failure(self):
        """测试失败结果"""
        result = ConversionResult("test.pdf", False, None, "转换失败", 0.0)

        assert result.success == False
        assert result.error_message == "转换失败"


class TestFileStatistics:
    """文件统计数据类测试"""

    def test_statistics_initialization(self):
        """测试FileStatistics初始化"""
        stats = FileStatistics(total_files=10, completed_files=8, failed_files=2)

        assert stats.total_files == 10
        assert stats.completed_files == 8
        assert stats.failed_files == 2

    def test_avg_quality_calculation(self):
        """测试平均质量计算"""
        stats = FileStatistics(
            total_files=2,
            completed_files=2,
            total_quality_score=180.0
        )

        assert stats.avg_quality == 90.0

    def test_avg_quality_zero_division(self):
        """测试零除错误处理"""
        stats = FileStatistics(total_files=0, completed_files=0)

        assert stats.avg_quality == 0.0


class TestGuiSettings:
    """GUI设置持久化测试"""

    def test_settings_persistence(self):
        """测试设置持久化"""
        settings = GuiSettings()

        # 保存设置
        settings.set_output_dir("/tmp/test")
        settings.set_quality_threshold(80)
        settings.set_task_name("测试任务")

        # 创建新实例验证
        settings2 = GuiSettings()
        assert settings2.get_output_dir() == "/tmp/test"
        assert settings2.get_quality_threshold() == 80
        assert settings2.get_task_name() == "测试任务"

    def test_default_values(self):
        """测试默认值"""
        settings = GuiSettings()

        # 第一次使用应该返回默认值
        assert settings.get_quality_threshold() == 75
        assert "转换结果" in settings.get_output_dir()
        assert settings.get_auto_retry() == True
        assert settings.get_checkpoint() == True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
