#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Document-to-Markdown Converter - PyQt6 简化原型
极简设计，专注核心功能
"""

import sys
import json
from pathlib import Path
from datetime import datetime
from typing import List

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QSpinBox, QCheckBox,
    QFileDialog, QTableWidget, QTableWidgetItem, QProgressBar,
    QTabWidget, QSplitter, QListWidget, QListWidgetItem, QMessageBox,
    QDialog, QSlider, QComboBox, QTextEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QFont, QColor

from gui_utils import convert_single_file, FileStatistics, ConversionResult


# ============================================================================
# 核心业务逻辑（连接到现有Python代码）
# ============================================================================

class ConversionWorker(QThread):
    """后台转换线程 - 集成真实转换逻辑"""
    progress_updated = pyqtSignal(int, int, str)  # 当前、总数、状态
    finished = pyqtSignal(bool, str, dict)  # 成功、信息、统计字典
    file_result = pyqtSignal(str, bool, str)  # 文件名、成功/失败、错误信息

    def __init__(self, files: List[str], output_dir: str, quality_threshold: int):
        super().__init__()
        self.files = [Path(f) for f in files]
        self.output_dir = Path(output_dir)
        self.quality_threshold = quality_threshold
        self.is_running = True
        self.statistics = FileStatistics(total_files=len(self.files))

    def run(self):
        """运行真实转换"""
        try:
            # 确保输出目录存在
            self.output_dir.mkdir(parents=True, exist_ok=True)

            for i, file_path in enumerate(self.files):
                if not self.is_running:
                    break

                # 更新进度
                status_text = f"处理中: {file_path.name}"
                self.progress_updated.emit(i, len(self.files), status_text)

                # 调用真实转换
                try:
                    result = convert_single_file(
                        file_path,
                        self.output_dir,
                        self.quality_threshold
                    )

                    # 更新统计
                    if result.success:
                        self.statistics.completed_files += 1
                        if result.quality_score < self.quality_threshold:
                            self.statistics.review_files += 1
                        self.statistics.total_quality_score += result.quality_score
                    else:
                        self.statistics.failed_files += 1

                    # 发出文件结果信号
                    self.file_result.emit(
                        result.file_name,
                        result.success,
                        result.error_message
                    )

                except Exception as e:
                    self.statistics.failed_files += 1
                    self.file_result.emit(
                        file_path.name,
                        False,
                        f"转换异常: {str(e)}"
                    )

            # 生成最终消息
            message = (
                f"转换完成！"
                f"成功: {self.statistics.completed_files}, "
                f"失败: {self.statistics.failed_files}, "
                f"待审核: {self.statistics.review_files}"
            )

            # 发出完成信号，包含统计数据
            stats_dict = {
                'completed': self.statistics.completed_files,
                'failed': self.statistics.failed_files,
                'review': self.statistics.review_files,
                'avg_quality': self.statistics.avg_quality
            }
            self.finished.emit(True, message, stats_dict)

        except Exception as e:
            self.finished.emit(False, f"错误: {str(e)}", {})

    def stop(self):
        """停止转换"""
        self.is_running = False


# ============================================================================
# UI 组件
# ============================================================================

class UploadPanel(QWidget):
    """文件上传面板"""
    files_added = pyqtSignal(list)
    
    def __init__(self):
        super().__init__()
        self.files = []
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 上传区域
        self.upload_btn = QPushButton("📁 选择文件")
        self.upload_btn.setMinimumHeight(60)
        self.upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 2px dashed #1a73e8;
                border-radius: 8px;
                font-size: 14px;
                color: #202124;
            }
            QPushButton:hover {
                background-color: #e8f0fe;
            }
        """)
        self.upload_btn.clicked.connect(self.select_files)
        
        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(150)
        
        # 清空按钮
        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(lambda: (self.file_list.clear(), self.files.clear()))
        
        layout.addWidget(QLabel("📄 文件选择"))
        layout.addWidget(self.upload_btn)
        layout.addWidget(QLabel("已添加文件:"))
        layout.addWidget(self.file_list)
        layout.addWidget(clear_btn)
        
        self.setLayout(layout)
    
    def select_files(self):
        """选择文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "选择文档",
            "",
            "Document Files (*.docx *.pdf *.xlsx);;All Files (*)"
        )
        if files:
            self.files.extend(files)
            self.update_list()
            self.files_added.emit(self.files)
    
    def update_list(self):
        """更新文件列表显示"""
        self.file_list.clear()
        for file in self.files:
            size = Path(file).stat().st_size / (1024 * 1024)  # MB
            item_text = f"{Path(file).name} ({size:.1f}MB)"
            self.file_list.addItem(item_text)


class SettingsPanel(QWidget):
    """设置面板"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 任务名称
        layout.addWidget(QLabel("⚙️ 转换设置"))
        
        layout.addWidget(QLabel("任务名称:"))
        self.task_name = QLineEdit("新建任务")
        layout.addWidget(self.task_name)
        
        # 输出目录
        layout.addWidget(QLabel("输出目录:"))
        output_layout = QHBoxLayout()
        self.output_dir = QLineEdit(str(Path.home() / "转换结果"))
        browse_btn = QPushButton("浏览...")
        browse_btn.setMaximumWidth(80)
        browse_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(self.output_dir)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)
        
        # 质量阈值
        layout.addWidget(QLabel("质量阈值:"))
        threshold_layout = QHBoxLayout()
        self.quality_slider = QSlider(Qt.Orientation.Horizontal)
        self.quality_slider.setRange(0, 100)
        self.quality_slider.setValue(75)
        self.quality_label = QLabel("75 分")
        self.quality_label.setMaximumWidth(50)
        self.quality_slider.valueChanged.connect(
            lambda v: self.quality_label.setText(f"{v} 分")
        )
        threshold_layout.addWidget(self.quality_slider)
        threshold_layout.addWidget(self.quality_label)
        layout.addLayout(threshold_layout)
        
        # 高级选项
        layout.addWidget(QLabel("选项:"))
        self.retry_check = QCheckBox("自动重试（失败文件）")
        self.retry_check.setChecked(True)
        self.checkpoint_check = QCheckBox("断点续传（中断后恢复）")
        self.checkpoint_check.setChecked(True)
        layout.addWidget(self.retry_check)
        layout.addWidget(self.checkpoint_check)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir.setText(dir_path)
    
    def get_config(self):
        """获取配置"""
        return {
            'task_name': self.task_name.text(),
            'output_dir': self.output_dir.text(),
            'quality_threshold': self.quality_slider.value(),
            'auto_retry': self.retry_check.isChecked(),
            'checkpoint': self.checkpoint_check.isChecked(),
        }


class ProgressPanel(QWidget):
    """进度追踪面板"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 统计信息
        stats_layout = QHBoxLayout()
        
        self.stat_completed = QLabel("✓ 已完成: 0")
        self.stat_processing = QLabel("⏳ 处理中: 0")
        self.stat_pending = QLabel("⊙ 待处理: 0")
        self.stat_failed = QLabel("✗ 失败: 0")
        
        stats_layout.addWidget(self.stat_completed)
        stats_layout.addWidget(self.stat_processing)
        stats_layout.addWidget(self.stat_pending)
        stats_layout.addWidget(self.stat_failed)
        
        layout.addLayout(stats_layout)
        
        # 进度条
        layout.addWidget(QLabel("总体进度:"))
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        layout.addWidget(self.progress_bar)
        
        # 当前状态
        layout.addWidget(QLabel("状态:"))
        self.status_label = QLabel("就绪")
        layout.addWidget(self.status_label)
        
        # 时间信息
        time_layout = QHBoxLayout()
        self.time_elapsed = QLabel("耗时: 00:00")
        self.time_remaining = QLabel("剩余: --:--")
        self.speed_info = QLabel("速度: -- 文件/分钟")
        time_layout.addWidget(self.time_elapsed)
        time_layout.addWidget(self.time_remaining)
        time_layout.addWidget(self.speed_info)
        layout.addLayout(time_layout)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def update_progress(self, current: int, total: int, status: str):
        """更新进度"""
        if total > 0:
            percent = int(current * 100 / total)
            self.progress_bar.setValue(percent)
            self.status_label.setText(status)
            
            # 更新统计
            self.stat_completed.setText(f"✓ 已完成: {current}")
            self.stat_pending.setText(f"⊙ 待处理: {total - current}")


class ResultsPanel(QWidget):
    """结果展示面板"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 统计
        layout.addWidget(QLabel("📊 转换结果"))
        
        stats_layout = QHBoxLayout()
        self.result_success = QLabel("✓ 成功: 0")
        self.result_review = QLabel("⚠️ 待审核: 0")
        self.result_failed = QLabel("✗ 失败: 0")
        self.result_quality = QLabel("质量评分: --")
        stats_layout.addWidget(self.result_success)
        stats_layout.addWidget(self.result_review)
        stats_layout.addWidget(self.result_failed)
        stats_layout.addWidget(self.result_quality)
        layout.addLayout(stats_layout)
        
        # 结果表格
        layout.addWidget(QLabel("文件详情:"))
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["文件名", "状态", "评分", "操作"])
        self.result_table.setMaximumHeight(200)
        layout.addWidget(self.result_table)
        
        # 导出按钮
        export_layout = QHBoxLayout()
        self.export_btn = QPushButton("📥 导出结果")
        self.export_btn.setEnabled(False)
        self.open_folder_btn = QPushButton("📁 打开文件夹")
        self.open_folder_btn.setEnabled(False)
        export_layout.addWidget(self.export_btn)
        export_layout.addWidget(self.open_folder_btn)
        layout.addLayout(export_layout)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def show_results(self, success: int, review: int, failed: int, quality: float):
        """显示结果"""
        self.result_success.setText(f"✓ 成功: {success}")
        self.result_review.setText(f"⚠️ 待审核: {review}")
        self.result_failed.setText(f"✗ 失败: {failed}")
        self.result_quality.setText(f"质量评分: {quality:.1f}")
        self.export_btn.setEnabled(True)
        self.open_folder_btn.setEnabled(True)


class LogPanel(QWidget):
    """日志面板"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("📋 日志"))
        
        # 日志显示
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        layout.addWidget(self.log_text)
        
        # 按钮
        btn_layout = QHBoxLayout()
        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        export_btn = QPushButton("导出日志")
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(export_btn)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def add_log(self, message: str):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def add_error(self, message: str):
        """添加错误日志"""
        self.add_log(f"❌ {message}")
    
    def add_success(self, message: str):
        """添加成功日志"""
        self.add_log(f"✓ {message}")


# ============================================================================
# 主窗口
# ============================================================================

class MainWindow(QMainWindow):
    """主应用窗口"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文档转换工具 - Document-to-Markdown Converter")
        self.setGeometry(100, 100, 1000, 700)
        
        self.conversion_worker = None
        self.init_ui()
    
    def init_ui(self):
        """初始化UI"""
        # 中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout()
        
        # ========== 左侧面板（上传和设置）==========
        left_layout = QVBoxLayout()
        
        # 选项卡切换
        self.left_tabs = QTabWidget()
        
        self.upload_panel = UploadPanel()
        self.settings_panel = SettingsPanel()
        
        self.left_tabs.addTab(self.upload_panel, "📄 文件")
        self.left_tabs.addTab(self.settings_panel, "⚙️ 设置")
        
        left_layout.addWidget(self.left_tabs)
        
        # 转换按钮
        self.convert_btn = QPushButton("▶ 开始转换")
        self.convert_btn.setMinimumHeight(40)
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #1a73e8;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #1557b0;
            }
            QPushButton:pressed {
                background-color: #0d47a1;
            }
        """)
        self.convert_btn.clicked.connect(self.start_conversion)
        left_layout.addWidget(self.convert_btn)
        
        # 暂停按钮（转换中显示）
        self.pause_btn = QPushButton("⏸ 暂停")
        self.pause_btn.setVisible(False)
        self.pause_btn.clicked.connect(self.pause_conversion)
        left_layout.addWidget(self.pause_btn)
        
        left_widget = QWidget()
        left_widget.setLayout(left_layout)
        left_widget.setMaximumWidth(350)
        
        # ========== 右侧面板（进度和结果）==========
        right_layout = QVBoxLayout()
        
        self.right_tabs = QTabWidget()
        
        self.progress_panel = ProgressPanel()
        self.results_panel = ResultsPanel()
        self.log_panel = LogPanel()
        
        self.right_tabs.addTab(self.progress_panel, "⏳ 进度")
        self.right_tabs.addTab(self.results_panel, "✓ 结果")
        self.right_tabs.addTab(self.log_panel, "📋 日志")
        
        right_layout.addWidget(self.right_tabs)
        
        right_widget = QWidget()
        right_widget.setLayout(right_layout)
        
        # ========== 组合布局 ==========
        main_layout.addWidget(left_widget, 1)
        main_layout.addWidget(right_widget, 1)
        
        central_widget.setLayout(main_layout)
        
        # 连接信号
        self.upload_panel.files_added.connect(self.on_files_added)
        
        # 初始化状态
        self.update_status("就绪")
    
    def on_files_added(self, files: List[str]):
        """文件添加时更新UI"""
        self.log_panel.add_success(f"已添加 {len(files)} 个文件")
    
    def start_conversion(self):
        """开始转换"""
        if not self.upload_panel.files:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return
        
        config = self.settings_panel.get_config()
        output_dir = Path(config['output_dir'])
        
        # 创建输出目录
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 启动转换线程
        self.conversion_worker = ConversionWorker(
            self.upload_panel.files,
            str(output_dir),
            config['quality_threshold']
        )
        
        self.conversion_worker.progress_updated.connect(self.on_progress_updated)
        self.conversion_worker.finished.connect(self.on_conversion_finished)
        
        self.conversion_worker.start()
        
        # 更新UI状态
        self.convert_btn.setVisible(False)
        self.pause_btn.setVisible(True)
        self.log_panel.add_success(f"开始转换 {len(self.upload_panel.files)} 个文件")
        self.right_tabs.setCurrentIndex(0)  # 显示进度面板
    
    def on_progress_updated(self, current: int, total: int, status: str):
        """进度更新回调"""
        self.progress_panel.update_progress(current, total, status)
        self.log_panel.add_log(status)
    
    def on_conversion_finished(self, success: bool, message: str, stats: dict):
        """转换完成回调"""
        self.convert_btn.setVisible(True)
        self.pause_btn.setVisible(False)

        if success:
            self.log_panel.add_success(message)
            self.results_panel.show_results(
                success=stats.get('completed', 0),
                review=stats.get('review', 0),
                failed=stats.get('failed', 0),
                quality=stats.get('avg_quality', 0.0)
            )
            self.right_tabs.setCurrentIndex(1)  # 显示结果面板
            QMessageBox.information(self, "完成", message)
        else:
            self.log_panel.add_error(message)
            QMessageBox.critical(self, "错误", message)

        self.update_status("就绪")
    
    def pause_conversion(self):
        """暂停转换"""
        if self.conversion_worker:
            self.conversion_worker.stop()
            self.log_panel.add_log("转换已暂停")
    
    def update_status(self, status: str):
        """更新状态栏"""
        self.statusBar().showMessage(f"状态: {status}")


# ============================================================================
# 应用入口
# ============================================================================

def main():
    app = QApplication(sys.argv)
    
    # 设置应用风格
    app.setStyle('Fusion')
    
    # 创建主窗口
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()