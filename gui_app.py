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
            import time

            # 确保输出目录存在
            self.output_dir.mkdir(parents=True, exist_ok=True)

            for i, file_path in enumerate(self.files):
                if not self.is_running:
                    break

                # 更新进度（i+1表示已处理的文件数）
                status_text = f"处理中: {file_path.name}"
                self.progress_updated.emit(i + 1, len(self.files), status_text)

                # 短暂延迟让UI能更新进度条显示（50ms）
                time.sleep(0.05)

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
        self.current_folder = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 按钮行
        btn_layout = QHBoxLayout()

        # 选择文件按钮
        self.upload_btn = QPushButton("📄 选择文件")
        self.upload_btn.setMinimumHeight(60)
        self.upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 2px dashed #1a73e8;
                border-radius: 8px;
                font-size: 13px;
                color: #202124;
            }
            QPushButton:hover {
                background-color: #e8f0fe;
            }
        """)
        self.upload_btn.clicked.connect(self.select_files)

        # 选择文件夹按钮
        self.folder_btn = QPushButton("📁 选择文件夹")
        self.folder_btn.setMinimumHeight(60)
        self.folder_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 2px dashed #1a73e8;
                border-radius: 8px;
                font-size: 13px;
                color: #202124;
            }
            QPushButton:hover {
                background-color: #e8f0fe;
            }
        """)
        self.folder_btn.clicked.connect(self.select_folder)

        btn_layout.addWidget(self.upload_btn)
        btn_layout.addWidget(self.folder_btn)

        # 当前文件夹显示
        self.folder_label = QLabel("📁 文件夹: 未选择")
        self.folder_label.setStyleSheet("color: #999; font-size: 11px; padding: 5px;")

        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(150)

        # 清空按钮
        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(lambda: (self.file_list.clear(), self.files.clear(), self.reset_folder_label()))

        layout.addWidget(QLabel("📄 文件选择"))
        layout.addLayout(btn_layout)
        layout.addWidget(self.folder_label)
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

    def select_folder(self):
        """选择文件夹并自动扫描支持的文件"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "选择文件夹"
        )
        if folder_path:
            self.current_folder = folder_path
            # 更新文件夹显示标签
            folder_name = Path(folder_path).name
            self.folder_label.setText(f"📁 文件夹: {folder_name}")

            # 扫描文件夹中的所有支持的文件
            folder = Path(folder_path)
            supported_extensions = {'.docx', '.pdf', '.xlsx'}
            found_files = []

            for ext in supported_extensions:
                found_files.extend([str(f) for f in folder.glob(f'*{ext}')])

            if found_files:
                # 只添加未在列表中的文件
                new_files = [f for f in found_files if f not in self.files]
                self.files.extend(new_files)
                self.update_list()
                self.files_added.emit(self.files)
            else:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.information(self, "提示", f"文件夹中没有找到DOCX、PDF或XLSX文件")

    def reset_folder_label(self):
        """重置文件夹显示"""
        self.current_folder = None
        self.folder_label.setText("📁 文件夹: 未选择")

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
        self.output_dir = None
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

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.open_folder_btn = QPushButton("📁 打开文件夹")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        btn_layout.addWidget(self.open_folder_btn)
        layout.addLayout(btn_layout)

        layout.addStretch()
        self.setLayout(layout)

    def show_results(self, success: int, review: int, failed: int, quality: float):
        """显示结果"""
        self.result_success.setText(f"✓ 成功: {success}")
        self.result_review.setText(f"⚠️ 待审核: {review}")
        self.result_failed.setText(f"✗ 失败: {failed}")
        self.result_quality.setText(f"质量评分: {quality:.1f}")
        self.open_folder_btn.setEnabled(True)

    def set_output_dir(self, path: str):
        """设置输出目录"""
        self.output_dir = Path(path)

    def open_output_folder(self):
        """打开输出文件夹"""
        if self.output_dir and self.output_dir.exists():
            import subprocess
            if sys.platform == "win32":
                subprocess.Popen(f'explorer /select,"{self.output_dir}"')
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(self.output_dir)])
            else:
                subprocess.Popen(["xdg-open", str(self.output_dir)])
        else:
            QMessageBox.warning(self, "提示", "输出目录不存在")


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
    """主应用窗口 - 向导式工作流设计"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("文档转换工具 - Document-to-Markdown Converter")
        self.setGeometry(100, 100, 1100, 800)

        self.conversion_worker = None
        self.log_panel = None
        self.current_workflow = None

        self.init_ui()

    def init_ui(self):
        """初始化UI - 使用QSplitter分割上下，日志始终可见"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)

        # 创建分割器（上：工作流容器，下：日志）
        splitter = QSplitter(Qt.Orientation.Vertical)

        # ========== 上半部分：工作流容器 ==========
        self.workflow_container = QWidget()
        self.workflow_layout = QVBoxLayout()
        self.workflow_layout.setContentsMargins(0, 0, 0, 0)
        self.workflow_container.setLayout(self.workflow_layout)

        splitter.addWidget(self.workflow_container)

        # ========== 下半部分：日志（始终可见）==========
        self.log_panel = LogPanel()
        splitter.addWidget(self.log_panel)

        # 设置分割器大小（上：600px，下：200px）
        splitter.setSizes([600, 200])
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)

        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)

        # 初始化状态
        self.update_status("就绪")
        self.show_menu()

    def show_menu(self):
        """显示主菜单"""
        # 清空容器
        while self.workflow_layout.count():
            self.workflow_layout.takeAt(0).widget().deleteLater()

        # 创建菜单
        menu_frame = MainMenuFrame(self)
        self.workflow_layout.addWidget(menu_frame)
        self.current_workflow = menu_frame

    def show_workflow(self, workflow_type: str):
        """显示指定工作流"""
        # 清空容器
        while self.workflow_layout.count():
            self.workflow_layout.takeAt(0).widget().deleteLater()

        # 创建工作流
        if workflow_type == "quick":
            workflow = QuickConversionFrame(self)
        elif workflow_type == "custom":
            workflow = CustomConversionFrame(self)
        elif workflow_type == "batch":
            workflow = BatchConversionFrame(self)
        else:
            return

        self.workflow_layout.addWidget(workflow)
        self.current_workflow = workflow

    def start_conversion(self, files: List[str], output_dir: str, quality_threshold: int):
        """开始转换（由工作流调用）"""
        if not files:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return

        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # 启动转换线程
        self.conversion_worker = ConversionWorker(
            files,
            str(output_path),
            quality_threshold
        )

        # 连接信号
        self.conversion_worker.file_result.connect(self.on_file_converted)
        self.conversion_worker.progress_updated.connect(self.on_progress_updated)
        self.conversion_worker.finished.connect(self.on_conversion_finished)

        self.conversion_worker.start()

        # 更新日志
        self.log_panel.add_success(f"开始转换 {len(files)} 个文件")
        self.update_status("转换中...")

    def on_progress_updated(self, current: int, total: int, status: str):
        """进度更新回调"""
        self.log_panel.add_log(status)
        # 通知工作流更新进度
        if self.current_workflow and hasattr(self.current_workflow, 'update_progress'):
            self.current_workflow.update_progress(current, total, status)

    def on_conversion_finished(self, success: bool, message: str, stats: dict):
        """转换完成回调"""
        if success:
            self.log_panel.add_success(message)
            # 通知工作流转换完成
            if self.current_workflow and hasattr(self.current_workflow, 'on_conversion_finished'):
                self.current_workflow.on_conversion_finished(success, message, stats)
        else:
            self.log_panel.add_error(message)
            QMessageBox.critical(self, "错误", message)

        self.update_status("就绪")

    def on_file_converted(self, file_name: str, success: bool, error_msg: str):
        """单个文件转换完成回调"""
        if success:
            self.log_panel.add_success(f"✓ {file_name}")
        else:
            self.log_panel.add_error(f"✗ {file_name}: {error_msg}")

    def pause_conversion(self):
        """暂停转换"""
        if self.conversion_worker:
            self.conversion_worker.stop()
            self.log_panel.add_log("转换已暂停")

    def update_status(self, status: str):
        """更新状态栏"""
        self.statusBar().showMessage(f"状态: {status}")


# ============================================================================
# 工作流框架 - 向导式UI
# ============================================================================

class MainMenuFrame(QWidget):
    """主菜单框架 - 选择工作流类型"""

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        # 标题
        title = QLabel("📄 文档转换工具")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        subtitle = QLabel("选择工作流开始转换")
        subtitle_font = QFont()
        subtitle_font.setPointSize(12)
        subtitle.setFont(subtitle_font)
        subtitle.setStyleSheet("color: #666;")
        layout.addWidget(subtitle)

        layout.addSpacing(20)

        # 快速转换
        quick_btn = self.create_menu_button(
            "▶ 快速转换",
            "一键转换，使用默认设置\n无需配置，立即开始",
            "quick"
        )
        layout.addWidget(quick_btn)

        # 自定义转换
        custom_btn = self.create_menu_button(
            "⚙️ 自定义转换",
            "逐步配置文件、参数、执行\n精细控制每个环节",
            "custom"
        )
        layout.addWidget(custom_btn)

        # 批量转换
        batch_btn = self.create_menu_button(
            "📦 批量转换",
            "处理整个文件夹\n自动扫描并转换所有文件",
            "batch"
        )
        layout.addWidget(batch_btn)

        layout.addStretch()
        self.setLayout(layout)

    def create_menu_button(self, title: str, description: str, workflow_type: str):
        """创建菜单按钮"""
        btn = QPushButton()
        btn.setMinimumHeight(80)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                border: 2px solid #e0e0e0;
                border-radius: 8px;
                text-align: left;
                padding: 15px;
                font-size: 13px;
                color: #202124;
            }
            QPushButton:hover {
                background-color: #e8f0fe;
                border: 2px solid #1a73e8;
            }
        """)

        # 创建布局显示标题和描述
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        title_label = QLabel(title)
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(11)
        title_label.setFont(title_font)
        layout.addWidget(title_label)

        desc_label = QLabel(description)
        desc_font = QFont()
        desc_font.setPointSize(10)
        desc_label.setFont(desc_font)
        desc_label.setStyleSheet("color: #666;")
        layout.addWidget(desc_label)

        # 无法直接在QPushButton中嵌套布局，所以用QWidget包装
        container = QWidget()
        container.setLayout(layout)
        btn.clicked.connect(lambda: self.main_window.show_workflow(workflow_type))

        return btn


class BaseWorkflowFrame(QWidget):
    """工作流基类"""

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setup_ui()

    def setup_ui(self):
        """由子类实现"""
        raise NotImplementedError

    def go_back(self):
        """返回主菜单"""
        self.main_window.show_menu()

    def create_step_label(self, step_num: int, total_steps: int, step_name: str):
        """创建步骤标签"""
        label = QLabel(f"【步骤 {step_num}/{total_steps}】{step_name}")
        font = QFont()
        font.setBold(True)
        font.setPointSize(11)
        label.setFont(font)
        return label

    def create_back_button(self):
        """创建返回按钮"""
        btn = QPushButton("< 返回主菜单")
        btn.setMaximumWidth(120)
        btn.clicked.connect(self.go_back)
        return btn


class QuickConversionFrame(BaseWorkflowFrame):
    """快速转换工作流"""

    def __init__(self, main_window):
        super().__init__(main_window)

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # 返回按钮
        back_layout = QHBoxLayout()
        back_layout.addWidget(self.create_back_button())
        back_layout.addSpacing(10)
        title = QLabel("快速转换")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(13)
        title.setFont(title_font)
        back_layout.addWidget(title)
        back_layout.addStretch()
        layout.addLayout(back_layout)

        layout.addSpacing(15)

        # 文件选择
        layout.addWidget(self.create_step_label(1, 1, "选择文件"))
        file_layout = QHBoxLayout()
        file_btn = QPushButton("📄 浏览文件")
        file_btn.setMinimumHeight(40)
        file_layout.addWidget(file_btn)
        folder_btn = QPushButton("📁 浏览文件夹")
        folder_btn.setMinimumHeight(40)
        file_layout.addWidget(folder_btn)
        layout.addLayout(file_layout)

        # 文件列表
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: #999; padding: 10px;")
        layout.addWidget(self.file_label)

        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(120)
        layout.addWidget(self.file_list)

        self.files = []
        file_btn.clicked.connect(self.select_files)
        folder_btn.clicked.connect(self.select_folder)

        layout.addSpacing(20)

        # 转换按钮
        self.convert_btn = QPushButton("▶ 开始转换")
        self.convert_btn.setMinimumHeight(45)
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
        """)
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)

        layout.addStretch()
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
            self.update_file_list()

    def select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            folder_path = Path(folder)
            found_files = []
            for ext in {'.docx', '.pdf', '.xlsx'}:
                found_files.extend([str(f) for f in folder_path.glob(f'*{ext}')])

            if found_files:
                new_files = [f for f in found_files if f not in self.files]
                self.files.extend(new_files)
                self.update_file_list()
            else:
                QMessageBox.information(self, "提示", "文件夹中没有找到支持的文件")

    def update_file_list(self):
        """更新文件列表"""
        self.file_list.clear()
        self.file_label.setText(f"已选择 {len(self.files)} 个文件")
        for file in self.files:
            size = Path(file).stat().st_size / (1024 * 1024)
            self.file_list.addItem(f"{Path(file).name} ({size:.1f}MB)")

    def start_conversion(self):
        """开始转换"""
        if not self.files:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return

        # 使用默认输出目录
        output_dir = str(Path.home() / "转换结果")
        self.main_window.start_conversion(self.files, output_dir, 75)
        self.convert_btn.setEnabled(False)


class CustomConversionFrame(BaseWorkflowFrame):
    """自定义转换工作流"""

    def __init__(self, main_window):
        super().__init__(main_window)
        self.files = []
        self.current_step = 1

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # 返回按钮
        back_layout = QHBoxLayout()
        back_layout.addWidget(self.create_back_button())
        back_layout.addSpacing(10)
        title = QLabel("自定义转换")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(13)
        title.setFont(title_font)
        back_layout.addWidget(title)
        back_layout.addStretch()
        layout.addLayout(back_layout)

        layout.addSpacing(15)

        # ========== 步骤1：文件选择 ==========
        layout.addWidget(self.create_step_label(1, 3, "选择文件"))
        file_layout = QHBoxLayout()
        file_btn = QPushButton("📄 浏览文件")
        file_btn.setMinimumHeight(40)
        file_layout.addWidget(file_btn)
        folder_btn = QPushButton("📁 浏览文件夹")
        folder_btn.setMinimumHeight(40)
        file_layout.addWidget(folder_btn)
        layout.addLayout(file_layout)

        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: #999; padding: 10px;")
        layout.addWidget(self.file_label)

        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(100)
        layout.addWidget(self.file_list)

        file_btn.clicked.connect(self.select_files)
        folder_btn.clicked.connect(self.select_folder)

        layout.addSpacing(15)

        # ========== 步骤2：参数配置 ==========
        layout.addWidget(self.create_step_label(2, 3, "配置参数"))

        # 任务名称
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("任务名称:"))
        self.task_name = QLineEdit("新建任务")
        name_layout.addWidget(self.task_name)
        layout.addLayout(name_layout)

        # 输出目录
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_dir = QLineEdit(str(Path.home() / "转换结果"))
        output_layout.addWidget(self.output_dir)
        browse_btn = QPushButton("浏览...")
        browse_btn.setMaximumWidth(80)
        browse_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)

        # 质量阈值
        quality_layout = QHBoxLayout()
        quality_layout.addWidget(QLabel("质量阈值:"))
        self.quality_slider = QSlider(Qt.Orientation.Horizontal)
        self.quality_slider.setRange(0, 100)
        self.quality_slider.setValue(75)
        quality_layout.addWidget(self.quality_slider)
        self.quality_label = QLabel("75 分")
        self.quality_label.setMaximumWidth(50)
        quality_layout.addWidget(self.quality_label)
        self.quality_slider.valueChanged.connect(
            lambda v: self.quality_label.setText(f"{v} 分")
        )
        layout.addLayout(quality_layout)

        # 选项
        self.retry_check = QCheckBox("自动重试失败文件")
        self.retry_check.setChecked(True)
        layout.addWidget(self.retry_check)

        layout.addSpacing(15)

        # ========== 步骤3：执行转换 ==========
        layout.addWidget(self.create_step_label(3, 3, "执行转换"))

        self.convert_btn = QPushButton("▶ 开始转换")
        self.convert_btn.setMinimumHeight(45)
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
        """)
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)

        self.pause_btn = QPushButton("⏸ 暂停")
        self.pause_btn.setVisible(False)
        self.pause_btn.clicked.connect(self.main_window.pause_conversion)
        layout.addWidget(self.pause_btn)

        # 进度信息
        self.progress_label = QLabel("就绪")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        layout.addStretch()
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
            self.update_file_list()

    def select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            folder_path = Path(folder)
            found_files = []
            for ext in {'.docx', '.pdf', '.xlsx'}:
                found_files.extend([str(f) for f in folder_path.glob(f'*{ext}')])

            if found_files:
                new_files = [f for f in found_files if f not in self.files]
                self.files.extend(new_files)
                self.update_file_list()
            else:
                QMessageBox.information(self, "提示", "文件夹中没有找到支持的文件")

    def select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir.setText(dir_path)

    def update_file_list(self):
        """更新文件列表"""
        self.file_list.clear()
        self.file_label.setText(f"已选择 {len(self.files)} 个文件")
        for file in self.files:
            size = Path(file).stat().st_size / (1024 * 1024)
            self.file_list.addItem(f"{Path(file).name} ({size:.1f}MB)")

    def start_conversion(self):
        """开始转换"""
        if not self.files:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return

        self.main_window.start_conversion(
            self.files,
            self.output_dir.text(),
            self.quality_slider.value()
        )

        self.convert_btn.setVisible(False)
        self.pause_btn.setVisible(True)
        self.progress_bar.setVisible(True)

    def update_progress(self, current: int, total: int, status: str):
        """更新进度"""
        if total > 0:
            self.progress_bar.setValue(int(current * 100 / total))
            self.progress_label.setText(status)

    def on_conversion_finished(self, success: bool, message: str, stats: dict):
        """转换完成"""
        self.convert_btn.setVisible(True)
        self.pause_btn.setVisible(False)
        self.progress_bar.setVisible(False)

        if success:
            completed = stats.get('completed', 0)
            review = stats.get('review', 0)
            failed = stats.get('failed', 0)
            avg_quality = stats.get('avg_quality', 0.0)

            result_msg = (
                f"转换完成！\n"
                f"成功: {completed - review}\n"
                f"待审核: {review}\n"
                f"失败: {failed}\n"
                f"平均质量: {avg_quality:.1f}"
            )
            QMessageBox.information(self, "转换完成", result_msg)


class BatchConversionFrame(BaseWorkflowFrame):
    """批量转换工作流"""

    def __init__(self, main_window):
        super().__init__(main_window)
        self.folder = None

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # 返回按钮
        back_layout = QHBoxLayout()
        back_layout.addWidget(self.create_back_button())
        back_layout.addSpacing(10)
        title = QLabel("批量转换")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(13)
        title.setFont(title_font)
        back_layout.addWidget(title)
        back_layout.addStretch()
        layout.addLayout(back_layout)

        layout.addSpacing(15)

        # ========== 步骤1：选择文件夹 ==========
        layout.addWidget(self.create_step_label(1, 2, "选择文件夹"))

        folder_layout = QHBoxLayout()
        self.folder_btn = QPushButton("📁 浏览文件夹")
        self.folder_btn.setMinimumHeight(40)
        folder_layout.addWidget(self.folder_btn)
        layout.addLayout(folder_layout)

        self.folder_label = QLabel("未选择文件夹")
        self.folder_label.setStyleSheet("color: #999; padding: 10px;")
        layout.addWidget(self.folder_label)

        self.file_count_label = QLabel("")
        self.file_count_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(self.file_count_label)

        self.folder_btn.clicked.connect(self.select_folder)

        layout.addSpacing(15)

        # ========== 步骤2：输出配置 ==========
        layout.addWidget(self.create_step_label(2, 2, "配置输出"))

        # 输出目录
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_dir = QLineEdit(str(Path.home() / "转换结果"))
        output_layout.addWidget(self.output_dir)
        browse_btn = QPushButton("浏览...")
        browse_btn.setMaximumWidth(80)
        browse_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(browse_btn)
        layout.addLayout(output_layout)

        layout.addSpacing(20)

        # 转换按钮
        self.convert_btn = QPushButton("▶ 开始转换")
        self.convert_btn.setMinimumHeight(45)
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
        """)
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        layout.addStretch()
        self.setLayout(layout)

    def select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.folder = folder
            folder_path = Path(folder)
            self.folder_label.setText(f"📁 {folder_path.name}")

            # 扫描文件
            files = []
            for ext in {'.docx', '.pdf', '.xlsx'}:
                files.extend([str(f) for f in folder_path.glob(f'*{ext}')])

            self.file_count_label.setText(f"找到 {len(files)} 个支持的文件")

    def select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir.setText(dir_path)

    def start_conversion(self):
        """开始转换"""
        if not self.folder:
            QMessageBox.warning(self, "提示", "请先选择文件夹")
            return

        # 扫描文件
        folder_path = Path(self.folder)
        files = []
        for ext in {'.docx', '.pdf', '.xlsx'}:
            files.extend([str(f) for f in folder_path.glob(f'*{ext}')])

        if not files:
            QMessageBox.warning(self, "提示", "文件夹中没有找到支持的文件")
            return

        self.main_window.start_conversion(files, self.output_dir.text(), 75)
        self.convert_btn.setEnabled(False)
        self.progress_bar.setVisible(True)

    def update_progress(self, current: int, total: int, status: str):
        """更新进度"""
        if total > 0:
            self.progress_bar.setValue(int(current * 100 / total))



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