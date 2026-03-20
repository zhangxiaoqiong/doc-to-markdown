#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import subprocess
import threading
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox, QProgressBar
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont


class ConversionWorker(QThread):
    """后台转换线程"""
    progress_signal = pyqtSignal(str, int, int)  # filename, current, total
    finished_signal = pyqtSignal(int, int, bool)  # success_count, failed_count, cancelled
    log_signal = pyqtSignal(str)  # log message

    def __init__(self, input_dir, output_dir):
        super().__init__()
        self.input_dir = input_dir
        self.output_dir = output_dir
        self._is_running = True

    def run(self):
        """后台线程执行转换"""
        success_count = 0
        failed_count = 0
        cancelled = False

        try:
            # 扫描输入目录中的文件
            files = self._get_files_to_convert()
            total = len(files)

            if total == 0:
                self.log_signal.emit("⚠️ 未找到待转换文件（PDF/DOCX/XLSX）")
                self.finished_signal.emit(0, 0, False)
                return

            # 直接调用 convert_all.py，让它处理整个文件夹
            self.log_signal.emit(f"📂 找到 {total} 个文件，开始批量转换...\n")

            try:
                self._convert_directory()
                success_count = total
                self.log_signal.emit(f"\n✓ 批量转换完成")
            except Exception as e:
                self.log_signal.emit(f"\n✗ 转换失败: {str(e)}")
                failed_count = total

        except Exception as e:
            self.log_signal.emit(f"❌ 转换异常: {str(e)}")
            failed_count += 1

        self.finished_signal.emit(success_count, failed_count, cancelled)

    def _get_files_to_convert(self):
        """获取待转换文件列表"""
        supported_exts = {'.pdf', '.docx', '.xlsx'}
        files = []

        if not os.path.isdir(self.input_dir):
            return files

        for filename in os.listdir(self.input_dir):
            if os.path.splitext(filename)[1].lower() in supported_exts:
                files.append(filename)

        return sorted(files)

    def _convert_directory(self):
        """调用convert_all.py转换整个目录，实时读取输出"""
        cmd = [
            sys.executable, "-u", "convert_all.py",  # -u 禁用缓冲
            "--input", self.input_dir,
            "--output", self.output_dir
        ]

        # 设置环境变量禁用缓冲
        env = os.environ.copy()
        env['PYTHONUNBUFFERED'] = '1'

        # 启动进程
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            env=env,
            bufsize=1,
            universal_newlines=True
        )

        # 启动线程读取输出（确保实时）
        def read_output():
            try:
                for line in iter(process.stdout.readline, ''):
                    if not self._is_running:
                        break
                    line = line.rstrip('\n')
                    if line.strip():
                        self.log_signal.emit(line)
            except:
                pass

        read_thread = threading.Thread(target=read_output, daemon=True)
        read_thread.start()

        # 等待进程完成
        process.wait()
        read_thread.join(timeout=1)

        if process.returncode != 0:
            raise Exception("转换过程出现错误，请查看日志")

    def stop(self):
        """停止转换"""
        self._is_running = False


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("📄 文档转换工具")
        self.setGeometry(100, 100, 1000, 750)
        self.setMinimumWidth(850)
        self.setMinimumHeight(650)

        # 设置窗口样式
        self.setStyleSheet("QMainWindow { background-color: #f5f7fa; }")

        self.input_dir = ""
        # 设置默认输出目录
        default_output_dir = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(default_output_dir, exist_ok=True)
        self.output_dir = default_output_dir

        self.worker = None

        self.init_ui()

        # 初始化后设置输出路径显示
        self.output_path.setText(default_output_dir)
        self.check_enable_convert()

    def init_ui(self):
        """初始化UI布局"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #f5f7fa;
            }
            QLabel {
                color: #1f2937;
            }
            QLineEdit {
                background-color: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 10px 12px;
                font-size: 10pt;
                selection-background-color: #3b82f6;
                color: #1f2937;
            }
            QLineEdit:focus {
                border: 2px solid #3b82f6;
                background-color: #ffffff;
            }
            QLineEdit::placeholder {
                color: #9ca3af;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: 600;
                font-size: 10pt;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
            QPushButton:disabled {
                background-color: #d1d5db;
                color: #9ca3af;
            }
            QProgressBar {
                background-color: #e5e7eb;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                height: 20px;
                text-align: center;
                color: #1f2937;
            }
            QProgressBar::chunk {
                background-color: #10b981;
                border-radius: 4px;
            }
        """)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(16)
        main_layout.setContentsMargins(24, 24, 24, 24)

        # ===== 标题区域 =====
        title_container = QWidget()
        title_layout = QVBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(6)

        title_label = QLabel("📄 文档转换工具")
        title_font = QFont("微软雅黑", 14, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #111827; margin: 0px;")
        title_layout.addWidget(title_label)

        subtitle_label = QLabel("支持格式: PDF • DOCX • XLSX")
        subtitle_font = QFont("微软雅黑", 9)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setStyleSheet("color: #6b7280; font-weight: normal;")
        title_layout.addWidget(subtitle_label)

        title_container.setLayout(title_layout)
        main_layout.addWidget(title_container)

        # ===== 分割线 =====
        divider = QLabel("")
        divider.setStyleSheet("background-color: #e5e7eb; height: 1px; margin: 8px 0px;")
        divider.setFixedHeight(1)
        main_layout.addWidget(divider)

        # ===== 输入文件夹 =====
        input_label_text = QLabel("📁 输入文件夹")
        input_label_text.setStyleSheet("color: #111827; font-weight: 600; font-size: 11pt;")
        input_label_text.setFont(QFont("微软雅黑", 11))
        main_layout.addWidget(input_label_text)

        input_layout = QHBoxLayout()
        input_layout.setSpacing(10)
        self.input_path = QLineEdit()
        self.input_path.setReadOnly(True)
        self.input_path.setPlaceholderText("选择包含待转换文档的文件夹")
        input_path_font = QFont("微软雅黑", 10)
        self.input_path.setFont(input_path_font)
        input_browse = QPushButton("浏览...")
        input_browse.setFixedWidth(90)
        input_browse.setFont(QFont("微软雅黑", 10))
        input_browse.clicked.connect(self.browse_input_dir)
        input_layout.addWidget(self.input_path, 1)
        input_layout.addWidget(input_browse)
        main_layout.addLayout(input_layout)

        # ===== 输出文件夹 =====
        output_label_text = QLabel("📂 输出文件夹")
        output_label_text.setStyleSheet("color: #111827; font-weight: 600; font-size: 11pt; margin-top: 8px;")
        output_label_text.setFont(QFont("微软雅黑", 11))
        main_layout.addWidget(output_label_text)

        output_layout = QHBoxLayout()
        output_layout.setSpacing(10)
        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        self.output_path.setPlaceholderText("选择存放转换结果的文件夹")
        self.output_path.setFont(input_path_font)
        output_browse = QPushButton("浏览...")
        output_browse.setFixedWidth(90)
        output_browse.setFont(QFont("微软雅黑", 10))
        output_browse.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(self.output_path, 1)
        output_layout.addWidget(output_browse)
        main_layout.addLayout(output_layout)

        # ===== 进度条 =====
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        main_layout.addWidget(self.progress_bar)

        # ===== 转换按钮 =====
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.convert_btn = QPushButton("🚀 开始转换")
        self.convert_btn.setFixedWidth(160)
        self.convert_btn.setFixedHeight(48)
        self.convert_btn.setEnabled(False)
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #059669;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-weight: 700;
                font-size: 11pt;
            }
            QPushButton:hover {
                background-color: #047857;
            }
            QPushButton:pressed {
                background-color: #065f46;
            }
            QPushButton:disabled {
                background-color: #d1d5db;
                color: #9ca3af;
            }
        """)
        self.convert_btn.clicked.connect(self.on_convert_clicked)
        button_font = QFont("微软雅黑", 11, QFont.Bold)
        self.convert_btn.setFont(button_font)
        button_layout.addWidget(self.convert_btn)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        # ===== 日志显示 =====
        log_label = QLabel("📋 转换日志")
        log_font = QFont("微软雅黑", 11, QFont.Bold)
        log_label.setFont(log_font)
        log_label.setStyleSheet("color: #111827; font-weight: 600; font-size: 11pt; margin-top: 12px;")
        main_layout.addWidget(log_label)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_font_mono = QFont("Consolas", 9)
        self.log_text.setFont(log_font_mono)
        # 优化日志显示样式
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 10px;
                line-height: 1.2;
            }
        """)
        # 设置行间距
        cursor = self.log_text.textCursor()
        fmt = cursor.blockFormat()
        fmt.setLineHeight(1.2, fmt.ProportionalHeight)
        cursor.setBlockFormat(fmt)
        main_layout.addWidget(self.log_text, 1)

        central_widget.setLayout(main_layout)

    def browse_input_dir(self):
        """选择输入文件夹"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输入文件夹")
        if dir_path:
            self.input_dir = dir_path
            self.input_path.setText(dir_path)
            self.check_enable_convert()

    def browse_output_dir(self):
        """选择输出文件夹"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出文件夹")
        if dir_path:
            self.output_dir = dir_path
            self.output_path.setText(dir_path)
            self.check_enable_convert()

    def check_enable_convert(self):
        """检查是否启用转换按钮"""
        enabled = bool(self.input_dir and self.output_dir)
        self.convert_btn.setEnabled(enabled)

    def on_convert_clicked(self):
        """转换按钮点击事件"""
        # 如果正在转换，点击则取消
        if self.worker is not None and self.worker.isRunning():
            self.log_text.append("\n⏹️  用户取消转换\n")
            self.worker.stop()
            self.worker.wait()
            self.convert_btn.setText("🚀 开始转换")
            return

        # 开始转换
        self.start_conversion()

    def start_conversion(self):
        """启动转换"""
        self.log_text.clear()
        self.log_text.append("📂 开始扫描文件...\n")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.worker = ConversionWorker(self.input_dir, self.output_dir)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.log_signal.connect(self.append_log)
        self.worker.finished_signal.connect(self.on_conversion_finished)

        self.convert_btn.setText("⏹️ 取消")
        self.convert_btn.setEnabled(True)

        self.worker.start()

    def update_progress(self, filename, current, total):
        """更新进度日志"""
        percentage = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(percentage)
        message = f"⏳ 处理中: {filename} ({current}/{total} 文件)"
        self.append_log(message)

    def append_log(self, message):
        """追加日志信息"""
        self.log_text.append(message)

    def on_conversion_finished(self, success_count, failed_count, cancelled):
        """转换完成回调"""
        self.convert_btn.setText("🚀 开始转换")
        self.progress_bar.setVisible(False)

        if cancelled:
            result_msg = "✋ 转换已取消"
        else:
            result_msg = f"✅ 转换完成\n\n成功: {success_count} 个\n失败: {failed_count} 个"

        self.log_text.append(f"\n{'='*50}\n{result_msg}\n{'='*50}\n")

        QMessageBox.information(self, "转换结果", result_msg)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
