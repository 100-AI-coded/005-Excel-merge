import os
import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QFileDialog, QListWidget, QCheckBox, 
                            QLineEdit, QProgressBar, QMessageBox, QGroupBox, QAbstractItemView,
                            QListWidgetItem, QSplashScreen, QRadioButton)  # 添加 QRadioButton
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QIcon, QDragEnterEvent, QDropEvent, QPixmap, QFont

# 添加内存监控功能
import psutil
import gc

class ExcelMergerThread(QThread):
    """处理Excel合并的线程类，避免UI卡顿"""
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    memory_signal = pyqtSignal(float)
    warning_signal = pyqtSignal(str)  # 添加警告信号
    
    def __init__(self, file_list, keep_all_headers, keep_first_header):
        super().__init__()
        self.file_list = file_list
        self.keep_all_headers = keep_all_headers
        self.keep_first_header = keep_first_header
        self.should_continue = None  # 添加标志位控制是否继续合并

    def run(self):
        try:
            if not self.file_list:
                self.error_signal.emit("请先选择Excel文件")
                return
                
            # 创建一个空的DataFrame用于存储合并结果
            merged_df = pd.DataFrame()
            total_files = len(self.file_list)
            
            # 如果选择只保留第一个文件的第一行，需要先检查所有文件的第一行是否一致
            if not self.keep_all_headers and self.keep_first_header:
                try:
                    # 读取第一个文件的第一行
                    first_file_df = pd.read_excel(self.file_list[0], engine='openpyxl', nrows=1, header=None)
                    first_row = first_file_df.iloc[0].tolist()
                    
                    # 检查其他文件的第一行
                    for i, file_path in enumerate(self.file_list[1:], 1):
                        current_df = pd.read_excel(file_path, engine='openpyxl', nrows=1, header=None)
                        current_row = current_df.iloc[0].tolist()
                        
                        if current_row != first_row:
                            self.warning_signal.emit(f"警告：{os.path.basename(file_path)}中的第一行与第一个文件的第一行内容不一致")
                            # 等待用户响应
                            self.should_continue = None
                            while self.should_continue is None:
                                self.msleep(100)
                            if not self.should_continue:
                                return
                except Exception as e:
                    self.error_signal.emit(f"检查文件第一行时出错: {str(e)}")
                    return
            
            # 继续原有的合并逻辑
            for i, file_path in enumerate(self.file_list):
                try:
                    # 更新进度
                    progress = int((i / total_files) * 100)
                    self.progress_signal.emit(progress)
                    
                    # 监控内存使用情况
                    memory_usage = psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024 / 1024  # GB
                    self.memory_signal.emit(memory_usage)
                    
                    # 使用chunks读取大文件以减少内存使用
                    file_size = os.path.getsize(file_path) / (1024 * 1024)  # MB
                    
                    if file_size > 100:  # 如果文件大于100MB，使用分块读取
                        chunks = pd.read_excel(file_path, engine='openpyxl', chunksize=10000, header=None)
                        first_chunk = next(chunks)
                        
                        # 处理第一行
                        if i > 0 and not self.keep_all_headers:  # 非第一个文件且不保留所有第一行
                            first_chunk = first_chunk.iloc[1:]

                        # 合并第一个chunk
                        if not first_chunk.empty:
                            merged_df = pd.concat([merged_df, first_chunk], ignore_index=True)
                        
                        # 处理剩余的chunks
                        for chunk in chunks:
                            merged_df = pd.concat([merged_df, chunk], ignore_index=True)
                            gc.collect()
                    else:
                        # 对于小文件，直接读取
                        df = pd.read_excel(file_path, engine='openpyxl', header=None)
                        
                        # 处理第一行
                        if i > 0 and not self.keep_all_headers:
                            df = df.iloc[1:]
                        
                        # 合并到结果DataFrame
                        merged_df = pd.concat([merged_df, df], ignore_index=True)
                    
                    # 强制垃圾回收
                    gc.collect()
                    
                except Exception as e:
                    self.error_signal.emit(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
                    continue
            
            if merged_df.empty:
                self.error_signal.emit("合并结果为空，请检查文件和设置")
                return
                
            # 保存合并后的文件
            output_dir = os.path.dirname(self.file_list[0])
            output_path = os.path.join(output_dir, "合并结果.xlsx")
            
            # 检查文件是否已存在，如果存在则添加序号
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(output_dir, f"合并结果_{counter}.xlsx")
                counter += 1
                
            # 使用较低内存的方式保存大文件
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, header=False, sheet_name='合并结果')
                
            self.progress_signal.emit(100)
            self.finished_signal.emit(output_path)
            
        except Exception as e:
            self.error_signal.emit(f"合并过程中发生错误: {str(e)}")
        finally:
            # 清理内存
            gc.collect()


class DragDropListWidget(QListWidget):
    """支持拖放操作的列表控件"""
    files_dropped = pyqtSignal(list)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setIconSize(QSize(16, 16))
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)
            
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            
            files = []
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.endswith(('.xlsx', '.xls')):
                    files.append(file_path)
            
            if files:
                self.files_dropped.emit(files)
        else:
            super().dropEvent(event)


class ExcelMergerApp(QMainWindow):
    """Excel合并工具主窗口"""
    def __init__(self):
        super().__init__()
        self.initUI()
        self.merger_thread = None
        self.excel_icon = QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "excel.png"))
        
        # 创建图标文件夹（如果不存在）
        icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
        if not os.path.exists(icon_dir):
            os.makedirs(icon_dir)
        
    def initUI(self):
        self.setWindowTitle("Excel表格合并器")
        self.setGeometry(100, 100, 900, 650)
        
        # 设置应用图标
        self.setWindowIcon(QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "app.png")))
        
        # 创建主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout()
        
        # 左侧文件列表区域
        left_layout = QVBoxLayout()
        
        file_group = QGroupBox("Excel文件列表")
        file_layout = QVBoxLayout()
        
        self.file_list = DragDropListWidget()
        self.file_list.setToolTip("可拖动调整文件合并顺序，从上到下依次合并")
        self.file_list.files_dropped.connect(self.add_dropped_files)
        
        file_buttons_layout = QHBoxLayout()
        self.add_btn = QPushButton("添加文件")
        self.add_btn.setIcon(QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "add.png")))
        self.add_btn.clicked.connect(self.add_files)
        
        self.remove_btn = QPushButton("移除选中")
        self.remove_btn.setIcon(QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "remove.png")))
        self.remove_btn.clicked.connect(self.remove_files)
        
        self.clear_btn = QPushButton("清空列表")
        self.clear_btn.setIcon(QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "clear.png")))
        self.clear_btn.clicked.connect(self.clear_files)
        
        file_buttons_layout.addWidget(self.add_btn)
        file_buttons_layout.addWidget(self.remove_btn)
        file_buttons_layout.addWidget(self.clear_btn)
        
        drag_label = QLabel("拖放Excel文件到此处或点击添加按钮:")
        drag_label.setStyleSheet("font-weight: bold;")
        
        file_layout.addWidget(drag_label)
        file_layout.addWidget(self.file_list)
        file_layout.addLayout(file_buttons_layout)
        file_group.setLayout(file_layout)
        
        left_layout.addWidget(file_group)
        
        # 右侧配置区域
        right_layout = QVBoxLayout()
        
        config_group = QGroupBox("合并配置")
        config_layout = QVBoxLayout()
        
        # 表头选项
        header_group = QGroupBox("第一行处理")
        header_layout = QVBoxLayout()
        
        self.keep_first_file_row = QRadioButton("只保留第一个文件的第一行")
        self.keep_first_file_row.setToolTip("合并后的文件只保留第一个Excel文件的第一行，其他文件的第一行将被删除")
        self.keep_first_file_row.setChecked(True)  # 默认选中
        
        self.keep_all_first_rows = QRadioButton("保留所有文件的第一行")
        self.keep_all_first_rows.setToolTip("合并后的文件将包含所有Excel文件的第一行")
        
        header_layout.addWidget(self.keep_first_file_row)
        header_layout.addWidget(self.keep_all_first_rows)
        header_group.setLayout(header_layout)
        
        # 添加配置组件到配置布局
        config_layout.addWidget(header_group)
        config_group.setLayout(config_layout)
        
        # 内存监控区域
        memory_group = QGroupBox("内存监控")
        memory_layout = QVBoxLayout()
        
        self.memory_label = QLabel("当前内存使用: 0 GB")
        memory_layout.addWidget(self.memory_label)
        memory_group.setLayout(memory_layout)
        
        # 合并按钮和进度条
        merge_group = QGroupBox("合并操作")
        merge_layout = QVBoxLayout()
        
        self.merge_btn = QPushButton("开始合并")
        self.merge_btn.setIcon(QIcon(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "merge.png")))
        self.merge_btn.clicked.connect(self.start_merge)
        self.merge_btn.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        merge_layout.addWidget(self.merge_btn)
        merge_layout.addWidget(self.progress_bar)
        merge_group.setLayout(merge_layout)
        
        # 添加到右侧布局
        right_layout.addWidget(config_group)
        right_layout.addWidget(memory_group)
        right_layout.addWidget(merge_group)
        right_layout.addStretch()
        
        # 设置主布局
        main_layout.addLayout(left_layout, 3)  # 左侧占3份
        main_layout.addLayout(right_layout, 2)  # 右侧占2份
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 状态栏
        self.statusBar().showMessage("准备就绪")
        
    def update_header_options(self, state):
        """更新表头选项的启用状态"""
        self.keep_first_header.setEnabled(not self.keep_all_headers.isChecked())
        
    def add_files(self):
        """添加Excel文件到列表"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if files:
            self.add_files_to_list(files)
            
    def add_dropped_files(self, files):
        """处理拖放添加的文件"""
        self.add_files_to_list(files)
        
    def add_files_to_list(self, files):
        """将文件添加到列表中并设置图标"""
        for file_path in files:
            item = QListWidgetItem(os.path.basename(file_path))
            item.setData(Qt.UserRole, file_path)  # 存储完整路径
            
            # 设置Excel图标
            try:
                item.setIcon(self.excel_icon)
            except:
                # 如果图标不存在，使用默认图标
                pass
                
            self.file_list.addItem(item)
            
    def remove_files(self):
        """从列表中移除选中的文件"""
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
            
    def clear_files(self):
        """清空文件列表"""
        self.file_list.clear()
        
    def get_file_list(self):
        """获取当前文件列表中的所有文件路径"""
        files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            # 获取存储的完整路径
            file_path = item.data(Qt.UserRole)
            files.append(file_path)
        return files
        
    def update_memory_usage(self, memory_gb):
        """更新内存使用显示"""
        self.memory_label.setText(f"当前内存使用: {memory_gb:.2f} GB")
        
        # 如果内存使用过高，显示警告
        if memory_gb > 3.0:  # 超过3GB显示警告
            self.memory_label.setStyleSheet("color: red; font-weight: bold;")
        elif memory_gb > 1.5:  # 超过1.5GB显示提醒
            self.memory_label.setStyleSheet("color: orange;")
        else:
            self.memory_label.setStyleSheet("")
        
    def start_merge(self):
        """开始合并操作"""
        files = self.get_file_list()
        if not files:
            QMessageBox.warning(self, "警告", "请先添加Excel文件")
            return
            
        # 获取配置选项
        keep_all_headers = self.keep_all_first_rows.isChecked()
        keep_first_header = not keep_all_headers
        
        # 禁用界面元素
        self.set_ui_enabled(False)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("正在合并...")
        
        # 创建并启动合并线程
        self.merger_thread = ExcelMergerThread(files, keep_all_headers, keep_first_header)
        self.merger_thread.progress_signal.connect(self.update_progress)
        self.merger_thread.finished_signal.connect(self.merge_completed)
        self.merger_thread.error_signal.connect(self.merge_error)
        self.merger_thread.memory_signal.connect(self.update_memory_usage)
        self.merger_thread.warning_signal.connect(self.handle_warning)  # 连接警告信号
        self.merger_thread.start()

    def handle_warning(self, warning_msg):
        """处理警告消息"""
        reply = QMessageBox.warning(
            self,
            "警告",
            f"{warning_msg}\n\n是否继续合并?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if hasattr(self, 'merger_thread') and self.merger_thread:
            self.merger_thread.should_continue = reply == QMessageBox.Yes
            
        if reply == QMessageBox.No:
            self.set_ui_enabled(True)
            self.statusBar().showMessage("合并已取消")
        
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
        
    def merge_completed(self, output_path):
        """合并完成后的处理"""
        self.set_ui_enabled(True)
        self.statusBar().showMessage(f"合并完成: {output_path}")
        
        reply = QMessageBox.information(
            self, 
            "合并完成", 
            f"Excel文件已成功合并，保存至:\n{output_path}\n\n是否打开文件?", 
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            os.startfile(output_path)
            
    def merge_error(self, error_msg):
        """处理合并过程中的错误"""
        self.set_ui_enabled(True)
        self.statusBar().showMessage("合并失败")
        QMessageBox.critical(self, "合并错误", error_msg)
        
    def set_ui_enabled(self, enabled):
        """设置UI元素的启用状态"""
        self.add_btn.setEnabled(enabled)
        self.remove_btn.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled)
        self.keep_all_first_rows.setEnabled(enabled)
        self.keep_first_file_row.setEnabled(enabled)
        self.merge_btn.setEnabled(enabled)
        self.file_list.setEnabled(enabled)
        
    def closeEvent(self, event):
        """关闭窗口时的处理"""
        if self.merger_thread and self.merger_thread.isRunning():
            reply = QMessageBox.question(
                self, 
                "确认退出", 
                "合并操作正在进行中，确定要退出吗?", 
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.merger_thread.terminate()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


# 创建图标生成函数
def create_default_icons():
    """创建默认图标文件"""
    icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
    if not os.path.exists(icon_dir):
        os.makedirs(icon_dir)
        
    # 这里可以添加代码生成简单的图标文件
    # 由于无法直接创建图像文件，这里只创建目录结构


if __name__ == "__main__":
    # 设置高DPI支持
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
        
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格，在所有平台上看起来一致
    
    # 创建默认图标
    create_default_icons()
    
    # 显示启动画面
    splash_pix = QPixmap(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "splash.png"))
    if not splash_pix.isNull():
        splash = QSplashScreen(splash_pix)
        splash.show()
        app.processEvents()
    
    # 创建并显示主窗口
    window = ExcelMergerApp()
    window.show()
    
    # 如果有启动画面，关闭它
    if 'splash' in locals():
        splash.finish(window)
    
    sys.exit(app.exec_())