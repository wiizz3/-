import sys
import os
import psutil
import shutil
import json
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QTextEdit, QPushButton,
                             QMessageBox, QFileDialog, QProgressBar, QGroupBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon


def resource_path(relative_path):
    """获取资源的绝对路径 - 用于 PyInstaller 打包后定位资源"""
    try:
        # PyInstaller 会创建一个临时文件夹，路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    path = os.path.join(base_path, relative_path)
    return path


class ProcessWorker(QThread):
    """后台线程用于处理耗时操作"""
    finished = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, operation, file_manager, *args):
        super().__init__()
        self.operation = operation
        self.file_manager = file_manager
        self.args = args

    def run(self):
        try:
            if self.operation == "refresh":
                result = self.file_manager.refresh_status()
                self.finished.emit(result)
            elif self.operation == "kill_and_move":
                result = self.file_manager.kill_and_move()
                self.finished.emit(result)
            elif self.operation == "restore":
                result = self.file_manager.restore_file()
                self.finished.emit(result)
        except Exception as e:
            self.finished.emit(f"错误: {str(e)}")


class FileManager:
    """文件管理逻辑类"""

    def __init__(self):
        self.target_name = "8021x.exe"
        self.detected_process_path = None
        self.last_known_path = None
        self.temp_folder = "D:\\临时存放"
        self.original_path = None
        self.config_file = Path(self.temp_folder) / "original_path.json"
        self.load_original_path()

    def load_original_path(self):
        """从配置文件加载原始路径"""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.original_path = config.get('original_path')
                    return f"从配置文件加载原始路径: {self.original_path}"
        except Exception as e:
            return f"加载配置文件时出错: {e}"
        return ""

    def save_original_path(self):
        """保存原始路径到配置文件"""
        try:
            temp_dir = Path(self.temp_folder)
            if not temp_dir.exists():
                temp_dir.mkdir(parents=True)

            config = {
                'original_path': self.original_path,
                'moved_at': str(os.path.getctime(self.original_path)) if self.original_path and os.path.exists(
                    self.original_path) else None
            }

            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)

            return f"原始路径已保存到配置文件: {self.original_path}"
        except Exception as e:
            return f"保存配置文件时出错: {e}"

    def refresh_status(self):
        """刷新进程状态"""
        try:
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                if proc.info['name'].lower() == '8021x.exe':
                    processes.append(proc)

            if processes:
                self.detected_process_path = processes[0].info['exe']
                self.last_known_path = self.detected_process_path

                result = f"发现 {len(processes)} 个8021x.exe进程正在运行\n\n"
                for i, proc in enumerate(processes, 1):
                    result += f"进程 {i}:\n"
                    result += f"  PID: {proc.info['pid']}\n"
                    result += f"  路径: {proc.info['exe']}\n\n"
                return result
            else:
                self.detected_process_path = None
                result = "未发现8021x.exe进程正在运行\n\n"

                if self.original_path:
                    result += f"检测到之前移动的8021x.exe记录:\n"
                    result += f"原始路径: {self.original_path}\n"
                    result += f"文件当前位于: {self.temp_folder}\\8021x.exe\n"
                    result += f"您可以选择还原文件到原始位置"
                elif self.last_known_path:
                    result += f"最后已知路径: {self.last_known_path}"

                return result

        except Exception as e:
            return f"检查进程状态时出错: {e}"

    def kill_and_move(self):
        """结束进程并将文件移动到D盘的临时文件夹"""
        try:
            # 结束进程
            killed_count = 0
            processes_to_kill = []

            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                if proc.info['name'].lower() == '8021x.exe':
                    processes_to_kill.append(proc)

            for proc in processes_to_kill:
                try:
                    proc.kill()
                    killed_count += 1
                except Exception as e:
                    return f"无法结束进程 {proc.info['pid']}: {e}"

            if killed_count == 0:
                return "未发现8021x.exe进程正在运行"

            # 确定要移动的文件路径
            file_to_move = None
            if self.detected_process_path and os.path.exists(self.detected_process_path):
                file_to_move = self.detected_process_path
            elif self.last_known_path and os.path.exists(self.last_known_path):
                file_to_move = self.last_known_path
            else:
                return "无法确定8021x.exe文件位置"

            # 记录原始路径
            self.original_path = file_to_move
            save_result = self.save_original_path()

            # 创建临时文件夹
            temp_dir = Path(self.temp_folder)
            if not temp_dir.exists():
                try:
                    temp_dir.mkdir(parents=True)
                except Exception as e:
                    return f"无法在D盘创建临时文件夹: {e}"

            # 生成目标路径
            original_name = Path(file_to_move).name
            destination_path = temp_dir / original_name

            # 如果目标文件已存在，先删除
            if destination_path.exists():
                try:
                    destination_path.unlink()
                except Exception as e:
                    return f"无法删除已存在的文件 {destination_path}: {e}"

            # 移动文件
            try:
                shutil.move(file_to_move, destination_path)
                self.detected_process_path = None
                self.last_known_path = str(destination_path)

                result = f"操作成功完成！\n\n"
                result += f"已结束 {killed_count} 个8021x.exe进程\n"
                result += f"原文件: {file_to_move}\n"
                result += f"新位置: {destination_path}\n"
                result += f"临时文件夹: {temp_dir.absolute()}\n\n"
                result += f"{save_result}"

                return result

            except Exception as e:
                return f"移动文件时出错: {e}"

        except Exception as e:
            return f"操作过程中出错: {e}"

    def restore_file(self):
        """还原文件到原始位置"""
        try:
            # 检查临时文件夹
            temp_dir = Path(self.temp_folder)
            if not temp_dir.exists():
                return "临时文件夹不存在"

            # 检查原始路径记录
            if not self.original_path and self.config_file.exists():
                self.load_original_path()

            if not self.original_path:
                return "未找到原始路径记录，无法还原文件"

            # 查找临时文件夹中的文件
            temp_files = list(temp_dir.glob("8021x.exe"))
            if not temp_files:
                return "在临时文件夹中未找到8021x.exe文件"

            temp_file_path = temp_files[0]

            # 检查原始目录
            original_dir = Path(self.original_path).parent
            if not original_dir.exists():
                try:
                    original_dir.mkdir(parents=True)
                except Exception as e:
                    return f"无法创建原始目录 {original_dir}: {e}"

            # 如果原始位置已存在文件，先删除
            if os.path.exists(self.original_path):
                try:
                    os.remove(self.original_path)
                except Exception as e:
                    return f"无法删除已存在的文件 {self.original_path}: {e}"

            # 移动文件回原始位置
            try:
                shutil.move(str(temp_file_path), self.original_path)
                result = f"还原操作成功完成！\n\n"
                result += f"文件已从: {temp_file_path}\n"
                result += f"还原到: {self.original_path}\n"

                # 删除配置文件和临时文件夹
                try:
                    if self.config_file.exists():
                        self.config_file.unlink()
                        result += "配置文件已删除\n"

                    if not any(temp_dir.iterdir()):
                        temp_dir.rmdir()
                        result += "临时文件夹已删除\n"
                    else:
                        result += "临时文件夹不为空，保留文件夹\n"
                except Exception as e:
                    result += f"清理文件时出错: {e}\n"

                # 更新状态
                self.last_known_path = self.original_path
                self.detected_process_path = None
                self.original_path = None

                return result

            except Exception as e:
                return f"还原文件时出错: {e}"

        except Exception as e:
            return f"还原过程中出错: {e}"


class MainWindow(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()
        self.file_manager = FileManager()
        self.init_ui()

    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('8021x.exe 管理工具')
        self.setGeometry(100, 100, 700, 600)

        # 设置图标 - 使用 resource_path 确保打包后也能找到
        self.setup_icon()

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建主布局
        layout = QVBoxLayout(central_widget)

        # 标题
        title_label = QLabel('8021x.exe 管理工具')
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # 状态组
        status_group = QGroupBox("当前状态")
        status_layout = QVBoxLayout()

        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMaximumHeight(150)
        status_layout.addWidget(self.status_text)

        status_group.setLayout(status_layout)
        layout.addWidget(status_group)

        # 操作组
        action_group = QGroupBox("操作")
        action_layout = QVBoxLayout()

        # 按钮行1
        button_row1 = QHBoxLayout()
        self.refresh_btn = QPushButton('刷新状态')
        self.kill_move_btn = QPushButton('结束并移动到临时文件夹')
        self.restore_btn = QPushButton('还原文件')

        button_row1.addWidget(self.refresh_btn)
        button_row1.addWidget(self.kill_move_btn)
        button_row1.addWidget(self.restore_btn)

        # 按钮行2
        button_row2 = QHBoxLayout()
        self.exit_btn = QPushButton('退出')
        button_row2.addWidget(self.exit_btn)

        action_layout.addLayout(button_row1)
        action_layout.addLayout(button_row2)
        action_group.setLayout(action_layout)
        layout.addWidget(action_group)

        # 信息显示组
        info_group = QGroupBox("详细信息")
        info_layout = QVBoxLayout()

        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        info_layout.addWidget(self.info_text)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 连接信号和槽
        self.refresh_btn.clicked.connect(self.refresh_status)
        self.kill_move_btn.clicked.connect(self.kill_and_move)
        self.restore_btn.clicked.connect(self.restore_file)
        self.exit_btn.clicked.connect(self.close)

        # 初始状态
        self.refresh_status()

    def setup_icon(self):
        """设置应用程序图标"""
        # 使用 resource_path 获取图标路径
        icon_path = resource_path("icon.ico")

        # 如果找不到图标文件，尝试其他可能的路径
        if not os.path.exists(icon_path):
            # 尝试当前目录下的图标
            icon_path = "icon.ico"

        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            print(f"图标设置成功: {icon_path}")
        else:
            # 如果找不到图标文件，使用 Qt 内置图标作为备选
            from PyQt5.QtWidgets import QStyle
            self.setWindowIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
            print("使用内置图标")

    def refresh_status(self):
        """刷新状态"""
        self.progress_bar.setVisible(True)
        self.worker = ProcessWorker("refresh", self.file_manager)
        self.worker.finished.connect(self.on_operation_finished)
        self.worker.start()

    def kill_and_move(self):
        """结束进程并移动文件"""
        reply = QMessageBox.question(self, '确认操作',
                                     '确定要结束8021x.exe进程并将其移动到临时文件夹吗？',
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.progress_bar.setVisible(True)
            self.worker = ProcessWorker("kill_and_move", self.file_manager)
            self.worker.finished.connect(self.on_operation_finished)
            self.worker.start()

    def restore_file(self):
        """还原文件"""
        reply = QMessageBox.question(self, '确认操作',
                                     '确定要还原8021x.exe文件到原始位置吗？',
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.progress_bar.setVisible(True)
            self.worker = ProcessWorker("restore", self.file_manager)
            self.worker.finished.connect(self.on_operation_finished)
            self.worker.start()

    def on_operation_finished(self, result):
        """操作完成后的回调"""
        self.progress_bar.setVisible(False)
        self.info_text.setText(result)

        # 如果是刷新操作，更新状态文本
        if "发现" in result or "未发现" in result:
            self.status_text.setText(result.split('\n\n')[0])


if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 设置应用样式
    app.setStyle('Fusion')

    # 也可以在这里设置应用程序图标
    try:
        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
    except:
        pass

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())