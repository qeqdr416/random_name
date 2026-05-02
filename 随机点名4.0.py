import sys
import random
import json
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout,
    QGraphicsDropShadowEffect, QDialog, QFormLayout, QSpinBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QHBoxLayout,
    QLineEdit, QDoubleSpinBox, QFileDialog, QMessageBox,
    QCheckBox, QTextEdit
)
from PyQt6.QtCore import Qt, QTimer, QPoint, QEvent, QUrl, QStandardPaths
from PyQt6.QtGui import QFont, QColor, QPixmap, QDesktopServices
from PyQt6.QtMultimedia import QSoundEffect
from PyQt6.QtNetwork import QLocalServer, QLocalSocket, QNetworkAccessManager, QNetworkRequest

# 尝试导入语音库（Qt6 支持）
try:
    from PyQt6.QtTextToSpeech import QTextToSpeech
    TTS_AVAILABLE = True
except ImportError:
    TTS_AVAILABLE = False

# 尝试导入 Excel 处理库（可选依赖）
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False

# -------------------- 全局常量 --------------------
VERSION = "4.0"  # 当前版本号
UPDATE_URL = "https://raw.githubusercontent.com/qeqdr416/random_name/refs/heads/main/version.json"  # 更新检测地址（可自行修改）

NAME_LIST = [
    "陈柏林", "陈楚轩", "陈秀瑶", "冯锦怡", "郭雨娴", "何家添", "何洛明", "何泳霖", "何子轩",
    "黄杰娴", "黄淑欣", "黄钰煊", "黄正杰", "黄政宇", "江钒", "赖霆骏", "赖心沅", "黎家鸿",
    "李浩基", "李贤岚", "李欣怡", "李泳妍", "梁恒之", "梁家荣", "梁凯晴", "梁乐儿",
    "梁铭烯", "梁梓瑜", "廖恺睿", "刘柏良", "刘绮童", "罗俊文", "蒙彩琳", "任语涵",
    "唐瑜彤", "翁舒雨", "吴辉健", "吴俊沛", "吴思淇", "伍柏裕", "杨宇悦", "张玉欣",
    "郑洁瑜", "朱虹炎", "朱紫涵", "黄宏轩", "彭威豪", "谢韵晴", "梁芷晴", "房尚炜"
]

LEFT_BOTTOM_IMAGE_PATH = "logo.png"
IMAGE_SIZE = (120, 120)
CLICK_SOUND_PATH = "click.wav"

DATA_DIR = os.path.join(QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation),
                        "RollCallApp")
DATA_FILE = os.path.join(DATA_DIR, "rollcall_data.json")
BACKUP_FILE = os.path.join(DATA_DIR, "rollcall_data.json.bak")

DEFAULT_REFRESH_INTERVAL = 20
DEFAULT_COOLDOWN_TIMES = 5
DEFAULT_NAME_LABEL_HEIGHT = 180
DEFAULT_NAME_FONT_SIZE = 45
DEFAULT_NAME_LABEL_WIDTH = 0
DEFAULT_CONTROL_BTN_WIDTH = 0
DEFAULT_DRAW_COUNT = 1
SERVER_NAME = "RollCallAppServer"


class RollCallApp(QWidget):
    def __init__(self):
        super().__init__()
        self.load_data()
        self.init_window()
        self.init_left_bottom_image()
        self.init_ui()
        self.init_variables()
        self.init_sound()
        self.init_speech()          # 初始化语音引擎
        self.init_local_server()
        self.set_style()
        self.reposition_controls()
        self.current_names = []     # 当前显示的名字列表（用于朗读）
        # 延迟启动更新检测，避免阻塞界面加载
        QTimer.singleShot(2000, self.check_for_update)

    # ---------- 窗口初始化 ----------
    def init_window(self):
        self.setWindowTitle(self.window_title)
        self.resize(self.window_width, self.window_height)
        self.setMinimumSize(400, 300)
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def init_left_bottom_image(self):
        if not LEFT_BOTTOM_IMAGE_PATH:
            return
        image_pix = QPixmap(LEFT_BOTTOM_IMAGE_PATH)
        if image_pix.isNull():
            print(f"图片加载失败，请检查路径：{LEFT_BOTTOM_IMAGE_PATH}")
            return
        self.image_label = QLabel(self)
        self.image_label.setObjectName("image_label")
        scaled_image = image_pix.scaled(
            IMAGE_SIZE[1], IMAGE_SIZE[1],
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        self.image_label.setPixmap(scaled_image)
        self.image_label.setFixedSize(IMAGE_SIZE[0], IMAGE_SIZE[1])
        self.image_label.lower()
        self.image_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

    def init_variables(self):
        self.is_running = False
        self.float_win = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_name)
        self.auto_stop_timer = QTimer(self)
        self.auto_stop_timer.setSingleShot(True)
        self.auto_stop_timer.timeout.connect(self.auto_stop)

    def init_sound(self):
        self.click_sound = QSoundEffect(self)
        if os.path.exists(CLICK_SOUND_PATH):
            self.click_sound.setSource(QUrl.fromLocalFile(os.path.abspath(CLICK_SOUND_PATH)))
            self.click_sound.setVolume(1.0)
        else:
            self.click_sound = None

    def play_click_sound(self):
        if self.click_sound:
            self.click_sound.play()

    # ---------- 语音朗读初始化 ----------
    def init_speech(self):
        self.speech = None
        self.speech_available = False
        self.speech_queue = []
        self.speech_index = 0
        if TTS_AVAILABLE:
            try:
                self.speech = QTextToSpeech(self)
                # 检查引擎是否可用
                if self.speech.state() == QTextToSpeech.State.Error:
                    self.speech = None
                else:
                    self.speech_available = True
            except Exception:
                self.speech = None

    def start_speaking(self, names):
        """朗读一组名字，逐个进行"""
        if not self.speech_available or not self.speech or not names:
            return
        # 停止之前的朗读
        self.stop_speaking()
        self.speech_queue = list(names)
        self.speech_index = 0
        try:
            self.speech.stateChanged.disconnect(self.on_speech_state_changed)
        except TypeError:
            pass
        self.speech.stateChanged.connect(self.on_speech_state_changed)
        self.speech.say(self.speech_queue[0])

    def on_speech_state_changed(self, state):
        if state == QTextToSpeech.State.Ready:
            self.speech_index += 1
            if self.speech_index < len(self.speech_queue):
                self.speech.say(self.speech_queue[self.speech_index])
            else:
                try:
                    self.speech.stateChanged.disconnect(self.on_speech_state_changed)
                except TypeError:
                    pass

    def stop_speaking(self):
        if self.speech_available and self.speech:
            self.speech.stop()
        self.speech_queue = []
        self.speech_index = 0

    # ---------- 单实例检测 ----------
    def init_local_server(self):
        self.server = QLocalServer(self)
        self.server.newConnection.connect(self.handle_new_connection)
        QLocalServer.removeServer(SERVER_NAME)
        if not self.server.listen(SERVER_NAME):
            print("警告：无法启动本地服务器，进程检测可能失效")

    def handle_new_connection(self):
        client_connection = self.server.nextPendingConnection()
        if client_connection:
            client_connection.readyRead.connect(lambda: self.read_command(client_connection))

    def read_command(self, conn):
        if conn.bytesAvailable():
            data = conn.readAll().data().decode()
            if data.strip() == "show":
                self.restore_from_float()
            conn.disconnectFromServer()

    # ---------- 数据持久化 ----------
    def safe_load_json(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return None

    def load_data(self):
        os.makedirs(DATA_DIR, exist_ok=True)
        data = self.safe_load_json(DATA_FILE)
        if data is None:
            data = self.safe_load_json(BACKUP_FILE)
            if data is not None:
                try:
                    with open(DATA_FILE, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
        if data is None:
            self._init_defaults()
            return

        try:
            self.draw_counts = data.get("draw_counts", {})
            self.cool_down = data.get("cooldown", {})
            self.refresh_interval = int(data.get("refresh_interval", DEFAULT_REFRESH_INTERVAL))
            self.cooldown_times = int(data.get("cooldown_times", DEFAULT_COOLDOWN_TIMES))
            self.window_width = int(data.get("window_width", 600))
            self.window_height = int(data.get("window_height", 450))
            self.window_title = data.get("window_title", "综合高中252班随机点名程序")
            self.float_opacity = float(data.get("float_opacity", 0.95))
            self.name_list = data.get("name_list", None)
            self.imported_file_path = data.get("imported_file_path", None)
            self.name_label_height = int(data.get("name_label_height", DEFAULT_NAME_LABEL_HEIGHT))
            self.name_font_size = int(data.get("name_font_size", DEFAULT_NAME_FONT_SIZE))
            self.name_label_width = int(data.get("name_label_width", DEFAULT_NAME_LABEL_WIDTH))
            self.control_btn_width = int(data.get("control_btn_width", DEFAULT_CONTROL_BTN_WIDTH))
            self.auto_mode = data.get("auto_mode", False)
            self.draw_count = int(data.get("draw_count", DEFAULT_DRAW_COUNT))
            self.read_aloud_enabled = data.get("read_aloud_enabled", False)
        except Exception:
            self._init_defaults()

        if not self.name_list or len(self.name_list) == 0:
            self.name_list = NAME_LIST[:]

        if self.imported_file_path:
            if not os.path.exists(self.imported_file_path):
                self.name_list = NAME_LIST[:]
                self.imported_file_path = None
            else:
                try:
                    with open(self.imported_file_path, 'r', encoding='utf-8') as f:
                        current_names = [line.strip() for line in f if line.strip()]
                    if not current_names or current_names != self.name_list:
                        self.name_list = NAME_LIST[:]
                        self.imported_file_path = None
                except Exception:
                    self.name_list = NAME_LIST[:]
                    self.imported_file_path = None
            self.save_data()

    def _init_defaults(self):
        self.draw_counts = {}
        self.cool_down = {}
        self.refresh_interval = DEFAULT_REFRESH_INTERVAL
        self.cooldown_times = DEFAULT_COOLDOWN_TIMES
        self.window_width = 600
        self.window_height = 450
        self.window_title = "252班随机点名程序"
        self.float_opacity = 0.95
        self.name_list = NAME_LIST[:]
        self.imported_file_path = None
        self.name_label_height = DEFAULT_NAME_LABEL_HEIGHT
        self.name_font_size = DEFAULT_NAME_FONT_SIZE
        self.name_label_width = DEFAULT_NAME_LABEL_WIDTH
        self.control_btn_width = DEFAULT_CONTROL_BTN_WIDTH
        self.auto_mode = False
        self.draw_count = DEFAULT_DRAW_COUNT
        self.read_aloud_enabled = False

    def save_data(self):
        data = {
            "draw_counts": self.draw_counts,
            "cooldown": self.cool_down,
            "refresh_interval": self.refresh_interval,
            "cooldown_times": self.cooldown_times,
            "window_width": self.width(),
            "window_height": self.height(),
            "window_title": self.window_title,
            "float_opacity": self.float_opacity,
            "name_list": self.name_list,
            "imported_file_path": self.imported_file_path,
            "name_label_height": self.name_label_height,
            "name_font_size": self.name_font_size,
            "name_label_width": self.name_label_width,
            "control_btn_width": self.control_btn_width,
            "auto_mode": self.auto_mode,
            "draw_count": self.draw_count,
            "read_aloud_enabled": self.read_aloud_enabled,
        }
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            temp_file = DATA_FILE + ".tmp"
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if os.path.exists(DATA_FILE):
                try:
                    os.replace(DATA_FILE, BACKUP_FILE)
                except Exception:
                    pass
            os.replace(temp_file, DATA_FILE)
        except Exception as e:
            print(f"保存数据失败: {e}")

    # ---------- 界面构建 ----------
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(50, 30, 50, 20)
        layout.setSpacing(0)
        self.setLayout(layout)

        self.name_label = QLabel("保证公平")
        self.name_label.setObjectName("name_label")
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.apply_name_label_style()
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(12)
        shadow.setColor(QColor(180, 180, 180))
        shadow.setOffset(4, 4)
        self.name_label.setGraphicsEffect(shadow)
        layout.addWidget(self.name_label)

        layout.addSpacing(30)

        btn_text = "自动抽选" if self.auto_mode else "点击抽选"
        self.control_btn = QPushButton(btn_text)
        control_font = QFont("黑体", 30)
        control_font.setWeight(QFont.Weight.Light)
        self.control_btn.setFont(control_font)
        self.control_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.control_btn.setMinimumHeight(100)
        self.apply_control_btn_width()
        self.control_btn.clicked.connect(self.toggle_status)
        self.control_btn.clicked.connect(self.play_click_sound)
        layout.addWidget(self.control_btn)

        layout.addSpacing(50)

        self.tip_label1 = QLabel("latest_version:4.0 项目已开源\n会自动引用Github仓库检查更新")
        self.tip_label1.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tip_label1.setFont(QFont("Microsoft YaHei", 12))
        self.tip_label1.setObjectName("tip_label")
        layout.addWidget(self.tip_label1)

        layout.addSpacing(5)

        self.tip_label2 = QLabel("点击最小化窗口创建悬浮窗")
        self.tip_label2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tip_label2.setFont(QFont("Microsoft YaHei", 12))
        self.tip_label2.setObjectName("tip_label")
        layout.addWidget(self.tip_label2)

        layout.addStretch()

        self.settings_btn = QPushButton("⚙", self)
        self.settings_btn.setObjectName("config_btn")
        self.settings_btn.setFixedSize(45, 45)
        self.settings_btn.setFont(QFont("黑体", 15))
        self.settings_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.settings_btn.setToolTip("配置")
        self.settings_btn.clicked.connect(self.open_config_dialog)
        self.settings_btn.clicked.connect(self.play_click_sound)

        self.stats_btn = QPushButton("📊", self)
        self.stats_btn.setObjectName("stats_btn")
        self.stats_btn.setFixedSize(45, 45)
        self.stats_btn.setFont(QFont("黑体", 15))
        self.stats_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.stats_btn.setToolTip("抽取统计")
        self.stats_btn.clicked.connect(self.show_statistics)
        self.stats_btn.clicked.connect(self.play_click_sound)

    def apply_name_label_style(self):
        name_font = QFont("黑体", self.name_font_size)
        name_font.setWeight(QFont.Weight.Light)
        self.name_label.setFont(name_font)
        self.name_label.setFixedHeight(self.name_label_height)
        if self.name_label_width > 0:
            self.name_label.setFixedWidth(self.name_label_width)
        else:
            self.name_label.setMinimumWidth(0)
            self.name_label.setMaximumWidth(16777215)

    def apply_control_btn_width(self):
        if self.control_btn_width > 0:
            self.control_btn.setFixedWidth(self.control_btn_width)
        else:
            self.control_btn.setMinimumWidth(0)
            self.control_btn.setMaximumWidth(16777215)

    def set_style(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f1f5f9;
            }
            QLabel {
                background-color: #ffffff;
                border: 2px solid #cbd5e1;
                border-radius: 12px;
                color: #0f172a;
            }
            QLabel#name_label {
                background-color: #ffffff;
                border: 3px solid #64748b;
                border-radius: 20px;
                color: #1e3a8a;
            }
            QLabel#tip_label {
                background-color: transparent;
                border: none;
                color: #64748b;
                padding: 0;
                min-height: 20px;
            }
            QLabel#image_label {
                background-color: transparent;
                border: none;
                border-radius: 0;
                padding: 0;
                margin: 0;
            }
            QPushButton {
                background-color: #4a9eff;
                color: white;
                border: none;
                border-radius: 16px;
                padding: 12px 0;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #2f8cff;
            }
            QPushButton:pressed {
                background-color: #1e40af;
                padding-left: 3px;
                padding-top: 3px;
            }
            QPushButton#config_btn, QPushButton#stats_btn {
                background-color: #64748b;
                color: #f8fafc;
                border: none;
                border-radius: 22px;
                min-width: 45px;
                max-width: 45px;
                min-height: 45px;
                max-height: 45px;
            }
            QPushButton#config_btn:hover, QPushButton#stats_btn:hover {
                background-color: #475569;
            }
            QCheckBox {
                spacing: 8px;
                font-size: 14px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border: 2px solid #64748b;
                border-radius: 4px;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #4a9eff;
                border-color: #4a9eff;
            }
        """)

    # ---------- 随机抽取逻辑 ----------
    def generate_random_names(self):
        available = [name for name in self.name_list if name not in self.cool_down]
        total_count = min(self.draw_count, 12)
        if len(available) >= total_count:
            chosen = random.sample(available, total_count)
        else:
            pool = self.name_list[:]
            if len(pool) >= total_count:
                chosen = random.sample(pool, total_count)
            else:
                chosen = [random.choice(pool) for _ in range(total_count)]
        return chosen

    def format_names_for_display(self, names):
        lines = []
        for i in range(0, len(names), 4):
            group = names[i:i + 4]
            formatted_group = []
            for name in group:
                if len(name) > 8:
                    formatted_group.append('\u200b'.join(name))
                else:
                    formatted_group.append(name)
            line = ' '.join(formatted_group)
            lines.append(line)
        return '\n'.join(lines)

    def toggle_status(self):
        # 开始新抽选时停止朗读
        self.stop_speaking()

        if self.auto_mode:
            if not self.is_running:
                self.is_running = True
                self.timer.start(self.refresh_interval)
                delay = random.randint(400, 1200)
                self.auto_stop_timer.start(delay)
                self.control_btn.setText("自动抽选中...")
                self.control_btn.setEnabled(False)
                self.play_click_sound()
            return

        if not self.is_running:
            self.is_running = True
            self.control_btn.setText("暂停")
            self.timer.start(self.refresh_interval)
        else:
            self.is_running = False
            self.control_btn.setText("继续")
            self.timer.stop()
            self.apply_cooldown_and_count()

    def auto_stop(self):
        if not self.is_running:
            return
        self.timer.stop()
        self.is_running = False
        self.apply_cooldown_and_count()
        self.control_btn.setText("自动抽选")
        self.control_btn.setEnabled(True)

    def apply_cooldown_and_count(self):
        """抽选结束后的处理：冷却、计数、触发朗读"""
        names = self.current_names  # 使用当前显示的名字，避免重复解析
        # 更新冷却池
        for name in list(self.cool_down.keys()):
            self.cool_down[name] -= 1
            if self.cool_down[name] <= 0:
                del self.cool_down[name]
        for name in names:
            self.draw_counts[name] = self.draw_counts.get(name, 0) + 1
            self.cool_down[name] = self.cooldown_times
        self.save_data()

        # 朗读抽取结果（如果启用）
        if self.read_aloud_enabled and names:
            self.start_speaking(names)

    def update_name(self):
        names = self.generate_random_names()
        self.current_names = names
        display_text = self.format_names_for_display(names)
        self.name_label.setText(display_text)

    # ---------- 窗口事件 ----------
    def changeEvent(self, event):
        if event.type() == QEvent.Type.WindowStateChange:
            if self.isMinimized():
                self.min_to_float()
        super().changeEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.reposition_controls()

    def min_to_float(self):
        self.hide()
        if self.float_win:
            self.float_win.close()
            self.float_win = None
        self.float_win = FloatWindow(self)
        self.float_win.show()

    def restore_from_float(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()
        if self.float_win:
            self.float_win.close()
            self.float_win = None

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self, "确认退出",
            "确定要退出点名程序吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.timer.stop()
            self.auto_stop_timer.stop()
            self.stop_speaking()
            if self.float_win:
                self.float_win.close()
            self.server.close()
            self.save_data()
            event.accept()
        else:
            event.ignore()

    def reposition_controls(self):
        if hasattr(self, 'settings_btn'):
            self.settings_btn.move(self.width() - 55, 15)
            self.stats_btn.move(self.width() - 55, 85)
        if hasattr(self, 'image_label'):
            pos_y = self.height() - IMAGE_SIZE[1] - 10
            self.image_label.move(20, pos_y)

    # ---------- 更新检测 ----------
    def check_for_update(self):
        manager = QNetworkAccessManager(self)
        reply = manager.get(QNetworkRequest(QUrl(UPDATE_URL)))
        reply.finished.connect(lambda: self.on_update_reply(reply))

    def on_update_reply(self, reply):
        try:
            data = json.loads(reply.readAll().data().decode())
            latest = data.get("version", "")
            download_url = data.get("url", "")
            if latest and latest > VERSION:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Information)
                msg.setWindowTitle("发现新版本")
                msg.setText(f"当前版本：{VERSION}\n最新版本：{latest}")
                msg.setInformativeText("是否前往下载页面？")
                msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if msg.exec() == QMessageBox.StandardButton.Yes and download_url:
                    QDesktopServices.openUrl(QUrl(download_url))
        except Exception:
            pass  # 网络错误或解析失败时静默处理

    # ---------- 导入功能（支持 TXT 与 Excel） ----------
    def import_name_list(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入抽奖名单", "",
            "文本文件 (*.txt);;Excel文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if not file_path:
            return
        try:
            if file_path.lower().endswith('.xlsx'):
                if not OPENPYXL_AVAILABLE:
                    QMessageBox.critical(self, "错误", "需要安装 openpyxl 库才能读取 .xlsx 文件。\n请执行：pip install openpyxl")
                    return
                wb = openpyxl.load_workbook(file_path)
                ws = wb.active
                new_names = [str(cell.value).strip() for cell in ws['A'] if cell.value is not None]
                wb.close()
            elif file_path.lower().endswith('.xls'):
                if not XLRD_AVAILABLE:
                    QMessageBox.critical(self, "错误", "需要安装 xlrd 库才能读取 .xls 文件。\n请执行：pip install xlrd")
                    return
                wb = xlrd.open_workbook(file_path)
                ws = wb.sheet_by_index(0)
                new_names = [str(ws.cell_value(row, 0)).strip() for row in range(ws.nrows) if ws.cell_value(row, 0) != '']
            else:
                # 默认按 TXT 处理
                with open(file_path, 'r', encoding='utf-8') as f:
                    new_names = [line.strip() for line in f if line.strip()]

            if new_names:
                self.name_list = new_names
                self.imported_file_path = file_path
                self.save_data()
                QMessageBox.information(self, "导入成功", f"已成功导入 {len(self.name_list)} 个名字！")
            else:
                QMessageBox.warning(self, "导入失败", "文件中没有找到有效的名字。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入失败：{e}")

    # ---------- 手动添加名字对话框 ----------
    def add_names_manually(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("手动添加名字")
        dialog.setFixedSize(420, 320)
        layout = QVBoxLayout(dialog)
        label = QLabel("请输入要添加的名字，每行一个：")
        label.setFont(QFont("Microsoft YaHei", 10))
        text_edit = QTextEdit()
        layout.addWidget(label)
        layout.addWidget(text_edit)

        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("确定添加")
        cancel_btn = QPushButton("取消")
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def on_ok():
            raw = text_edit.toPlainText().strip()
            if not raw:
                QMessageBox.warning(dialog, "提示", "没有输入任何名字。")
                return
            new = [n.strip() for n in raw.splitlines() if n.strip()]
            if new:
                # 添加并去重（保持原有顺序）
                existing = set(self.name_list)
                for n in new:
                    if n not in existing:
                        self.name_list.append(n)
                        existing.add(n)
                self.save_data()
                QMessageBox.information(dialog, "成功", f"已添加 {len(new)} 个名字。")
                dialog.accept()
            else:
                QMessageBox.warning(dialog, "提示", "未识别到有效名字。")

        ok_btn.clicked.connect(on_ok)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec()

    # ---------- 配置对话框（新增朗读开关、手动添加按钮）----------
    def open_config_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("系统配置")
        dialog.setFixedSize(420, 720)

        main_layout = QVBoxLayout(dialog)
        form = QFormLayout()

        interval_spin = QSpinBox()
        interval_spin.setRange(5, 200)
        interval_spin.setValue(self.refresh_interval)
        interval_spin.setSuffix(" 毫秒")
        form.addRow("名字滚动刷新间隔:", interval_spin)

        cooldown_spin = QSpinBox()
        cooldown_spin.setRange(1, 30)
        cooldown_spin.setValue(self.cooldown_times)
        cooldown_spin.setSuffix(" 次")
        form.addRow("防重复冷却次数:", cooldown_spin)

        draw_count_spin = QSpinBox()
        draw_count_spin.setRange(1, 12)
        draw_count_spin.setValue(self.draw_count)
        draw_count_spin.setSuffix(" 人")
        form.addRow("每次抽选人数:", draw_count_spin)

        title_edit = QLineEdit(self.window_title)
        form.addRow("程序标题:", title_edit)

        opacity_spin = QDoubleSpinBox()
        opacity_spin.setRange(0.3, 1.0)
        opacity_spin.setSingleStep(0.05)
        opacity_spin.setDecimals(2)
        opacity_spin.setValue(self.float_opacity)
        form.addRow("悬浮窗透明度:", opacity_spin)

        label_width_spin = QSpinBox()
        label_width_spin.setRange(0, 1000)
        label_width_spin.setValue(self.name_label_width)
        label_width_spin.setSpecialValueText("自动")
        label_width_spin.setSuffix(" px")
        form.addRow("名字标签宽度:", label_width_spin)

        label_height_spin = QSpinBox()
        label_height_spin.setRange(60, 400)
        label_height_spin.setValue(self.name_label_height)
        label_height_spin.setSuffix(" px")
        form.addRow("名字标签高度:", label_height_spin)

        font_size_spin = QSpinBox()
        font_size_spin.setRange(12, 100)
        font_size_spin.setValue(self.name_font_size)
        font_size_spin.setSuffix(" px")
        form.addRow("名字字体大小:", font_size_spin)

        btn_width_spin = QSpinBox()
        btn_width_spin.setRange(0, 1000)
        btn_width_spin.setValue(self.control_btn_width)
        btn_width_spin.setSpecialValueText("自动")
        btn_width_spin.setSuffix(" px")
        form.addRow("抽选按钮宽度:", btn_width_spin)

        self.auto_checkbox = QCheckBox("启用自动抽选")
        self.auto_checkbox.setChecked(self.auto_mode)
        form.addRow("自动模式:", self.auto_checkbox)

        # 朗读功能复选框
        self.read_aloud_checkbox = QCheckBox("抽取后朗读名字")
        self.read_aloud_checkbox.setChecked(self.read_aloud_enabled)
        if not self.speech_available:
            self.read_aloud_checkbox.setEnabled(False)
            self.read_aloud_checkbox.setToolTip("系统不支持语音功能")
        form.addRow("朗读:", self.read_aloud_checkbox)

        import_btn = QPushButton("📂 选择文件导入")
        import_btn.clicked.connect(self.import_name_list)
        form.addRow("外部导入名单:", import_btn)

        # 手动添加按钮
        manual_btn = QPushButton("✏️ 手动添加名字")
        manual_btn.clicked.connect(self.add_names_manually)
        form.addRow("现场添加:", manual_btn)

        main_layout.addLayout(form)

        btn_layout = QHBoxLayout()
        save_btn = QPushButton("保存并应用")
        cancel_btn = QPushButton("取消")
        save_btn.clicked.connect(lambda: self._save_config(
            dialog, interval_spin, cooldown_spin, draw_count_spin,
            title_edit, opacity_spin,
            label_width_spin, label_height_spin, font_size_spin, btn_width_spin
        ))
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        main_layout.addLayout(btn_layout)

        dialog.exec()

    def _save_config(self, dialog, interval_spin, cooldown_spin, draw_count_spin,
                     title_edit, opacity_spin,
                     label_width_spin, label_height_spin, font_size_spin, btn_width_spin):
        if self.is_running:
            self.timer.stop()
            self.is_running = False
            self.stop_speaking()
            if self.auto_mode:
                self.auto_stop_timer.stop()
                self.control_btn.setEnabled(True)

        self.refresh_interval = interval_spin.value()
        self.cooldown_times = cooldown_spin.value()
        self.draw_count = draw_count_spin.value()
        self.window_title = title_edit.text().strip() or "综合高中252班随机点名程序"
        self.float_opacity = opacity_spin.value()
        self.name_label_width = label_width_spin.value()
        self.name_label_height = label_height_spin.value()
        self.name_font_size = font_size_spin.value()
        self.control_btn_width = btn_width_spin.value()
        self.auto_mode = self.auto_checkbox.isChecked()
        self.read_aloud_enabled = self.read_aloud_checkbox.isChecked()

        self.apply_name_label_style()
        self.apply_control_btn_width()

        if self.auto_mode:
            self.control_btn.setText("自动抽选")
        else:
            self.control_btn.setText("点击抽选")

        self.setWindowTitle(self.window_title)
        self.reposition_controls()
        self.save_data()
        dialog.accept()

    # ---------- 统计面板 ----------
    def show_statistics(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("抽取记录统计")
        dialog.setFixedSize(450, 580)

        layout = QVBoxLayout(dialog)
        total = sum(self.draw_counts.values())
        title = QLabel(f"<h2 style='color:#1e3a8a'>总抽取次数: {total}</h2>")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["姓名", "被抽次数"])
        table.setRowCount(len(self.draw_counts))

        sorted_counts = sorted(self.draw_counts.items(), key=lambda x: x[1], reverse=True)
        for row, (name, count) in enumerate(sorted_counts):
            table.setItem(row, 0, QTableWidgetItem(name))
            count_item = QTableWidgetItem(str(count))
            count_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            table.setItem(row, 1, count_item)

        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec()


class FloatWindow(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.drag_pos = QPoint()
        self.is_dragging = False
        self.init_float_ui()
        self.set_float_style()

    def init_float_ui(self):
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(85, 30)
        self.setWindowOpacity(self.parent.float_opacity)

        self.float_label = QLabel("点击点名", self)
        float_font = QFont("黑体", 12)
        float_font.setWeight(QFont.Weight.Normal)
        self.float_label.setFont(float_font)
        self.float_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.float_label.setFixedSize(85, 30)
        self.float_label.setCursor(Qt.CursorShape.PointingHandCursor)

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 3)
        self.setGraphicsEffect(shadow)

        screen_geo = self.screen().availableGeometry()
        x = screen_geo.width() - self.width() - 15
        y = screen_geo.height() - self.height() - 60
        self.move(x, y)

    def set_float_style(self):
        self.float_label.setStyleSheet("""
            QLabel {
                background-color: rgba(44, 62, 80, 90%);
                color: white;
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 20%);
            }
            QLabel:hover {
                background-color: rgba(64, 158, 255, 95%);
                border: 1px solid rgba(255, 255, 255, 30%);
            }
        """)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and not self.is_dragging:
            self.parent.restore_from_float()
        self.is_dragging = False
        self.setWindowOpacity(1.0)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.RightButton:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            self.is_dragging = False

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.RightButton and not self.drag_pos.isNull():
            self.is_dragging = True
            self.setWindowOpacity(0.8)
            new_pos = event.globalPosition().toPoint() - self.drag_pos
            self.move(new_pos)

    def closeEvent(self, event):
        self.parent.restore_from_float()
        event.accept()


def show_already_running_dialog():
    dialog = QDialog()
    dialog.setWindowTitle("点名工具")
    dialog.setFixedSize(380, 180)

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(30, 20, 30, 20)

    msg = QLabel("点名工具已开启，请留意桌面右下角。")
    msg.setFont(QFont("Microsoft YaHei", 11))
    msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
    msg.setWordWrap(True)
    layout.addWidget(msg)

    layout.addSpacing(20)

    btn_layout = QHBoxLayout()
    btn_layout.setSpacing(20)

    confirm_btn = QPushButton("确定")
    confirm_btn.setFixedSize(100, 35)
    confirm_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    confirm_btn.clicked.connect(dialog.accept)

    open_btn = QPushButton("打开程序")
    open_btn.setFixedSize(100, 35)
    open_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    open_btn.setStyleSheet("""
        QPushButton {
            background-color: #4a9eff;
            color: white;
            border: none;
            border-radius: 8px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #2f8cff;
        }
    """)

    def open_existing():
        socket = QLocalSocket()
        socket.connectToServer(SERVER_NAME)
        if socket.waitForConnected(1000):
            socket.write(b"show")
            socket.flush()
            socket.disconnectFromServer()
        dialog.accept()

    open_btn.clicked.connect(open_existing)

    btn_layout.addStretch()
    btn_layout.addWidget(confirm_btn)
    btn_layout.addWidget(open_btn)
    btn_layout.addStretch()
    layout.addLayout(btn_layout)

    dialog.exec()


def is_already_running():
    socket = QLocalSocket()
    socket.connectToServer(SERVER_NAME)
    if socket.waitForConnected(500):
        socket.disconnectFromServer()
        return True
    return False


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei UI"))

    if is_already_running():
        show_already_running_dialog()
        sys.exit(0)

    window = RollCallApp()
    window.show()
    sys.exit(app.exec())
