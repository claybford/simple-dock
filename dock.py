import sys
import os
import signal
from PyQt5.QtWidgets import QApplication, QWidget, QHBoxLayout, QPushButton, QSpacerItem, QSizePolicy, QLabel
from PyQt5.QtCore import Qt, QTimer, QSize, QEvent
from PyQt5.QtGui import QCursor, QIcon, QPixmap, QPainter, QColor, QFont, QFontMetrics
import win32gui
import win32ui
import win32con
import win32api
import pywintypes
from PIL import Image
import io
import json

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the path to the configuration file
config_path = os.path.join(script_dir, 'config.json')

# Load configuration from JSON
with open(config_path, 'r') as config_file:
    config = json.load(config_file)

ICON_SIZE = config['ICON_SIZE']
DOCK_BACKGROUND_COLOR = tuple(config['DOCK_BACKGROUND_COLOR'])
HOVER_BACKGROUND_COLOR = tuple(config['HOVER_BACKGROUND_COLOR'])
ICON_PADDING = config['ICON_PADDING']
HOVER_TEXT_COLOR = tuple(config['HOVER_TEXT_COLOR'])
HOVER_TEXT_FONT = config['HOVER_TEXT_FONT']
HOVER_TEXT_SIZE = config['HOVER_TEXT_SIZE']
ACTIVATION_PERCENTAGE = config['ACTIVATION_PERCENTAGE']

TOTAL_BUTTON_SIZE = ICON_SIZE + (2 * ICON_PADDING)

print(f"Configuration loaded: ICON_SIZE={ICON_SIZE}, ICON_PADDING={ICON_PADDING}")
print(f"HOVER_TEXT_COLOR={HOVER_TEXT_COLOR}, HOVER_TEXT_FONT={HOVER_TEXT_FONT}, HOVER_TEXT_SIZE={HOVER_TEXT_SIZE}")

class HoverLabel(QLabel):
    def __init__(self, screen, parent=None):
        super().__init__(parent)
        self.screen = screen
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Format color values
        text_color = f"rgba({HOVER_TEXT_COLOR[0]}, {HOVER_TEXT_COLOR[1]}, {HOVER_TEXT_COLOR[2]}, {HOVER_TEXT_COLOR[3] / 255.0})"
        
        # Escape font name if it contains spaces
        font_name = f"'{HOVER_TEXT_FONT}'" if ' ' in HOVER_TEXT_FONT else HOVER_TEXT_FONT
        
        self.setStyleSheet(f"""
            QLabel {{
                color: {text_color};
                font-family: {font_name};
                font-size: {HOVER_TEXT_SIZE}px;
            }}
        """)
        self.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        self.label_height = self.calculate_height()
        geometry = self.calculate_position(self.screen)
        self.setGeometry(*geometry)
    
    def calculate_height(self):
        font = QFont(HOVER_TEXT_FONT, HOVER_TEXT_SIZE)
        fm = QFontMetrics(font)
        typical_text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        return fm.boundingRect(typical_text).height()

    def calculate_position(self, screen):
        screen_geometry = screen.geometry()
        global_geometry = screen.availableGeometry()
        x = global_geometry.left()
        y = global_geometry.top() + TOTAL_BUTTON_SIZE + ICON_PADDING
        return x, y, screen_geometry.width(), self.label_height
    
    def setText(self, text):
        super().setText(text)

class DockButton(QPushButton):
    def __init__(self, name, parent=None):
        super().__init__(parent)
        self.name = name
        button_color = (
            f"rgba({HOVER_BACKGROUND_COLOR[0]}, {HOVER_BACKGROUND_COLOR[1]}, "
            f"{HOVER_BACKGROUND_COLOR[2]}, {HOVER_BACKGROUND_COLOR[3] / 255.0})"
        )
       
        self.setStyleSheet(f"""
            QPushButton {{
                border: none;
                background-color: transparent;
                padding: {ICON_PADDING}px;
                margin: 0px;
                outline: none;
            }}
            QPushButton:focus {{
                outline: none;
            }}
            QPushButton:hover {{
                background-color: {button_color};
            }}
            QPushButton:pressed {{
                background-color: {button_color};
            }}
            QPushButton:!hover {{
                background-color: transparent;
            }}
        """)
        print(f"DockButton created: {name}")
   
    def event(self, event):
        if event.type() == QEvent.Enter:
            self.parent().setCursor(Qt.PointingHandCursor)
            self.parent().show_hover_label(self.name)
        elif event.type() == QEvent.Leave:
            self.parent().setCursor(Qt.ArrowCursor)
            self.parent().hide_hover_label()
        return super().event(event)

class Dock(QWidget):
    def __init__(self, screen):
        super().__init__()
        self.screen = screen
        self.hover_label = HoverLabel(self.screen)
        self.initUI()
        
    def initUI(self):
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        shortcuts = config['shortcuts']
        
        for shortcut in shortcuts:
            if shortcut.get('name') == 'Spacer':
                layout.addSpacerItem(QSpacerItem(TOTAL_BUTTON_SIZE, TOTAL_BUTTON_SIZE, QSizePolicy.Fixed, QSizePolicy.Fixed))
            else:
                btn = self.create_button(shortcut['name'], shortcut.get('path'), shortcut.get('icon_path'))
                layout.addWidget(btn)
        
        layout.insertStretch(0, 1)
        layout.addStretch(1)
        
        self.setLayout(layout)
        screen_geometry = self.screen.geometry()
        self.setGeometry(screen_geometry.left(), screen_geometry.top(), screen_geometry.width(), TOTAL_BUTTON_SIZE)
        self.hide()

    def create_button(self, name, path, icon_path):
        btn = DockButton(name, self)
        btn.clicked.connect(lambda _, p=path: self.launch_app(p))
        
        icon = self.load_icon(icon_path, path, name)
        btn.setIcon(icon)
        btn.setIconSize(QSize(ICON_SIZE, ICON_SIZE))
        btn.setFixedSize(TOTAL_BUTTON_SIZE, TOTAL_BUTTON_SIZE)
        
        return btn

    def load_icon(self, icon_path, exe_path, name):
        print(f"Loading icon for {name}...")
        
        if icon_path and os.path.exists(icon_path):
            print(f"Attempting to load PNG icon from {icon_path}")
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                print("Successfully loaded PNG icon")
                return QIcon(self.scale_pixmap(pixmap))
        
        if exe_path and os.path.exists(exe_path):
            print(f"Attempting to extract icon from executable: {exe_path}")
            pixmap = self.extract_icon_from_exe(exe_path)
            if pixmap:
                print("Successfully extracted icon from executable")
                return QIcon(self.scale_pixmap(pixmap))
        
        print("Creating placeholder icon")
        return self.create_placeholder_icon(name)

    def scale_pixmap(self, pixmap):
        if pixmap.width() != ICON_SIZE or pixmap.height() != ICON_SIZE:
            print(f"Scaling icon from {pixmap.width()}x{pixmap.height()} to {ICON_SIZE}x{ICON_SIZE}")
            return pixmap.scaled(ICON_SIZE, ICON_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        print(f"Icon size matches ICON_SIZE ({ICON_SIZE}x{ICON_SIZE}). Skipping scaling.")
        return pixmap

    def extract_icon_from_exe(self, exe_path):
        try:
            print("Starting icon extraction process...")
            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
            print(f"System icon size: {ico_x}x{ico_y}")

            large, small = win32gui.ExtractIconEx(exe_path, 0, 1)
            if large:
                print("Successfully extracted large icon")
                win32gui.DestroyIcon(small[0])

                hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                hbmp = win32ui.CreateBitmap()
                hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
                hdc = hdc.CreateCompatibleDC()

                hdc.SelectObject(hbmp)
                hdc.DrawIcon((0, 0), large[0])

                bmpstr = hbmp.GetBitmapBits(True)
                pil_img = Image.frombytes('RGBA', (ico_x, ico_y), bmpstr, 'raw', 'BGRA')

                buffer = io.BytesIO()
                pil_img.save(buffer, format='PNG')
                pixmap = QPixmap()
                pixmap.loadFromData(buffer.getvalue())

                win32gui.DestroyIcon(large[0])

                print(f"Extracted icon size: {pixmap.width()}x{pixmap.height()}")
                return pixmap
            else:
                print("Failed to extract large icon")
        except pywintypes.error as e:
            print(f"Failed to extract icon from {exe_path}: {e}")
        return None
    
    def create_placeholder_icon(self, name):
        print(f"Creating placeholder icon for {name}")
        pixmap = QPixmap(ICON_SIZE, ICON_SIZE)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(200, 200, 200))
        painter.drawEllipse(0, 0, ICON_SIZE, ICON_SIZE)
        painter.setPen(QColor(100, 100, 100))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, name[0].upper())
        painter.end()
        return QIcon(pixmap)
        
    def launch_app(self, path):
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                print(f"Error launching {path}: {e}")
        else:
            print(f"Path does not exist: {path}")
    
    def show_hover_label(self, text):
        self.hover_label.setText(text)
        self.hover_label.show()

    def hide_hover_label(self):
        self.hover_label.hide()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(*DOCK_BACKGROUND_COLOR))
        painter.drawRect(self.rect())
        super().paintEvent(event)

class DockManager:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.docks = []
        self.create_docks()
        
        self.check_mouse_timer = QTimer()
        self.check_mouse_timer.timeout.connect(self.check_mouse_position)
        self.check_mouse_timer.start(50)

        self.app.screenAdded.connect(self.screen_added)
        self.app.screenRemoved.connect(self.screen_removed)
        
        signal.signal(signal.SIGINT, self.signal_handler)

    def create_docks(self):
        self.docks.clear()
        for screen in self.app.screens():
            dock = Dock(screen)
            self.docks.append(dock)
        print(f"Created {len(self.docks)} docks")

    def screen_added(self, screen):
        print(f"Screen added: {screen.name()}")
        dock = Dock(screen)
        self.docks.append(dock)
    
    def screen_removed(self, screen):
        print(f"Screen removed: {screen.name()}")
        self.docks = [dock for dock in self.docks if dock.screen != screen]

    def check_mouse_position(self):
        cursor = QCursor.pos()
        for dock in self.docks:
            if dock.screen is None or dock.screen not in self.app.screens():
                print(f"Skipping dock on invalid screen: {dock.screen}")
                continue

            screen_geometry = dock.screen.geometry()
            dock_width = screen_geometry.width()
            center_width = dock_width * ACTIVATION_PERCENTAGE
            center_left = screen_geometry.left() + (dock_width - center_width) / 2
            center_right = center_left + center_width

            if (cursor.y() == screen_geometry.top() and 
                center_left <= cursor.x() <= center_right):
                if not dock.isVisible():
                    dock.show()
            elif not dock.geometry().contains(cursor):
                if dock.isVisible():
                    dock.hide()
                    dock.hide_hover_label()

    def signal_handler(self, sig, frame):
        print("\nCtrl+C pressed. Shutting down gracefully...")
        self.app.quit()

    def run(self):
        print("Starting DockManager...")
        sys.exit(self.app.exec_())

if __name__ == '__main__':
    manager = DockManager()
    manager.run()