import sys
import os
import signal
from PyQt5.QtWidgets import QApplication, QWidget, QHBoxLayout, QPushButton, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtGui import QCursor, QIcon, QPixmap, QPainter, QColor, QImage
import subprocess
import win32gui
import win32ui
import win32con
import win32api
import pywintypes
from PIL import Image
import io
import os
import json

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the path to the configuration file
config_path = os.path.join(script_dir, 'config.json')

# Load configuration from JSON
with open(config_path, 'r') as config_file:
    config = json.load(config_file)

DOCK_SIZE = config['DOCK_SIZE']
DOCK_SPACING = config['DOCK_SPACING']
DOCK_BACKGROUND_COLOR = tuple(config['DOCK_BACKGROUND_COLOR'])  # Ensure it's a tuple for QColor usage

class DockButton(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QPushButton {
                border: none;
                background-color: transparent;
                padding: 0px;
                margin: 0px;
            }
        """)
    
    def event(self, event):
        if event.type() == event.Enter:
            self.parent().setCursor(Qt.PointingHandCursor)
        elif event.type() == event.Leave:
            self.parent().setCursor(Qt.ArrowCursor)
        return super().event(event)

class Dock(QWidget):
    def __init__(self, screen):
        super().__init__()
        self.screen = screen
        self.initUI()
        
    def initUI(self):
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(DOCK_SPACING)
        
        # Use shortcuts from the configuration file
        shortcuts = config['shortcuts']
        
        for shortcut in shortcuts:
            if shortcut.get('name') == 'Spacer':
                layout.addSpacerItem(QSpacerItem(DOCK_SIZE, DOCK_SIZE, QSizePolicy.Fixed, QSizePolicy.Fixed))
            else:
                btn = self.create_button(shortcut['name'], shortcut.get('path'), shortcut.get('icon_path'))
                layout.addWidget(btn)
        
        layout.insertStretch(0, 1)
        layout.addStretch(1)
        
        self.setLayout(layout)
        screen_geometry = self.screen.geometry()
        self.setGeometry(screen_geometry.left(), screen_geometry.top(), screen_geometry.width(), DOCK_SIZE)
        self.hide()

    def create_button(self, name, path, icon_path):
        btn = DockButton(self)
        btn.clicked.connect(lambda _, p=path: self.launch_app(p))
        
        icon = self.load_icon(icon_path, path, name)
        btn.setIcon(icon)
        btn.setIconSize(QSize(DOCK_SIZE, DOCK_SIZE))
        btn.setFixedSize(DOCK_SIZE, DOCK_SIZE)
        
        return btn

    def load_icon(self, icon_path, exe_path, name):
        print(f"Loading icon for {name}...")
        
        # Try to load the provided PNG
        if icon_path and os.path.exists(icon_path):
            print(f"Attempting to load PNG icon from {icon_path}")
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                print("Successfully loaded PNG icon")
                return QIcon(pixmap)
            else:
                print("Failed to load PNG icon")
        
        # If PNG not available, try to extract icon from the executable
        if exe_path and os.path.exists(exe_path):
            print(f"Attempting to extract icon from executable: {exe_path}")
            icon = self.extract_icon_from_exe(exe_path)
            if icon:
                print("Successfully extracted icon from executable")
                return icon
            else:
                print("Failed to extract icon from executable")
        
        # If both methods fail, create a placeholder
        print("Creating placeholder icon")
        return self.create_placeholder_icon(name)

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

                # Get icon info
                hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                hbmp = win32ui.CreateBitmap()
                hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
                hdc = hdc.CreateCompatibleDC()

                hdc.SelectObject(hbmp)
                hdc.DrawIcon((0, 0), large[0])

                # Convert to PIL Image
                bmpstr = hbmp.GetBitmapBits(True)
                pil_img = Image.frombytes('RGBA', (ico_x, ico_y), bmpstr, 'raw', 'BGRA')

                # Convert PIL Image to QPixmap
                buffer = io.BytesIO()
                pil_img.save(buffer, format='PNG')
                pixmap = QPixmap()
                pixmap.loadFromData(buffer.getvalue())

                print(f"Extracted icon size: {pixmap.width()}x{pixmap.height()}")

                # Scale if necessary
                if pixmap.width() != DOCK_SIZE or pixmap.height() != DOCK_SIZE:
                    print(f"Scaling icon from {pixmap.width()}x{pixmap.height()} to {DOCK_SIZE}x{DOCK_SIZE}")
                    scaled_pixmap = pixmap.scaled(DOCK_SIZE, DOCK_SIZE, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                else:
                    print(f"Icon size matches DOCK_SIZE ({DOCK_SIZE}x{DOCK_SIZE}). Skipping scaling.")
                    scaled_pixmap = pixmap

                # Clean up
                win32gui.DestroyIcon(large[0])

                print(f"Icon processed to {DOCK_SIZE}x{DOCK_SIZE}")
                return QIcon(scaled_pixmap)
            else:
                print("Failed to extract large icon")
        except pywintypes.error as e:
            print(f"Failed to extract icon from {exe_path}: {e}")
        return None

    def create_placeholder_icon(self, name):
        print(f"Creating placeholder icon for {name}")
        pixmap = QPixmap(DOCK_SIZE, DOCK_SIZE)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(200, 200, 200))  # Light gray color
        painter.drawEllipse(0, 0, DOCK_SIZE, DOCK_SIZE)
        painter.setPen(QColor(100, 100, 100))  # Dark gray color for text
        painter.drawText(pixmap.rect(), Qt.AlignCenter, name[0].upper())
        painter.end()
        return QIcon(pixmap)
        
    def launch_app(self, path):
        if path and os.path.exists(path):
            try:
                subprocess.Popen(path)
            except subprocess.CalledProcessError as e:
                print(f"Error launching {path}: {e}")
        else:
            print(f"Path does not exist: {path}")

    def mousePressEvent(self, event):
        # This ensures clicks are registered even at the top and bottom pixels
        for child in self.children():
            if isinstance(child, DockButton):
                if child.geometry().contains(event.pos()):
                    child.click()
                    return
        super().mousePressEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(*DOCK_BACKGROUND_COLOR))  # Use the global background color
        painter.drawRect(self.rect())
        super().paintEvent(event)

class DockManager:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.docks = []
        self.create_docks()
        
        self.check_mouse_timer = QTimer()
        self.check_mouse_timer.timeout.connect(self.check_mouse_position)
        self.check_mouse_timer.start(50)  # Check every 50 ms

        # Connect signals for screen changes
        self.app.screenAdded.connect(self.screen_added)
        self.app.screenRemoved.connect(self.screen_removed)
        
        # Set up signal handling for graceful termination
        signal.signal(signal.SIGINT, self.signal_handler)

    def create_docks(self):
        # Clear existing docks to avoid duplications
        self.docks.clear()
        
        # Create docks for each screen
        for screen in self.app.screens():
            dock = Dock(screen)
            self.docks.append(dock)

    def screen_added(self, screen):
        print(f"Screen added: {screen.name()}")
        dock = Dock(screen)
        self.docks.append(dock)
    
    def screen_removed(self, screen):
        print(f"Screen removed: {screen.name()}")
        # Remove docks associated with the removed screen
        self.docks = [dock for dock in self.docks if dock.screen != screen]

    def check_mouse_position(self):
        cursor = QCursor.pos()
        for dock in self.docks:
            if dock.screen is None or dock.screen not in self.app.screens():
                print(f"Skipping dock on invalid screen: {dock.screen}")
                continue

            screen_geometry = dock.screen.geometry()
            dock_width = screen_geometry.width()
            center_width = dock_width * 0.25
            center_left = screen_geometry.left() + (dock_width - center_width) / 2
            center_right = center_left + center_width

            if (cursor.y() == screen_geometry.top() and 
                center_left <= cursor.x() <= center_right):
                dock.show()
            elif not dock.geometry().contains(cursor):
                dock.hide()

    def signal_handler(self, sig, frame):
        print("\nCtrl+C pressed. Shutting down gracefully...")
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec_())

if __name__ == '__main__':
    manager = DockManager()
    manager.run()