#!/usr/bin/env python3
"""
Serial Terminal - Standalone Application
A PyQt6 serial terminal application with split pane support and advanced formatting.
"""

import sys
import os
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QSplashScreen
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QFont, QColor
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtSvg import QSvgRenderer

from ui.dialogs.terminal_dialog import SerialMonitorWindow
from ui.resources import resource_manager
from constants import AppInfo


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = Path(__file__).parent

    return os.path.join(base_path, relative_path)


def create_terminal_icon():
    """Create the terminal application icon from app.svg file"""
    # Try to load from assets/icons/app.svg
    icon_path = get_resource_path(os.path.join('assets', 'icons', 'app.svg'))

    if os.path.exists(icon_path):
        # Load SVG from file
        renderer = QSvgRenderer(icon_path)

        pixmap = QPixmap(256, 256)
        pixmap.fill(Qt.GlobalColor.transparent)

        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()

        return QIcon(pixmap)
    else:
        # Fallback to empty icon if file not found
        print(f"Warning: Icon file not found at {icon_path}")
        return QIcon()


class TerminalSplashScreen(QSplashScreen):
    """Simple splash screen for Serial Terminal"""

    def __init__(self, pixmap):
        super().__init__(pixmap)

        self.setWindowFlags(
            Qt.WindowType.SplashScreen |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint
        )

        self.status_text = "Loading components..."
        self.progress = 0
        self.version_text = f"Version {AppInfo.VERSION}"

    def set_progress(self, value):
        """Set progress value (0-100)"""
        self.progress = max(0, min(100, value))
        self.update()

    def update_status(self, status):
        """Update status text"""
        self.status_text = status
        self.update()

    def paintEvent(self, event):
        """Paint splash screen using palette colors like VirtualPortManager"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        rect = self.rect()
        width = rect.width()
        height = rect.height()

        # Get colors from palette
        from PyQt6.QtGui import QPalette
        palette = self.palette()
        bg_color = palette.color(QPalette.ColorRole.Window)
        text_color = palette.color(QPalette.ColorRole.WindowText)
        disabled_color = palette.color(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText)
        highlight_color = palette.color(QPalette.ColorRole.Highlight)
        base_color = palette.color(QPalette.ColorRole.Base)

        # Draw background
        painter.fillRect(rect, bg_color)

        # Application name
        name_font = resource_manager.get_app_font(size=12, weight=QFont.Weight.Medium)
        painter.setFont(name_font)
        painter.setPen(text_color)
        painter.drawText(20, 32, AppInfo.NAME)

        # Version text
        version_font = resource_manager.get_app_font(size=9)
        painter.setFont(version_font)
        painter.setPen(disabled_color)
        painter.drawText(20, 48, self.version_text)

        # Status text
        painter.setPen(highlight_color)
        painter.drawText(20, height - 45, self.status_text)

        # Progress bar
        bar_width = width - 40
        bar_height = 2
        bar_x = 20
        bar_y = height - 25

        painter.fillRect(bar_x, bar_y, bar_width, bar_height, base_color)
        progress_width = int(bar_width * self.progress / 100)
        painter.fillRect(bar_x, bar_y, progress_width, bar_height, highlight_color)

        painter.end()


def create_splash_screen():
    """Create and return splash screen"""
    splash_pixmap = QPixmap(320, 180)
    splash_pixmap.fill(Qt.GlobalColor.transparent)
    return TerminalSplashScreen(splash_pixmap)


def main():
    """Main entry point for the Serial Terminal application"""
    # Disable platform theme detection
    os.environ['QT_QPA_PLATFORMTHEME'] = ''

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    # Set application properties
    app.setApplicationName(AppInfo.NAME)
    app.setApplicationVersion(AppInfo.VERSION)
    app.setOrganizationName(AppInfo.ORG_NAME)
    app.setOrganizationDomain(AppInfo.ORG_DOMAIN)

    # Use Fusion style (like VirtualPortManager)
    app.setStyle('fusion')

    # Create and show splash screen
    splash = create_splash_screen()
    splash.show()
    splash.raise_()
    splash.activateWindow()

    # Center splash on screen
    screen = splash.screen()
    if screen:
        screen_geometry = screen.availableGeometry()
        splash_geometry = splash.geometry()
        center_x = (screen_geometry.width() - splash_geometry.width()) // 2
        center_y = (screen_geometry.height() - splash_geometry.height()) // 2
        splash.move(center_x, center_y)

    app.processEvents()

    # Update splash
    splash.update_status("Initializing...")
    splash.set_progress(20)
    app.processEvents()

    # Load custom fonts
    splash.update_status("Loading fonts...")
    splash.set_progress(40)
    app.processEvents()

    resource_manager.load_custom_fonts("Poppins")
    resource_manager.load_custom_fonts("JetBrainsMono")

    # Set application default font
    app.setFont(resource_manager.get_app_font())

    # Create terminal icon
    splash.update_status("Loading icons...")
    splash.set_progress(60)
    app.processEvents()
    terminal_icon = create_terminal_icon()

    # Update splash
    splash.update_status("Loading interface...")
    splash.set_progress(80)
    app.processEvents()

    # Create terminal window
    terminal_window = SerialMonitorWindow()
    terminal_window.setWindowIcon(terminal_icon)
    terminal_window.setWindowTitle(AppInfo.NAME)

    # Final splash update
    splash.update_status("Ready...")
    splash.set_progress(100)
    app.processEvents()

    # Show window after splash delay
    QTimer.singleShot(1500, lambda: (
        splash.finish(terminal_window),
        terminal_window.show()
    ))

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
