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


# --- Terminal SVG Icon ---
TERMINAL_ICON_SVG = """
<svg width="256" height="256" viewBox="0 0 256 256" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <!-- Terminal-specific gradients -->
    <linearGradient id="terminalGradient" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" style="stop-color:#00ff88;stop-opacity:1" />
      <stop offset="50%" style="stop-color:#00cc66;stop-opacity:1" />
      <stop offset="100%" style="stop-color:#009944;stop-opacity:1" />
    </linearGradient>

    <linearGradient id="terminalHighlight" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" style="stop-color:#ffffff;stop-opacity:0.3" />
      <stop offset="100%" style="stop-color:#ffffff;stop-opacity:0" />
    </linearGradient>

    <!-- Terminal screen effect -->
    <filter id="screenGlow">
      <feGaussianBlur stdDeviation="3" result="coloredBlur"/>
      <feMerge>
        <feMergeNode in="coloredBlur"/>
        <feMergeNode in="SourceGraphic"/>
      </feMerge>
    </filter>

    <filter id="shadow">
      <feDropShadow dx="0" dy="3" stdDeviation="4" flood-opacity="0.3"/>
    </filter>
  </defs>

  <!-- Terminal screen representation -->
  <g transform="translate(128, 128)">

    <!-- Terminal frame -->
    <rect x="-100" y="-80" width="200" height="160" rx="15" fill="#1a1a1a" filter="url(#shadow)"/>
    <rect x="-95" y="-75" width="190" height="150" rx="10" fill="#000000"/>

    <!-- Terminal screen -->
    <rect x="-85" y="-60" width="170" height="100" rx="5" fill="#0d1117" filter="url(#screenGlow)"/>

    <!-- Terminal text lines -->
    <g fill="url(#terminalGradient)" font-family="monospace" font-size="10">
      <text x="-75" y="-40">$ serial-terminal</text>
      <text x="-75" y="-25">Connected: COM3</text>
      <text x="-75" y="-10">Baud: 115200</text>
      <text x="-75" y="5">Data: 8N1</text>
      <text x="-75" y="20">Status: Ready</text>
    </g>

    <!-- Cursor blink -->
    <rect x="-20" y="15" width="8" height="12" fill="url(#terminalGradient)">
      <animate attributeName="opacity" values="1;0;1" dur="1s" repeatCount="indefinite"/>
    </rect>

    <!-- Control buttons -->
    <g>
      <!-- Power button -->
      <circle cx="70" cy="-55" r="8" fill="url(#terminalGradient)" opacity="0.8"/>
      <circle cx="70" cy="-55" r="4" fill="#ffffff" opacity="0.9"/>

      <!-- Activity indicators -->
      <circle cx="55" cy="-55" r="4" fill="#ff4444">
        <animate attributeName="opacity" values="0.3;1;0.3" dur="2s" repeatCount="indefinite"/>
      </circle>
      <circle cx="85" cy="-55" r="4" fill="#44ff44">
        <animate attributeName="opacity" values="1;0.3;1" dur="2s" repeatCount="indefinite"/>
      </circle>
    </g>

    <!-- Bottom edge highlight -->
    <rect x="-100" y="60" width="200" height="20" rx="15" fill="url(#terminalHighlight)"/>

    <!-- Connection ports -->
    <g transform="translate(0, 85)">
      <rect x="-15" y="-5" width="30" height="10" rx="5" fill="url(#terminalGradient)"/>
      <rect x="-12" y="-3" width="6" height="6" rx="2" fill="#ffffff" opacity="0.9"/>
      <rect x="-2" y="-3" width="6" height="6" rx="2" fill="#ffffff" opacity="0.9"/>
      <rect x="8" y="-3" width="6" height="6" rx="2" fill="#ffffff" opacity="0.9"/>
    </g>
  </g>
</svg>
"""


def create_terminal_icon():
    """Create the terminal application icon from SVG data"""
    svg_bytes = TERMINAL_ICON_SVG.encode('utf-8')
    renderer = QSvgRenderer(svg_bytes)

    pixmap = QPixmap(256, 256)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()

    return QIcon(pixmap)


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
        from ui.resources import resource_manager
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
