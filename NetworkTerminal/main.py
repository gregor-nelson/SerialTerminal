#!/usr/bin/env python3
"""
Network Terminal - TCP/UDP Terminal Application
"""

import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPalette, QColor
from PyQt6.QtCore import Qt

from constants import AppInfo
from ui.resources import resource_manager


def apply_dark_palette(app: QApplication):
    """Apply dark color palette to application"""
    palette = QPalette()

    palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)

    app.setPalette(palette)


def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    app.setApplicationName(AppInfo.NAME)
    app.setOrganizationName(AppInfo.ORG_NAME)

    # Set Fusion style for consistent look
    app.setStyle("Fusion")

    # Apply dark palette
    apply_dark_palette(app)

    # Load custom fonts
    resource_manager.load_custom_fonts("Poppins")
    resource_manager.load_custom_fonts("JetBrainsMono")

    # Import main window (after app created for QSettings)
    from ui.dialogs.terminal_dialog import NetworkMonitorWindow

    # Create and show main window
    window = NetworkMonitorWindow()
    window.resize(1200, 800)
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
