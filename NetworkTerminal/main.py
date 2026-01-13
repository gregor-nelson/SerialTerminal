#!/usr/bin/env python3
"""
Network Terminal - Entry Point
TCP/UDP Terminal Application
"""

import sys
from PyQt6.QtWidgets import QApplication

from constants import AppInfo
from ui.resources import resource_manager
from ui.dialogs.terminal_dialog import NetworkMonitorWindow


def main():
    """Main entry point for Network Terminal"""
    app = QApplication(sys.argv)
    app.setApplicationName(AppInfo.NAME)
    app.setOrganizationName(AppInfo.ORG_NAME)

    # Set Fusion style like Serial Terminal
    app.setStyle("Fusion")

    # Load custom fonts
    resource_manager.load_custom_fonts("Poppins")
    resource_manager.load_custom_fonts("JetBrainsMono")

    # Create and show main window
    window = NetworkMonitorWindow()
    window.resize(1200, 800)
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
