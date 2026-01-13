#!/usr/bin/env python3
"""
Network Terminal - Entry Point
Minimal version for Phase 0 verification
"""

import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel
from PyQt6.QtCore import Qt

# Verify all copied modules import correctly
from constants import AppInfo, TerminalColors
from ui.resources import resource_manager
from ui.windows.terminal_formatter import TerminalStreamFormatter
from ui.components.ribbon_toolbar import RibbonToolbar
from ui.common.icons import Icons


def main():
    """Main entry point for verification"""
    app = QApplication(sys.argv)
    app.setApplicationName(AppInfo.NAME)
    app.setOrganizationName(AppInfo.ORG_NAME)

    # Set Fusion style like Serial Terminal
    app.setStyle("Fusion")

    # Load custom fonts
    resource_manager.load_custom_fonts("Poppins")
    resource_manager.load_custom_fonts("JetBrainsMono")

    # Create a simple test window
    window = QMainWindow()
    window.setWindowTitle(f"{AppInfo.NAME} - Phase 0 Verification")
    window.setMinimumSize(800, 600)

    # Add a label to confirm it works
    label = QLabel("Phase 0 Complete!\n\nAll modules imported successfully.")
    label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    label.setFont(resource_manager.get_monospace_font(size=14))
    window.setCentralWidget(label)

    # Test that formatter initializes
    formatter = TerminalStreamFormatter()
    print(f"Formatter colors loaded: {len(formatter.colors)} types")
    print(f"NMEA colors loaded: {len(formatter.nmea_colors)} types")

    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
