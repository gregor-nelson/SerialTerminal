"""Ribbon-style toolbar for Serial Terminal commands."""

from PyQt6.QtWidgets import QToolBar, QWidget, QHBoxLayout, QPushButton
from PyQt6.QtCore import pyqtSignal, QSize
from PyQt6.QtGui import QIcon

from ..resources import resource_manager


class RibbonButton(QPushButton):
    """Large ribbon-style button with icon and text."""

    def __init__(self, text: str, icon_name: str = None, parent=None):
        super().__init__(parent)
        self.setText(text)

        # Set icon if provided using resource manager
        if icon_name:
            icon = resource_manager.get_toolbar_icon(icon_name)
            if not icon.isNull():
                self.setIcon(icon)
                self.setIconSize(QSize(16, 16))

        # Use default button styling

    def update_icon(self, icon_name: str):
        """Update button icon dynamically."""
        icon = resource_manager.get_toolbar_icon(icon_name)
        if not icon.isNull():
            self.setIcon(icon)


class RibbonToolbar(QToolBar):
    """Ribbon-style toolbar with Serial Terminal commands."""

    # Signals for terminal actions (simplified to 5 primary actions)
    new_connection = pyqtSignal()
    refresh_ports = pyqtSignal()
    toggle_connection = pyqtSignal()
    clear_terminal = pyqtSignal()
    show_settings = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.setup_actions()

    def setup_ui(self):
        """Set up the ribbon toolbar UI."""
        self.setMovable(False)
        self.setFloatable(False)
        self.setMinimumHeight(48)
        self.setMaximumHeight(48)

        # Main widget to hold ribbon buttons (flat layout, no groups)
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setSpacing(4)  # 4px spacing between buttons like SerialRouter
        main_layout.setContentsMargins(5, 5, 5, 5)

        # Create 5 buttons in flat layout
        self.new_button = RibbonButton("New", "new")
        self.new_button.setToolTip("New connection (Ctrl+N)")

        self.refresh_button = RibbonButton("Refresh", "refresh")
        self.refresh_button.setToolTip("Refresh available ports")

        self.connect_button = RibbonButton("Connect", "enable")
        self.connect_button.setToolTip("Connect to serial port")

        self.clear_button = RibbonButton("Clear", "remove")
        self.clear_button.setToolTip("Clear terminal output")

        self.settings_button = RibbonButton("Settings", "configure")
        self.settings_button.setToolTip("Application settings")

        # Add buttons to layout
        main_layout.addWidget(self.new_button)
        main_layout.addWidget(self.refresh_button)
        main_layout.addWidget(self.connect_button)
        main_layout.addWidget(self.clear_button)
        main_layout.addWidget(self.settings_button)
        main_layout.addStretch()

        # Add main widget to toolbar
        self.addWidget(main_widget)

    def setup_actions(self):
        """Set up button actions."""
        self.new_button.clicked.connect(self.new_connection.emit)
        self.refresh_button.clicked.connect(self.refresh_ports.emit)
        self.connect_button.clicked.connect(self.toggle_connection.emit)
        self.clear_button.clicked.connect(self.clear_terminal.emit)
        self.settings_button.clicked.connect(self.show_settings.emit)

    def set_connection_state(self, is_connected: bool):
        """Update connect/disconnect button based on connection state."""
        if is_connected:
            self.connect_button.setText("Disconnect")
            self.connect_button.setToolTip("Disconnect from serial port")
            self.connect_button.update_icon("disable")
        else:
            self.connect_button.setText("Connect")
            self.connect_button.setToolTip("Connect to serial port")
            self.connect_button.update_icon("enable")

    def set_pane_actions_enabled(self, enabled: bool):
        """Enable/disable pane-specific actions based on context."""
        self.connect_button.setEnabled(enabled)
        self.clear_button.setEnabled(enabled)
