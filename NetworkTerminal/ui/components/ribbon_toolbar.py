"""Ribbon-style toolbar for Network Terminal commands."""

from PyQt6.QtWidgets import QToolBar, QWidget, QHBoxLayout, QPushButton, QMenu
from PyQt6.QtCore import pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QFont

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

        # Set medium font weight for better readability
        font = self.font()
        font.setWeight(QFont.Weight.Medium)
        self.setFont(font)

    def update_icon(self, icon_name: str):
        """Update button icon dynamically."""
        icon = resource_manager.get_toolbar_icon(icon_name)
        if not icon.isNull():
            self.setIcon(icon)


class RibbonMenuButton(RibbonButton):
    """Ribbon button that shows a dropdown menu when clicked."""

    def __init__(self, text: str, icon_name: str = None, parent=None):
        super().__init__(text, icon_name, parent)
        self._menu = QMenu(self)
        self.clicked.connect(self._show_menu)

    def menu(self) -> QMenu:
        """Get the dropdown menu."""
        return self._menu

    def _show_menu(self):
        """Show the dropdown menu below the button."""
        pos = self.mapToGlobal(self.rect().bottomLeft())
        self._menu.exec(pos)


class RibbonToolbar(QToolBar):
    """Ribbon-style toolbar for Network Terminal commands."""

    # Signals for terminal actions
    new_connection = pyqtSignal()
    refresh_ports = pyqtSignal()  # Kept for API compatibility, but hidden
    toggle_connection = pyqtSignal()
    clear_terminal = pyqtSignal()
    show_settings = pyqtSignal()
    show_history = pyqtSignal()
    recent_selected = pyqtSignal(object)  # Emits NetworkConfig

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
        main_layout.setSpacing(4)  # 4px spacing between buttons
        main_layout.setContentsMargins(5, 5, 5, 5)

        # Create buttons
        self.new_button = RibbonButton("New", "new")
        self.new_button.setToolTip("New connection (Ctrl+N)")

        # Recent connections dropdown
        self.recent_button = RibbonMenuButton("Recent", "refresh")
        self.recent_button.setToolTip("Recent connections")

        # History button
        self.history_button = RibbonButton("History", "configure")
        self.history_button.setToolTip("Connection history and favorites")

        # Refresh button hidden for network (kept for API compatibility)
        self.refresh_button = RibbonButton("Refresh", "refresh")
        self.refresh_button.setToolTip("Refresh")
        self.refresh_button.setVisible(False)  # Hidden for network

        self.connect_button = RibbonButton("Connect", "enable")
        self.connect_button.setToolTip("Connect to network")

        self.clear_button = RibbonButton("Clear", "remove")
        self.clear_button.setToolTip("Clear terminal output")

        self.settings_button = RibbonButton("Settings", "configure")
        self.settings_button.setToolTip("Terminal settings")

        # Add buttons to layout
        main_layout.addWidget(self.new_button)
        main_layout.addWidget(self.recent_button)
        main_layout.addWidget(self.history_button)
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
        self.history_button.clicked.connect(self.show_history.emit)

    def populate_recent_menu(self, recent_configs):
        """Populate the recent connections dropdown menu."""
        menu = self.recent_button.menu()
        menu.clear()

        if not recent_configs:
            no_recent = menu.addAction("No recent connections")
            no_recent.setEnabled(False)
            return

        for config in recent_configs:
            action = menu.addAction(f"{config.protocol} {config.host}:{config.port}")
            action.triggered.connect(
                lambda checked, c=config: self.recent_selected.emit(c)
            )

    def set_connection_state(self, is_connected: bool):
        """Update connect/disconnect button based on connection state."""
        if is_connected:
            self.connect_button.setText("Disconnect")
            self.connect_button.setToolTip("Disconnect from network")
            self.connect_button.update_icon("disable")
        else:
            self.connect_button.setText("Connect")
            self.connect_button.setToolTip("Connect to network")
            self.connect_button.update_icon("enable")

    def set_pane_actions_enabled(self, enabled: bool):
        """Enable/disable pane-specific actions based on context."""
        self.connect_button.setEnabled(enabled)
        self.clear_button.setEnabled(enabled)
