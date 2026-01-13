#!/usr/bin/env python3
"""
Connection Dialog - Configure network connections

This module contains:
- QuickConnectDialog: Configure protocol, mode, host, port for new connections
- ConnectionHistoryDialog: Browse favorites and connection history
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QGroupBox, QRadioButton, QLineEdit, QSpinBox,
    QPushButton, QCheckBox, QDoubleSpinBox,
    QTabWidget, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt, pyqtSignal

from core.network_config import NetworkConfig


class QuickConnectDialog(QDialog):
    """
    Quick connect dialog for network connections.

    Allows user to configure:
    - Protocol (TCP/UDP)
    - Mode (Client/Server)
    - Host address
    - Port number
    - Advanced TCP/UDP options
    """

    def __init__(self, parent=None, initial_config: NetworkConfig = None):
        super().__init__(parent)
        self.setWindowTitle("New Connection")
        self.setModal(True)
        self.setMinimumWidth(400)

        # Default config
        self.config = initial_config or NetworkConfig(
            host="127.0.0.1",
            port=5000,
            protocol='TCP',
            mode='client'
        )

        self._setup_ui()
        self._connect_signals()
        self._update_ui_state()

    def _setup_ui(self):
        """Setup dialog UI"""
        layout = QVBoxLayout(self)
        layout.setSpacing(16)

        # === Protocol Selection ===
        protocol_group = QGroupBox("Protocol")
        protocol_layout = QHBoxLayout(protocol_group)

        self.tcp_radio = QRadioButton("TCP")
        self.udp_radio = QRadioButton("UDP")
        self.tcp_radio.setChecked(self.config.protocol == 'TCP')
        self.udp_radio.setChecked(self.config.protocol == 'UDP')

        protocol_layout.addWidget(self.tcp_radio)
        protocol_layout.addWidget(self.udp_radio)
        protocol_layout.addStretch()

        layout.addWidget(protocol_group)

        # === Mode Selection ===
        mode_group = QGroupBox("Mode")
        mode_layout = QHBoxLayout(mode_group)

        self.client_radio = QRadioButton("Client (Connect to)")
        self.server_radio = QRadioButton("Server (Listen on)")
        self.client_radio.setChecked(self.config.mode == 'client')
        self.server_radio.setChecked(self.config.mode == 'server')

        mode_layout.addWidget(self.client_radio)
        mode_layout.addWidget(self.server_radio)
        mode_layout.addStretch()

        layout.addWidget(mode_group)

        # === Connection Details ===
        details_group = QGroupBox("Connection Details")
        details_layout = QFormLayout(details_group)

        # Host input
        self.host_input = QLineEdit(self.config.host)
        self.host_input.setPlaceholderText("hostname or IP address")
        details_layout.addRow("Host:", self.host_input)

        # Port input
        self.port_input = QSpinBox()
        self.port_input.setRange(1, 65535)
        self.port_input.setValue(self.config.port)
        details_layout.addRow("Port:", self.port_input)

        layout.addWidget(details_group)

        # === Advanced Options (collapsible) ===
        self.advanced_group = QGroupBox("Advanced Options")
        self.advanced_group.setCheckable(True)
        self.advanced_group.setChecked(False)
        advanced_layout = QFormLayout(self.advanced_group)

        # TCP options
        self.keepalive_check = QCheckBox("Enable TCP Keepalive")
        self.keepalive_check.setChecked(self.config.keepalive)
        advanced_layout.addRow("", self.keepalive_check)

        self.nodelay_check = QCheckBox("Disable Nagle (TCP_NODELAY)")
        self.nodelay_check.setChecked(self.config.nodelay)
        advanced_layout.addRow("", self.nodelay_check)

        # UDP options
        self.broadcast_check = QCheckBox("Enable Broadcast")
        self.broadcast_check.setChecked(self.config.broadcast)
        advanced_layout.addRow("", self.broadcast_check)

        # Timeout
        self.timeout_spin = QDoubleSpinBox()
        self.timeout_spin.setRange(0.1, 60.0)
        self.timeout_spin.setValue(self.config.timeout)
        self.timeout_spin.setSuffix(" sec")
        advanced_layout.addRow("Timeout:", self.timeout_spin)

        layout.addWidget(self.advanced_group)

        # === Buttons ===
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.connect_btn = QPushButton("Connect")
        self.connect_btn.setDefault(True)
        self.connect_btn.clicked.connect(self.accept)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        button_layout.addWidget(self.connect_btn)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def _connect_signals(self):
        """Connect UI signals"""
        self.tcp_radio.toggled.connect(self._update_ui_state)
        self.udp_radio.toggled.connect(self._update_ui_state)
        self.client_radio.toggled.connect(self._update_ui_state)
        self.server_radio.toggled.connect(self._update_ui_state)

    def _update_ui_state(self):
        """Update UI based on current selections"""
        is_tcp = self.tcp_radio.isChecked()
        is_server = self.server_radio.isChecked()

        # Host field: disabled for server mode (always binds to 0.0.0.0)
        self.host_input.setEnabled(not is_server)
        if is_server:
            self.host_input.setText("0.0.0.0")
            self.host_input.setPlaceholderText("(listening on all interfaces)")
        else:
            if self.host_input.text() == "0.0.0.0":
                self.host_input.setText("127.0.0.1")
            self.host_input.setPlaceholderText("hostname or IP address")

        # TCP options only visible for TCP
        self.keepalive_check.setVisible(is_tcp)
        self.nodelay_check.setVisible(is_tcp)

        # Broadcast only for UDP
        self.broadcast_check.setVisible(not is_tcp)

        # Update button text
        if is_server:
            self.connect_btn.setText("Start Listening")
        else:
            self.connect_btn.setText("Connect")

    def get_config(self) -> NetworkConfig:
        """
        Get the configured NetworkConfig.

        Returns:
            NetworkConfig with user-specified settings
        """
        return NetworkConfig(
            host=self.host_input.text().strip() or "127.0.0.1",
            port=self.port_input.value(),
            protocol='TCP' if self.tcp_radio.isChecked() else 'UDP',
            mode='server' if self.server_radio.isChecked() else 'client',
            keepalive=self.keepalive_check.isChecked(),
            nodelay=self.nodelay_check.isChecked(),
            broadcast=self.broadcast_check.isChecked(),
            timeout=self.timeout_spin.value()
        )


class ConnectionHistoryDialog(QDialog):
    """
    Dialog showing connection history and favorites.

    Allows users to:
    - Browse saved favorite connections
    - Browse recent connection history
    - Select a connection to open
    """

    configSelected = pyqtSignal(object)  # Emits NetworkConfig

    def __init__(self, connection_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Connection History")
        self.setMinimumSize(400, 300)
        self.connection_manager = connection_manager
        self._setup_ui()
        self._load_data()

    def _setup_ui(self):
        """Setup dialog UI"""
        layout = QVBoxLayout(self)

        # Tab widget for Favorites / History
        tabs = QTabWidget()

        # Favorites tab
        self.favorites_list = QListWidget()
        self.favorites_list.itemDoubleClicked.connect(self._on_item_selected)
        tabs.addTab(self.favorites_list, "Favorites")

        # History tab
        self.history_list = QListWidget()
        self.history_list.itemDoubleClicked.connect(self._on_item_selected)
        tabs.addTab(self.history_list, "Recent")

        layout.addWidget(tabs)

        # Buttons
        button_layout = QHBoxLayout()

        connect_btn = QPushButton("Connect")
        connect_btn.clicked.connect(self._connect_selected)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(connect_btn)
        button_layout.addWidget(cancel_btn)

        layout.addLayout(button_layout)

    def _load_data(self):
        """Load favorites and history"""
        # Load favorites
        for config in self.connection_manager.get_favorites():
            item = QListWidgetItem(
                f"{config.protocol} {config.host}:{config.port}"
            )
            item.setData(Qt.ItemDataRole.UserRole, config)
            self.favorites_list.addItem(item)

        # Load history
        for hist in self.connection_manager.get_history():
            config = hist.config
            item = QListWidgetItem(
                f"{config.protocol} {config.host}:{config.port} "
                f"({hist.connect_count}x)"
            )
            item.setData(Qt.ItemDataRole.UserRole, config)
            self.history_list.addItem(item)

    def _on_item_selected(self, item):
        """Handle double-click on item"""
        config = item.data(Qt.ItemDataRole.UserRole)
        if config:
            self.configSelected.emit(config)
            self.accept()

    def _connect_selected(self):
        """Connect to selected item"""
        # Check favorites first, then history
        for list_widget in [self.favorites_list, self.history_list]:
            item = list_widget.currentItem()
            if item:
                self._on_item_selected(item)
                return

    def get_selected_config(self) -> NetworkConfig:
        """
        Get the selected configuration (if any).

        Returns:
            NetworkConfig or None if nothing selected
        """
        for list_widget in [self.favorites_list, self.history_list]:
            item = list_widget.currentItem()
            if item:
                return item.data(Qt.ItemDataRole.UserRole)
        return None


# Export public API
__all__ = ['QuickConnectDialog', 'ConnectionHistoryDialog']
