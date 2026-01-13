#!/usr/bin/env python3
"""
Network Terminal Dialog
Adapted from Serial Terminal for TCP/UDP connections

This module contains:
- NetworkTerminalPane: Individual terminal display for network connections
- SplitContainer: Manages recursive split pane layouts
- NetworkMonitorWindow: Main application window
"""

import sys
from typing import Optional, Dict, List
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTextEdit, QMenu, QMainWindow,
    QTabWidget, QStatusBar, QSplitter, QSizePolicy, QTabBar, QApplication,
    QFormLayout, QLineEdit, QSpinBox, QComboBox, QGroupBox, QPushButton,
    QCheckBox, QDoubleSpinBox, QDialog, QLabel, QWidgetAction
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QSize, QByteArray, QEvent
from PyQt6.QtGui import (
    QTextCursor, QTextCharFormat, QColor, QShortcut, QKeySequence,
    QPainter, QPixmap, QIcon, QPalette
)
from PyQt6.QtSvg import QSvgRenderer

from ui.windows.terminal_formatter import TerminalStreamFormatter
from ui.resources import resource_manager
from ui.components.ribbon_toolbar import RibbonToolbar
from ui.common.icons import Icons
from core.network_config import NetworkConfig
from core.network_worker import NetworkWorker, create_network_worker
from core.core import ResponsiveWindowManager


# ===== NETWORK TERMINAL PANE =====
class NetworkTerminalPane(QWidget):
    """
    Individual terminal display for network connections.

    Adapted from TerminalPane - key changes:
    - Uses NetworkConfig instead of SerialConfig
    - Uses NetworkWorker instead of SerialWorker
    - Removed baud rate detection logic
    - Added protocol-specific options
    """

    # Signals - SAME as TerminalPane
    focusChanged = pyqtSignal(bool)
    splitRequested = pyqtSignal(object, str)  # (source_pane, direction)
    closeRequested = pyqtSignal(object)        # source_pane

    def __init__(self, config: NetworkConfig, parent=None,
                 main_window=None, container=None):
        super().__init__(parent)
        self.config = config
        self.main_window = main_window
        self.container = container
        self.formatter = TerminalStreamFormatter()
        self.network_worker: Optional[NetworkWorker] = None
        self.is_connected = False
        self.rx_bytes = 0
        self.tx_bytes = 0

        # Line buffering for proper data handling
        self.line_buffer = ""
        self.buffer_timer = QTimer()
        self.buffer_timer.setSingleShot(True)
        self.buffer_timer.timeout.connect(self._flush_buffer)

        # Display settings
        self.encoding = 'utf-8'
        self.hex_display_mode = False
        self.local_echo_enabled = True

        # Help display management
        self.help_displayed = False
        self.auto_scroll_state_before_help = True

        self._setup_ui()
        self._setup_context_menu()

    def _setup_ui(self):
        """Setup the UI components"""
        # Main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Terminal display
        self.terminal = QTextEdit()
        self.terminal.setReadOnly(True)
        self.terminal.setUndoRedoEnabled(False)
        self.terminal.document().setMaximumBlockCount(10000)
        self.terminal.setFont(resource_manager.get_monospace_font(size=10))
        self.terminal.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        layout.addWidget(self.terminal)

        # Focus handling
        self.terminal.installEventFilter(self)

        # Install event filter for ESC key handling
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.installEventFilter(self)

    def _setup_context_menu(self):
        """Setup right-click context menu"""
        self.terminal.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.terminal.customContextMenuRequested.connect(self._show_context_menu)

    def _show_context_menu(self, position):
        """Show context menu at position"""
        menu = self._create_terminal_menu()
        menu.exec(self.terminal.mapToGlobal(position))

    def _create_terminal_menu(self) -> QMenu:
        """Create terminal context menu - adapted for network"""
        menu = QMenu(self)
        menu.addSeparator()

        # Connection section
        if self.is_connected:
            disconnect = menu.addAction("Disconnect")
            disconnect.triggered.connect(self.disconnect)
        else:
            connect = menu.addAction("Connect")
            connect.triggered.connect(self.connect)

        # Display Settings section
        menu.addSeparator()

        auto_scroll = menu.addAction(
            self.checkbox_icon(self.formatter.is_auto_scroll_enabled()),
            "Auto-scroll"
        )
        auto_scroll.triggered.connect(
            lambda: self._toggle_auto_scroll(not self.formatter.is_auto_scroll_enabled())
        )

        hex_mode = menu.addAction(
            self.checkbox_icon(self.hex_display_mode),
            "Hex Display Mode"
        )
        hex_mode.triggered.connect(
            lambda: self._toggle_hex_mode(not self.hex_display_mode)
        )

        local_echo = menu.addAction(
            self.checkbox_icon(self.local_echo_enabled),
            "Local Echo"
        )
        local_echo.triggered.connect(
            lambda: self._toggle_local_echo(not self.local_echo_enabled)
        )

        menu.addSeparator()

        # Font size submenu (KEEP)
        font_menu = menu.addMenu("Font Size")
        self._create_font_size_menu(font_menu)

        # Connection settings submenu (NEW - replaces baud rate/COM port menus)
        conn_menu = menu.addMenu("Connection Settings")
        self._create_connection_settings_menu(conn_menu)

        clear = menu.addAction("Clear Terminal")
        clear.triggered.connect(self._clear_terminal)

        menu.addSeparator()

        # Pane Management (KEEP)
        split_v = menu.addAction("Split Pane Vertically")
        split_v.setShortcut("Alt+Shift+-")
        split_v.triggered.connect(lambda: self.splitRequested.emit(self, 'vertical'))

        split_h = menu.addAction("Split Pane Horizontally")
        split_h.setShortcut("Alt+Shift++")
        split_h.triggered.connect(lambda: self.splitRequested.emit(self, 'horizontal'))

        close = menu.addAction("Close Pane")
        close.setShortcut("Ctrl+Shift+W")
        close.triggered.connect(lambda: self.closeRequested.emit(self))
        if self.container and len(self.container.panes) <= 1:
            close.setEnabled(False)

        menu.addSeparator()

        # Edit Actions (KEEP)
        copy = menu.addAction("Copy")
        copy.setShortcut("Ctrl+C")
        copy.triggered.connect(self.terminal.copy)
        copy.setEnabled(self.terminal.textCursor().hasSelection())

        select_all = menu.addAction("Select All")
        select_all.setShortcut("Ctrl+A")
        select_all.triggered.connect(self.terminal.selectAll)

        scroll_bottom = menu.addAction("Scroll to Bottom")
        scroll_bottom.triggered.connect(self._scroll_to_bottom)

        menu.addSeparator()

        help_action = menu.addAction("Help")
        help_action.triggered.connect(self._show_help)

        return menu

    def _create_connection_settings_menu(self, menu: QMenu):
        """Create connection settings submenu with editable host/port and favorite toggle"""
        # Show current connection info (read-only protocol/mode)
        info = menu.addAction(f"{self.config.protocol} {self.config.mode.title()}")
        info.setEnabled(False)

        menu.addSeparator()

        # === Editable Host Field ===
        host_widget = QWidget()
        host_layout = QHBoxLayout(host_widget)
        host_layout.setContentsMargins(8, 4, 8, 4)
        host_label = QLabel("Host:")
        host_label.setFixedWidth(40)
        self._host_edit = QLineEdit(self.config.host)
        self._host_edit.setMinimumWidth(150)
        # Disable host editing in server mode (bound to 0.0.0.0)
        if self.config.mode == 'server':
            self._host_edit.setEnabled(False)
            self._host_edit.setToolTip("Host cannot be changed in server mode")
        host_layout.addWidget(host_label)
        host_layout.addWidget(self._host_edit)

        host_action = QWidgetAction(menu)
        host_action.setDefaultWidget(host_widget)
        menu.addAction(host_action)

        # === Editable Port Field ===
        port_widget = QWidget()
        port_layout = QHBoxLayout(port_widget)
        port_layout.setContentsMargins(8, 4, 8, 4)
        port_label = QLabel("Port:")
        port_label.setFixedWidth(40)
        self._port_spin = QSpinBox()
        self._port_spin.setRange(1, 65535)
        self._port_spin.setValue(self.config.port)
        self._port_spin.setMinimumWidth(150)
        port_layout.addWidget(port_label)
        port_layout.addWidget(self._port_spin)

        port_action = QWidgetAction(menu)
        port_action.setDefaultWidget(port_widget)
        menu.addAction(port_action)

        menu.addSeparator()

        # === Apply & Reconnect Button ===
        if self.is_connected:
            apply_action = menu.addAction("Apply && Reconnect")
            apply_action.triggered.connect(self._apply_connection_settings)
        else:
            apply_action = menu.addAction("Apply Settings")
            apply_action.triggered.connect(self._apply_connection_settings)

        menu.addSeparator()

        # Favorite toggle
        if self.main_window and hasattr(self.main_window, 'connection_manager'):
            is_favorite = self.main_window.connection_manager.is_favorite(self.config)
            fav_text = "Remove from Favorites" if is_favorite else "Add to Favorites"
            favorite_action = menu.addAction(self.checkbox_icon(is_favorite), fav_text)
            favorite_action.triggered.connect(self._toggle_favorite)

    def _toggle_favorite(self):
        """Toggle favorite status for current connection"""
        if not self.main_window or not hasattr(self.main_window, 'connection_manager'):
            return

        manager = self.main_window.connection_manager

        if manager.is_favorite(self.config):
            manager.remove_favorite(self.config)
            self.formatter.append_status(self.terminal, "Removed from favorites", "status")
        else:
            manager.add_favorite(self.config)
            self.formatter.append_status(self.terminal, "Added to favorites", "status")

        # Refresh favorites dropdown in ribbon
        if hasattr(self.main_window, '_populate_favorites_menu'):
            self.main_window._populate_favorites_menu()

    def _apply_connection_settings(self):
        """Apply changed host/port settings and reconnect if needed"""
        # Get new values from the menu widgets
        new_host = self._host_edit.text().strip() if hasattr(self, '_host_edit') else self.config.host
        new_port = self._port_spin.value() if hasattr(self, '_port_spin') else self.config.port

        # Check if anything changed
        if new_host == self.config.host and new_port == self.config.port:
            self.formatter.append_status(self.terminal, "No changes to apply", "status")
            return

        # Validate host
        if not new_host:
            self.formatter.append_status(self.terminal, "Host cannot be empty", "error")
            return

        old_host = self.config.host
        old_port = self.config.port
        was_connected = self.is_connected

        # Create updated config (preserving all other settings)
        self.config = self.config.copy(host=new_host, port=new_port)

        # Update tab title if we have access to main window
        if self.main_window and hasattr(self.main_window, 'tab_widget'):
            tab_widget = self.main_window.tab_widget
            # Find our tab by checking containers
            for i in range(tab_widget.count()):
                widget = tab_widget.widget(i)
                if hasattr(widget, 'panes') and self in widget.panes:
                    tab_widget.setTabText(i, self.config.get_short_name())
                    break

        # Log the change
        if was_connected:
            self.formatter.append_status(
                self.terminal,
                f"Settings changed: {old_host}:{old_port} → {new_host}:{new_port}",
                "status"
            )
            # Disconnect and reconnect with new settings
            self.disconnect()
            # Give a small delay before reconnecting
            QTimer.singleShot(100, self.connect)
            self.formatter.append_status(self.terminal, "Reconnecting...", "status")
        else:
            self.formatter.append_status(
                self.terminal,
                f"Settings updated: {new_host}:{new_port}",
                "status"
            )

    def _create_font_size_menu(self, menu: QMenu):
        """Create font size submenu matching main GUI pattern"""
        common_sizes = [8, 10, 12, 14, 16]
        current_size = self.terminal.font().pointSize()

        for size in common_sizes:
            action = menu.addAction(f"{size}pt")
            action.triggered.connect(lambda checked, s=size: self._set_font_size(s))
            if size == current_size:
                action.setIcon(self.checkbox_icon(True))

        menu.addSeparator()

        # All sizes submenu
        all_sizes_menu = menu.addMenu("All Sizes")
        for size in range(6, 25):
            if size not in common_sizes:
                action = all_sizes_menu.addAction(f"{size}pt")
                action.triggered.connect(lambda checked, s=size: self._set_font_size(s))
                if size == current_size:
                    action.setIcon(self.checkbox_icon(True))

        menu.addSeparator()

        # Quick actions
        increase_font = menu.addAction("Increase")
        increase_font.setShortcut("Ctrl++")
        increase_font.triggered.connect(self._increase_font_size)

        decrease_font = menu.addAction("Decrease")
        decrease_font.setShortcut("Ctrl+-")
        decrease_font.triggered.connect(self._decrease_font_size)

        menu.addSeparator()

        reset_font = menu.addAction("Reset to Default")
        reset_font.triggered.connect(self._reset_font_size)

    def checkbox_icon(self, checked: bool) -> QIcon:
        """Generate checkbox icon using palette colors"""
        palette = self.palette()
        border_color = palette.color(QPalette.ColorRole.Mid).name()
        bg_color = palette.color(QPalette.ColorRole.Base).name()
        check_color = palette.color(QPalette.ColorRole.Highlight).name()

        if checked:
            svg = f'''<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
                <rect x="0.5" y="0.5" width="15" height="15" fill="{border_color}" stroke="{border_color}" stroke-width="1"/>
                <rect x="2" y="2" width="12" height="12" fill="{bg_color}"/>
                <path d="M4 8l2 2 6-6" stroke="{check_color}" stroke-width="1.5" fill="none" stroke-linecap="round"/>
            </svg>'''
        else:
            svg = f'''<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
                <rect x="0.5" y="0.5" width="15" height="15" fill="{border_color}" stroke="{border_color}" stroke-width="1"/>
            </svg>'''

        renderer = QSvgRenderer(QByteArray(svg.encode()))
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return QIcon(pixmap)

    def eventFilter(self, obj, event):
        """Handle focus events and keyboard input for local echo"""
        if obj == self.terminal:
            if event.type() == QEvent.Type.FocusIn:
                self.focusChanged.emit(True)
            elif event.type() == QEvent.Type.FocusOut:
                self.focusChanged.emit(False)
            elif event.type() == QEvent.Type.KeyPress and self.local_echo_enabled:
                return self._handle_key_press(event)
        elif obj == self and event.type() == QEvent.Type.KeyPress:
            if event.key() == Qt.Key.Key_Escape and self.help_displayed:
                self._dismiss_help()
                return True
        return super().eventFilter(obj, event)

    def _handle_key_press(self, event):
        """Handle key press events for local echo"""
        if not self.is_connected:
            return False

        key = event.key()
        text = event.text()

        # Handle special keys
        if key == Qt.Key.Key_Return or key == Qt.Key.Key_Enter:
            data_to_send = "\r\n"
            self._send_raw_data(data_to_send)
            self._echo_local_data(data_to_send)
            return True
        elif key == Qt.Key.Key_Backspace:
            return True
        elif key == Qt.Key.Key_Tab:
            data_to_send = "\t"
            self._send_raw_data(data_to_send)
            self._echo_local_data(data_to_send)
            return True
        elif text and text.isprintable():
            self._send_raw_data(text)
            self._echo_local_data(text)
            return True

        return False

    def _send_raw_data(self, data: str):
        """Send raw data to network socket without local echo formatting"""
        if self.network_worker and self.is_connected:
            try:
                bytes_data = data.encode(self.encoding)
                self.network_worker.write(bytes_data)
                self.tx_bytes += len(bytes_data)
            except UnicodeEncodeError as e:
                self.formatter.append_status(
                    self.terminal,
                    f"Send encoding error: {str(e)}",
                    "error"
                )

    def _echo_local_data(self, data: str):
        """Echo data locally without timestamps or formatting"""
        cursor = self.terminal.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        # Set color for local echo (different from received data)
        format = QTextCharFormat()
        format.setForeground(QColor("#90EE90"))  # Light green for local echo
        cursor.setCharFormat(format)

        # Insert the text
        cursor.insertText(data)

        # Auto-scroll if enabled
        if self.formatter.is_auto_scroll_enabled():
            self.formatter.force_scroll_to_bottom(self.terminal)

    def connect(self):
        """Connect using network worker"""
        if not self.network_worker:
            self.network_worker = create_network_worker(self.config)
            self.network_worker.dataReceived.connect(
                self._on_data_received, Qt.ConnectionType.QueuedConnection
            )
            self.network_worker.errorOccurred.connect(
                self._on_error, Qt.ConnectionType.QueuedConnection
            )
            self.network_worker.connectionStateChanged.connect(
                self._on_connection_state_changed, Qt.ConnectionType.QueuedConnection
            )
            self.network_worker.start()

    def disconnect(self):
        """Disconnect from network"""
        self.cleanup()

    def cleanup(self):
        """Single point of cleanup for terminal pane"""
        old_is_connected = self.is_connected
        self.is_connected = False

        # Block signals from worker EARLY to prevent race conditions
        # This ensures no new signals are processed during cleanup
        if self.network_worker:
            self.network_worker.blockSignals(True)

        # Update toolbar/ribbon
        if old_is_connected and self.main_window and hasattr(self.main_window, '_update_ribbon_connection_state'):
            self.main_window._update_ribbon_connection_state()

        # Show disconnection message
        if old_is_connected:
            self._format_connection_end()

        # Stop buffer timer
        if self.buffer_timer:
            try:
                self.buffer_timer.stop()
                try:
                    self.buffer_timer.timeout.disconnect()
                except (TypeError, RuntimeError):
                    pass
            except (RuntimeError, AttributeError):
                pass

        # Stop worker thread
        if self.network_worker:
            self.network_worker.stop()

            if not self.network_worker.wait(5000):
                print(f"Warning: Network worker did not stop gracefully for "
                      f"{self.config.host}:{self.config.port}")

            try:
                self.network_worker.dataReceived.disconnect(self._on_data_received)
                self.network_worker.errorOccurred.disconnect(self._on_error)
                self.network_worker.connectionStateChanged.disconnect(self._on_connection_state_changed)
            except (TypeError, RuntimeError):
                pass

            self.network_worker = None

        self.line_buffer = ""

    def _on_data_received(self, data: bytes):
        """Handle received data - simplified from serial version (no baud rate detection)"""
        if not self.network_worker:
            return

        try:
            self.rx_bytes += len(data)

            # Handle hex display mode
            if self.hex_display_mode:
                hex_data = ' '.join(f'{b:02X}' for b in data)
                ascii_data = ''.join(chr(b) if 32 <= b < 127 else '.' for b in data)
                formatted_data = f"HEX: {hex_data} | ASCII: {ascii_data}"

                self.formatter.append_data(
                    self.terminal,
                    formatted_data,
                    "incoming",
                    show_timestamp=True
                )
                return

            # Decode bytes - simplified error handling (no baud rate detection)
            try:
                text_data = data.decode(self.encoding)
            except UnicodeDecodeError:
                text_data = data.decode(self.encoding, errors='replace')

            # Normalize line endings
            text_data = text_data.replace('\r\n', '\n').replace('\r', '\n')

            # Add to line buffer
            self.line_buffer += text_data

            # Process complete lines
            lines = self.line_buffer.split('\n')

            # Display all complete lines
            for line in lines[:-1]:
                if line:
                    self.formatter.append_data(
                        self.terminal,
                        line,
                        "incoming",
                        show_timestamp=True
                    )

            # Keep the last (potentially incomplete) line
            self.line_buffer = lines[-1]

            # Start timer for incomplete lines
            if self.line_buffer:
                self.buffer_timer.stop()
                self.buffer_timer.start(1000)

        except Exception as e:
            self.formatter.append_status(
                self.terminal,
                f"Data processing error: {e}",
                "error"
            )

    def _on_error(self, error_msg: str):
        """Handle network errors"""
        if not self.network_worker:
            return
        self.formatter.append_status(self.terminal, error_msg, "error")

    def _on_connection_state_changed(self, connected: bool):
        """Handle connection state changes"""
        if not self.network_worker:
            return

        if connected:
            self.is_connected = True
            if self.main_window and hasattr(self.main_window, '_update_ribbon_connection_state'):
                self.main_window._update_ribbon_connection_state()
            self._format_connection_start()
        else:
            self.cleanup()

    def _format_connection_start(self):
        """Format connection start message for network"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        mode_str = "Listening on" if self.config.mode == 'server' else "Connected to"
        self.formatter.append_separator(self.terminal, f"Connection established - {timestamp}")
        self.formatter.append_status(
            self.terminal,
            f"{self.config.protocol} {mode_str} {self.config.host}:{self.config.port}",
            "status"
        )
        self.formatter.append_separator(self.terminal)

    def _format_connection_end(self):
        """Format connection end message for network"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.formatter.append_separator(self.terminal, f"Connection closed - {timestamp}")
        self.formatter.append_status(
            self.terminal,
            f"{self.config.protocol} {self.config.host}:{self.config.port} disconnected",
            "status"
        )
        self.formatter.append_separator(self.terminal)

    def send_data(self, data: str):
        """Send data to network"""
        if self.network_worker and self.is_connected:
            try:
                bytes_data = data.encode(self.encoding)
                self.network_worker.write(bytes_data)
                self.tx_bytes += len(bytes_data)
                self.formatter.append_data(
                    self.terminal,
                    data.strip(),
                    "outgoing",
                    show_timestamp=True
                )
            except UnicodeEncodeError as e:
                self.formatter.append_status(
                    self.terminal,
                    f"Send encoding error: {str(e)}",
                    "error"
                )

    def get_status_info(self) -> str:
        """Get status information for status bar"""
        status = "Disconnected"
        if self.is_connected:
            if self.config.mode == 'server':
                status = "Listening"
            else:
                status = "Connected"

        rx_str = self._format_bytes(self.rx_bytes)
        tx_str = self._format_bytes(self.tx_bytes)

        echo_indicator = " | Echo: ON" if self.local_echo_enabled else ""

        # Network format: "TCP 192.168.1.1:5000: Connected | RX: 1KB | TX: 256B"
        return (f"{self.config.protocol} {self.config.host}:{self.config.port}: "
                f"{status} | RX: {rx_str} | TX: {tx_str}{echo_indicator}")

    def _format_bytes(self, bytes_count: int) -> str:
        """Format byte count for display"""
        if bytes_count < 1024:
            return f"{bytes_count}B"
        elif bytes_count < 1024 * 1024:
            return f"{bytes_count / 1024:.1f}KB"
        else:
            return f"{bytes_count / (1024 * 1024):.1f}MB"

    def _flush_buffer(self):
        """Flush remaining data in buffer"""
        if self.line_buffer:
            self.formatter.append_data(
                self.terminal,
                self.line_buffer,
                "incoming",
                show_timestamp=True
            )
            self.line_buffer = ""

    def _toggle_auto_scroll(self, enabled: bool):
        """Toggle auto-scroll with formatter integration"""
        self.formatter.set_auto_scroll_enabled(enabled)
        if enabled:
            self.formatter.force_scroll_to_bottom(self.terminal)

    def _toggle_hex_mode(self, enabled: bool):
        """Toggle hex display mode"""
        self.hex_display_mode = enabled
        if enabled:
            self.formatter.append_status(self.terminal, "Hex display mode enabled", "status")
        else:
            self.formatter.append_status(self.terminal, "Hex display mode disabled", "status")

    def _toggle_local_echo(self, enabled: bool):
        """Toggle local echo mode"""
        self.local_echo_enabled = enabled
        if enabled:
            self.formatter.append_status(
                self.terminal,
                "Local echo enabled - Start typing to send data",
                "status"
            )
        else:
            self.formatter.append_status(
                self.terminal,
                "Local echo disabled - Terminal is read-only",
                "status"
            )
        self.focusChanged.emit(True)

    def _scroll_to_bottom(self):
        """Manually scroll to bottom"""
        self.formatter.force_scroll_to_bottom(self.terminal)

    def _clear_terminal(self):
        """Clear the terminal display with proper formatter integration"""
        self.formatter.clear(self.terminal)
        self.line_buffer = ""
        self.buffer_timer.stop()
        self.help_displayed = False

        if self.is_connected:
            self.formatter.append_separator(self.terminal, "Terminal cleared")

    def _set_font_size(self, size: int):
        """Set terminal font size"""
        font = self.terminal.font()
        font.setPointSize(size)
        self.terminal.setFont(font)

    def _increase_font_size(self):
        """Increase terminal font size"""
        font = self.terminal.font()
        if font.pointSize() < 24:
            font.setPointSize(font.pointSize() + 1)
            self.terminal.setFont(font)

    def _decrease_font_size(self):
        """Decrease terminal font size"""
        font = self.terminal.font()
        if font.pointSize() > 8:
            font.setPointSize(font.pointSize() - 1)
            self.terminal.setFont(font)

    def _reset_font_size(self):
        """Reset font size to default"""
        self.terminal.setFont(resource_manager.get_monospace_font(size=10))

    def _show_help(self):
        """Show help inline in the terminal window"""
        if self.help_displayed:
            self.formatter.append_status(
                self.terminal,
                "Help already displayed. Press ESC to return to auto-scroll mode.",
                "warning"
            )
            return

        self.auto_scroll_state_before_help = self.formatter.is_auto_scroll_enabled()
        self.formatter.set_auto_scroll_enabled(False)
        self.help_displayed = True

        self.formatter.append_separator(self.terminal, "NETWORK TERMINAL HELP")
        self.formatter.append_status(self.terminal, "Start Of Help Content", "help")

        help_sections = [
            ("CONNECTION", [
                "- Supports TCP and UDP protocols",
                "- Client mode: Connect to remote hosts",
                "- Server mode: Listen for incoming connections",
                "- Real-time data display"
            ]),
            ("DISPLAY OPTIONS", [
                "- Auto-scroll: Enabled by default",
                "- Hex Display Mode: Show data as hex",
                "- Local Echo: Echo typed characters"
            ]),
            ("KEYBOARD SHORTCUTS", [
                "Navigation:",
                "- Alt+Arrow Keys: Navigate between panes",
                "- Ctrl+Tab: Next tab",
                "- Ctrl+Shift+Tab: Previous tab",
                "",
                "Pane Management:",
                "- Alt+Shift+-: Split pane vertically",
                "- Alt+Shift++: Split pane horizontally",
                "- Ctrl+Shift+W: Close current pane",
                "",
                "Terminal Actions:",
                "- Ctrl+C: Copy selected text",
                "- Ctrl+A: Select all text",
                "- Ctrl++: Increase font size",
                "- Ctrl+-: Decrease font size",
                "",
                "Window Management:",
                "- Ctrl+N: New connection",
                "- Ctrl+W: Close current tab",
                "- F1: Show this help"
            ])
        ]

        cursor = self.terminal.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        for section_title, section_items in help_sections:
            cursor.insertText(f"\n{section_title}\n", self.formatter._get_format('help', bold=True))
            for item in section_items:
                if item:
                    cursor.insertText(f"{item}\n", self.formatter._get_format('help'))
                else:
                    cursor.insertText("\n")
            cursor.insertText("\n")

        self.formatter.append_status(
            self.terminal,
            "End Of Help Content - Press ESC to return to auto-scroll mode",
            "help"
        )
        self.formatter.append_separator(self.terminal)

    def _dismiss_help(self):
        """Dismiss help display and restore auto-scroll state"""
        if not self.help_displayed:
            return

        self.help_displayed = False
        self.formatter.set_auto_scroll_enabled(self.auto_scroll_state_before_help)
        self.formatter.append_status(self.terminal, "Help dismissed - Auto-scroll restored", "status")

        if self.auto_scroll_state_before_help:
            self.formatter.force_scroll_to_bottom(self.terminal)


# ===== WELCOME CONFIG WIDGET =====
class WelcomeConfigWidget(QWidget):
    """Welcome screen with network connection configuration"""

    connectionRequested = pyqtSignal(object)  # NetworkConfig

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        """Setup UI components"""
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setSpacing(0)
        self.main_layout.setContentsMargins(8, 8, 8, 8)

        self.main_layout.addStretch()

        # Create centered container
        self.center_container = QWidget()
        self.center_container.setMaximumWidth(400)
        self.center_container.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        self._create_config_section()

        h_layout = QHBoxLayout()
        h_layout.addStretch()
        h_layout.addWidget(self.center_container)
        h_layout.addStretch()

        self.main_layout.addLayout(h_layout)
        self.main_layout.addStretch()

    def _create_config_section(self):
        """Create the connection configuration section"""
        form_layout = QFormLayout(self.center_container)
        form_layout.setSpacing(8)

        # Protocol selection
        self.protocol_combo = QComboBox()
        self.protocol_combo.addItems(["TCP", "UDP"])
        self.protocol_combo.currentTextChanged.connect(self._on_protocol_changed)
        form_layout.addRow("Protocol:", self.protocol_combo)

        # Mode selection
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Client", "Server"])
        self.mode_combo.currentTextChanged.connect(self._on_mode_changed)
        form_layout.addRow("Mode:", self.mode_combo)

        # Host input
        self.host_input = QLineEdit("127.0.0.1")
        self.host_input.setPlaceholderText("hostname or IP address")
        form_layout.addRow("Host:", self.host_input)

        # Port input
        self.port_input = QSpinBox()
        self.port_input.setRange(1, 65535)
        self.port_input.setValue(5000)
        form_layout.addRow("Port:", self.port_input)

        # === Advanced Options ===
        # TCP options
        self.keepalive_check = QCheckBox("Enable TCP Keepalive")
        self.keepalive_check.setChecked(False)
        form_layout.addRow("", self.keepalive_check)

        self.nodelay_check = QCheckBox("Disable Nagle (TCP_NODELAY)")
        self.nodelay_check.setChecked(False)
        form_layout.addRow("", self.nodelay_check)

        # UDP options
        self.broadcast_check = QCheckBox("Enable Broadcast")
        self.broadcast_check.setChecked(False)
        self.broadcast_check.setVisible(False)  # Hidden by default (TCP selected)
        form_layout.addRow("", self.broadcast_check)

        # Timeout
        self.timeout_spin = QDoubleSpinBox()
        self.timeout_spin.setRange(0.1, 60.0)
        self.timeout_spin.setValue(5.0)
        self.timeout_spin.setSuffix(" sec")
        form_layout.addRow("Timeout:", self.timeout_spin)

        # Connect button
        self.connect_btn = QPushButton("Connect")
        self.connect_btn.clicked.connect(self._on_connect)
        form_layout.addRow("", self.connect_btn)

    def _on_protocol_changed(self, protocol: str):
        """Handle protocol change - show/hide relevant options"""
        is_tcp = protocol == "TCP"
        self.keepalive_check.setVisible(is_tcp)
        self.nodelay_check.setVisible(is_tcp)
        self.broadcast_check.setVisible(not is_tcp)

    def _on_mode_changed(self, mode: str):
        """Handle mode change"""
        if mode == "Server":
            self.host_input.setText("0.0.0.0")
            self.host_input.setEnabled(False)
            self.connect_btn.setText("Start Listening")
        else:
            self.host_input.setText("127.0.0.1")
            self.host_input.setEnabled(True)
            self.connect_btn.setText("Connect")

    def _on_connect(self):
        """Handle connect button click"""
        config = NetworkConfig(
            host=self.host_input.text().strip() or "127.0.0.1",
            port=self.port_input.value(),
            protocol=self.protocol_combo.currentText(),
            mode=self.mode_combo.currentText().lower(),
            keepalive=self.keepalive_check.isChecked(),
            nodelay=self.nodelay_check.isChecked(),
            broadcast=self.broadcast_check.isChecked(),
            timeout=self.timeout_spin.value()
        )
        self.connectionRequested.emit(config)


# ===== SPLIT CONTAINER =====
class SplitContainer(QWidget):
    """Manages split pane layout with recursive splitting"""

    activePaneChanged = pyqtSignal(object)  # NetworkTerminalPane

    def __init__(self, initial_config: NetworkConfig, parent=None, main_window=None):
        super().__init__(parent)
        self.panes: List[NetworkTerminalPane] = []
        self.active_pane: Optional[NetworkTerminalPane] = None
        self.main_window = main_window
        self._setup_ui(initial_config)

    def _setup_ui(self, config: NetworkConfig):
        """Setup initial UI"""
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        # Create initial pane
        initial_pane = self._create_pane(config)
        self.main_layout.addWidget(initial_pane)
        self._set_active_pane(initial_pane)

    def _create_pane(self, config: NetworkConfig) -> NetworkTerminalPane:
        """Create a new terminal pane"""
        pane = NetworkTerminalPane(config, main_window=self.main_window, container=self)
        pane.splitRequested.connect(self._split_pane)
        pane.closeRequested.connect(self._close_pane)
        pane.focusChanged.connect(lambda focused: self._on_pane_focus(pane, focused))

        self.panes.append(pane)
        return pane

    def _create_welcome_pane(self):
        """Create a pane with welcome configuration widget"""
        welcome_pane = QWidget()
        layout = QVBoxLayout(welcome_pane)
        layout.setContentsMargins(0, 0, 0, 0)

        welcome_widget = WelcomeConfigWidget()
        welcome_widget.connectionRequested.connect(
            lambda config: self._replace_welcome_with_terminal(welcome_pane, config)
        )

        layout.addWidget(welcome_widget)
        return welcome_pane

    def _replace_welcome_with_terminal(self, welcome_pane, config: NetworkConfig):
        """Replace welcome pane with actual terminal pane"""
        parent = welcome_pane.parent()
        if isinstance(parent, QSplitter):
            index = parent.indexOf(welcome_pane)

            terminal_pane = self._create_pane(config)
            parent.replaceWidget(index, terminal_pane)
            welcome_pane.deleteLater()

            terminal_pane.connect()
            terminal_pane.terminal.setFocus()
            self._set_active_pane(terminal_pane)

    def _split_pane(self, source_pane: NetworkTerminalPane, direction: str):
        """Split a pane horizontally or vertically"""
        new_pane = self._create_welcome_pane()

        parent_widget = source_pane.parent()

        if isinstance(parent_widget, QSplitter):
            index = parent_widget.indexOf(source_pane)
            new_orientation = Qt.Orientation.Horizontal if direction == 'vertical' else Qt.Orientation.Vertical

            if parent_widget.orientation() == new_orientation:
                parent_widget.insertWidget(index + 1, new_pane)
            else:
                nested_splitter = self._create_splitter(new_orientation)
                parent_widget.replaceWidget(index, nested_splitter)
                nested_splitter.addWidget(source_pane)
                nested_splitter.addWidget(new_pane)
                nested_splitter.setSizes([500, 500])
        else:
            orientation = Qt.Orientation.Horizontal if direction == 'vertical' else Qt.Orientation.Vertical
            splitter = self._create_splitter(orientation)

            self.main_layout.replaceWidget(source_pane, splitter)
            splitter.addWidget(source_pane)
            splitter.addWidget(new_pane)
            splitter.setSizes([500, 500])

        new_pane.setFocus()

    def _create_splitter(self, orientation: Qt.Orientation) -> QSplitter:
        """Create a styled splitter"""
        splitter = QSplitter(orientation)
        splitter.setHandleWidth(4)
        return splitter

    def cleanup(self):
        """Single point of cleanup for split container"""
        for pane in self.panes:
            try:
                pane.cleanup()
            except Exception as e:
                print(f"Error cleaning up pane: {e}")

        self.panes.clear()
        self.active_pane = None

    def _close_pane(self, pane: NetworkTerminalPane):
        """Close a pane and reorganize layout"""
        if len(self.panes) == 1:
            return

        pane.cleanup()
        self.panes.remove(pane)

        parent = pane.parent()

        if isinstance(parent, QSplitter):
            pane.setParent(None)
            pane.deleteLater()

            if parent.count() == 1:
                remaining_widget = parent.widget(0)
                grandparent = parent.parent()

                if isinstance(grandparent, QSplitter):
                    index = grandparent.indexOf(parent)
                    grandparent.replaceWidget(index, remaining_widget)
                else:
                    self.main_layout.replaceWidget(parent, remaining_widget)

                parent.deleteLater()
        else:
            pane.setParent(None)
            pane.deleteLater()

        if self.panes:
            self._set_active_pane(self.panes[0])
            self.panes[0].terminal.setFocus()

    def _on_pane_focus(self, pane: NetworkTerminalPane, focused: bool):
        """Handle pane focus changes"""
        if focused and pane in self.panes:
            self._set_active_pane(pane)

    def _set_active_pane(self, pane: NetworkTerminalPane):
        """Set the active pane"""
        self.active_pane = pane
        self.activePaneChanged.emit(pane)

    def navigate_panes(self, direction: str):
        """Navigate between panes using keyboard"""
        if not self.active_pane or len(self.panes) < 2:
            return

        current_index = self.panes.index(self.active_pane)

        if direction in ['left', 'up']:
            new_index = (current_index - 1) % len(self.panes)
        else:
            new_index = (current_index + 1) % len(self.panes)

        self.panes[new_index].terminal.setFocus()


# ===== MAIN WINDOW =====
class NetworkMonitorWindow(QMainWindow):
    """Main window for Network Terminal with tab management"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Terminal")
        self.setMinimumSize(800, 600)

        # Add ConnectionManager integration
        from core.connection_manager import ConnectionManager
        self.connection_manager = ConnectionManager()

        # Set custom window icon
        icon_pixmap = QPixmap(64, 64)
        icon_pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(icon_pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        svg_renderer = QSvgRenderer()
        svg_renderer.load(Icons.terminal_settings(self.palette()).encode())
        svg_renderer.render(painter)
        painter.end()
        self.setWindowIcon(QIcon(icon_pixmap))

        self.tabs: Dict[QWidget, SplitContainer] = {}
        self.close_button_icon = None

        self._setup_ui()
        self._setup_shortcuts()
        self._apply_window_style()

        # Show connection dialog after window is shown
        QTimer.singleShot(200, self._show_initial_connection_dialog)

    def _setup_ui(self):
        """Setup main window UI"""
        # Create ribbon toolbar (no menu bar - all features in ribbon)
        self.ribbon = RibbonToolbar()
        self.addToolBar(self.ribbon)

        # Connect ribbon signals
        self._connect_ribbon_signals()

        # Populate dropdown menus
        self._populate_recent_menu()
        self._populate_favorites_menu()

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self._close_tab)
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

        self._setup_close_button_icon()
        main_layout.addWidget(self.tab_widget)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Status update timer
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self._update_status_bar)
        self.status_timer.start(1000)

    def _connect_ribbon_signals(self):
        """Connect ribbon toolbar signals"""
        self.ribbon.new_connection.connect(self._new_connection)
        self.ribbon.refresh_ports.connect(self._refresh)
        self.ribbon.toggle_connection.connect(self._toggle_connection)
        self.ribbon.clear_terminal.connect(self._clear_current_terminal)
        self.ribbon.show_settings.connect(self._show_settings_menu)
        self.ribbon.recent_selected.connect(self._create_tab)
        self.ribbon.favorite_selected.connect(self._create_tab)

    def _populate_recent_menu(self):
        """Populate recent connections dropdown in ribbon"""
        recent = self.connection_manager.get_recent_configs(limit=10)
        self.ribbon.populate_recent_menu(recent)

    def _populate_favorites_menu(self):
        """Populate favorites dropdown in ribbon"""
        favorites = self.connection_manager.get_favorites()
        self.ribbon.populate_favorites_menu(favorites)

    def _setup_close_button_icon(self):
        """Set up custom close button - done in _apply_window_style"""
        pass

    def _apply_window_style(self):
        """Apply window style"""
        palette = self.palette()
        window_bg = palette.color(QPalette.ColorRole.Window)

        tab_palette = self.tab_widget.palette()
        tab_palette.setColor(QPalette.ColorRole.Base, window_bg)
        self.tab_widget.setPalette(tab_palette)

        # Create custom white X close button icon
        close_icon_svg = """<svg width="16" height="16" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">
            <line x1="4" y1="4" x2="12" y2="12" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/>
            <line x1="12" y1="4" x2="4" y2="12" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/>
        </svg>"""

        renderer = QSvgRenderer(QByteArray(close_icon_svg.encode()))
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()

        self.close_button_icon = QIcon(pixmap)
        self._apply_close_icon_to_tabs()

    def _apply_close_icon_to_tabs(self):
        """Apply custom close button icon to all tabs"""
        if not self.close_button_icon:
            return

        for i in range(self.tab_widget.count()):
            button = self.tab_widget.tabBar().tabButton(i, QTabBar.ButtonPosition.RightSide)
            if button:
                button.setIcon(self.close_button_icon)
                button.setIconSize(QSize(12, 12))
                button.setStyleSheet("border: none; background: transparent;")

    def _setup_shortcuts(self):
        """Setup global keyboard shortcuts"""
        QShortcut(QKeySequence("Ctrl+N"), self, self._new_connection)
        QShortcut(QKeySequence("Ctrl+W"), self, lambda: self._close_tab(self.tab_widget.currentIndex()))
        QShortcut(QKeySequence("Ctrl+Tab"), self, self._next_tab)
        QShortcut(QKeySequence("Ctrl+Shift+Tab"), self, self._prev_tab)

        QShortcut(QKeySequence("Alt+Left"), self, lambda: self._navigate_panes("left"))
        QShortcut(QKeySequence("Alt+Right"), self, lambda: self._navigate_panes("right"))
        QShortcut(QKeySequence("Alt+Up"), self, lambda: self._navigate_panes("up"))
        QShortcut(QKeySequence("Alt+Down"), self, lambda: self._navigate_panes("down"))

        QShortcut(QKeySequence("Alt+Shift+-"), self, lambda: self._split_current_pane("vertical"))
        QShortcut(QKeySequence("Alt+Shift++"), self, lambda: self._split_current_pane("horizontal"))

        QShortcut(QKeySequence("Ctrl+Shift+W"), self, self._close_current_pane)

        # Font size shortcuts
        QShortcut(QKeySequence("Ctrl++"), self, self._increase_active_pane_font)
        QShortcut(QKeySequence("Ctrl+="), self, self._increase_active_pane_font)  # For keyboards without numpad
        QShortcut(QKeySequence("Ctrl+-"), self, self._decrease_active_pane_font)

    def _increase_active_pane_font(self):
        """Increase font size of the active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container.active_pane._increase_font_size()

    def _decrease_active_pane_font(self):
        """Decrease font size of the active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container.active_pane._decrease_font_size()

    def _show_initial_connection_dialog(self):
        """Show welcome tab when window opens"""
        if self.tab_widget.count() == 0:
            self._show_welcome_tab()

    def _refresh(self):
        """Refresh - placeholder for network (no port scanning needed)"""
        pass

    def _show_welcome_tab(self):
        """Show welcome tab with network configuration"""
        if self._has_welcome_tab():
            return

        try:
            welcome_widget = WelcomeConfigWidget()
            welcome_widget.connectionRequested.connect(self._handle_welcome_connection)

            index = self.tab_widget.addTab(welcome_widget, "New tab")
            self.tab_widget.setCurrentIndex(index)
            self._apply_close_icon_to_tabs()
        except Exception as e:
            print(f"Error creating welcome tab: {e}")

    def _has_welcome_tab(self) -> bool:
        """Check if a welcome tab already exists"""
        for i in range(self.tab_widget.count()):
            if self.tab_widget.tabText(i) == "New tab":
                return True
        return False

    def _handle_welcome_connection(self, config: NetworkConfig):
        """Handle connection request from welcome widget"""
        try:
            self._create_tab(config)
            self._remove_welcome_tab()
        except Exception as e:
            print(f"Error handling welcome connection: {e}")

    def _remove_welcome_tab(self):
        """Safely remove welcome tab"""
        try:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == "New tab":
                    widget = self.tab_widget.widget(i)
                    self.tab_widget.removeTab(i)
                    if widget:
                        widget.deleteLater()
                    break
        except Exception as e:
            print(f"Error removing welcome tab: {e}")

    def _new_connection(self):
        """Create new tab with inline connection settings"""
        try:
            welcome_widget = WelcomeConfigWidget()
            welcome_widget.connectionRequested.connect(
                lambda config, w=welcome_widget: self._handle_inline_connection(config, w)
            )

            index = self.tab_widget.addTab(welcome_widget, "New tab")
            self.tab_widget.setCurrentIndex(index)
            self._apply_close_icon_to_tabs()
        except Exception as e:
            print(f"Error creating new tab: {e}")

    def _handle_inline_connection(self, config: NetworkConfig, welcome_widget):
        """Handle connection from inline welcome widget and replace it with terminal"""
        try:
            # Find and remove this specific welcome widget
            for i in range(self.tab_widget.count()):
                if self.tab_widget.widget(i) == welcome_widget:
                    self.tab_widget.removeTab(i)
                    welcome_widget.deleteLater()
                    break

            # Create the actual connection tab
            self._create_tab(config)
        except Exception as e:
            print(f"Error handling inline connection: {e}")

    def _create_tab(self, config: NetworkConfig):
        """Create a new tab with split container"""
        container = SplitContainer(config, main_window=self)
        container.activePaneChanged.connect(self._on_active_pane_changed)

        self.tabs[container] = container

        # Use short name for tab title
        tab_title = config.get_short_name()
        index = self.tab_widget.addTab(container, tab_title)
        self.tab_widget.setCurrentIndex(index)
        self._apply_close_icon_to_tabs()

        # Record in history
        self.connection_manager.record_connection(config)

        # Refresh recent menu
        self._populate_recent_menu()

        # Auto-connect the first pane
        if container.active_pane:
            container.active_pane.connect()

    def cleanup(self):
        """Cleanup all tabs"""
        for container in self.tabs.values():
            container.cleanup()

    def _close_tab(self, index: int):
        """Close a tab and cleanup"""
        if index < 0 or index >= self.tab_widget.count():
            return

        widget = self.tab_widget.widget(index)
        if not widget:
            return

        # Don't allow closing the last welcome tab
        if (self.tab_widget.count() == 1 and
            self.tab_widget.tabText(index) == "New tab"):
            return

        if widget in self.tabs:
            container = self.tabs[widget]
            try:
                container.cleanup()
                del self.tabs[widget]
            except Exception as e:
                print(f"Error cleaning up container: {e}")

        self.tab_widget.removeTab(index)

        if widget:
            widget.deleteLater()

        QTimer.singleShot(0, self._check_empty_tabs)

    def _check_empty_tabs(self):
        """Check if tabs are empty and show welcome tab if needed"""
        try:
            if self.tab_widget.count() == 0:
                self._show_welcome_tab()
        except Exception as e:
            print(f"Error checking empty tabs: {e}")

    def _next_tab(self):
        """Switch to next tab"""
        current = self.tab_widget.currentIndex()
        count = self.tab_widget.count()
        if count > 0:
            self.tab_widget.setCurrentIndex((current + 1) % count)

    def _prev_tab(self):
        """Switch to previous tab"""
        current = self.tab_widget.currentIndex()
        count = self.tab_widget.count()
        if count > 0:
            self.tab_widget.setCurrentIndex((current - 1) % count)

    def _on_tab_changed(self, index: int):
        """Handle tab change"""
        self._update_status_bar()
        self._update_ribbon_connection_state()

    def _get_current_container(self) -> Optional[SplitContainer]:
        """Get current tab's split container"""
        widget = self.tab_widget.currentWidget()
        return self.tabs.get(widget)

    def _navigate_panes(self, direction: str):
        """Navigate panes in current tab"""
        container = self._get_current_container()
        if container:
            container.navigate_panes(direction)

    def _split_current_pane(self, direction: str):
        """Split the current active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container._split_pane(container.active_pane, direction)

    def _close_current_pane(self):
        """Close the current active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container._close_pane(container.active_pane)

    def _clear_current_terminal(self):
        """Clear the current active terminal pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container.active_pane._clear_terminal()

    def _toggle_connection(self):
        """Toggle connection state of active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            pane = container.active_pane
            if pane.is_connected:
                pane.disconnect()
            else:
                pane.connect()
            self._update_ribbon_connection_state()

    def _update_ribbon_connection_state(self):
        """Update ribbon toolbar connection button based on active pane state"""
        container = self._get_current_container()
        if container and container.active_pane:
            is_connected = container.active_pane.is_connected
            self.ribbon.set_connection_state(is_connected)
        else:
            self.ribbon.set_connection_state(False)

    def _on_active_pane_changed(self, pane):
        """Handle active pane change"""
        self._update_status_bar()
        self._update_ribbon_connection_state()

    def _update_status_bar(self):
        """Update status bar with active pane info"""
        container = self._get_current_container()
        if container and container.active_pane:
            status = container.active_pane.get_status_info()
            self.status_bar.showMessage(status)
        else:
            self.status_bar.showMessage("No active connection")

    def closeEvent(self, event):
        """Handle window close with cleanup"""
        try:
            self.status_timer.stop()

            for container in self.tabs.values():
                container.cleanup()

            event.accept()
        except Exception as e:
            print(f"Error during close: {e}")
            event.accept()

    def _show_settings_menu(self):
        """Show settings menu for the active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            active_pane = container.active_pane
            button_global_pos = self.ribbon.settings_button.mapToGlobal(
                self.ribbon.settings_button.rect().bottomLeft()
            )
            menu = active_pane._create_terminal_menu()
            menu.exec(button_global_pos)
        else:
            menu = QMenu(self)
            no_connection = menu.addAction("No active connection")
            no_connection.setEnabled(False)

            button_global_pos = self.ribbon.settings_button.mapToGlobal(
                self.ribbon.settings_button.rect().bottomLeft()
            )
            menu.exec(button_global_pos)


# ===== MAIN ENTRY POINT =====
def main():
    """Main application entry point for testing"""
    app = QApplication(sys.argv)

    app.setApplicationName("Network Terminal")
    app.setOrganizationName("NetworkTerminal")

    app.setStyle("Fusion")

    window = NetworkMonitorWindow()
    window.resize(1200, 800)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
