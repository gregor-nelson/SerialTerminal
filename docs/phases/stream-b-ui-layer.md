# Stream B: UI Layer

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Phase 0](./phase-0-project-setup.md)

---

## Overview

| Attribute | Value |
|-----------|-------|
| **Stream** | B - UI Layer |
| **Owner** | Developer 2 |
| **Dependencies** | Phase 0 complete |
| **Parallel With** | Stream A, Stream C |
| **Estimated Effort** | 4 tasks, ~200 new + ~400 adapted lines |
| **Status** | Not Started |

---

## Purpose

Adapt the Serial Terminal UI components for network connections. This involves:
1. Extracting reusable components (SplitContainer)
2. Adapting TerminalPane for NetworkConfig
3. Creating new connection dialogs
4. Adapting the main window

---

## Source File Reference

All UI adaptations are based on:
- **Source**: `SerialTerminal/ui/dialogs/terminal_dialog.py` (2,598 lines)
- **Key Classes**:
  - `SerialWorker` (lines 55-169) → Replace with NetworkWorker import
  - `TerminalPane` (lines 172-1399) → Adapt to `NetworkTerminalPane`
  - `SplitContainer` (search for class) → Copy unchanged
  - `SerialMonitorWindow` (search for class) → Adapt to `NetworkMonitorWindow`

---

## Task Checklist

- [ ] [Task B.1: Extract SplitContainer](#task-b1-extract-splitcontainer)
- [ ] [Task B.2: Create NetworkTerminalPane](#task-b2-create-networktarminalpane)
- [ ] [Task B.3: Create Connection Dialog](#task-b3-create-connection-dialog)
- [ ] [Task B.4: Create NetworkMonitorWindow](#task-b4-create-networkmonitorwindow)

---

## Task B.1: Extract SplitContainer

### Objective

Extract the `SplitContainer` class from Serial Terminal - it's 100% reusable.

### Source Location

`SerialTerminal/ui/dialogs/terminal_dialog.py` - search for `class SplitContainer`

### Instructions

1. Find `class SplitContainer` in the source file
2. Copy the entire class including all methods
3. Place in `NetworkTerminal/ui/dialogs/terminal_dialog.py`

### What SplitContainer Does

- Manages recursive split pane layout using QSplitter
- Tracks active pane
- Handles split/close operations
- Emits signals for pane changes

### Key Signals (preserve these)

```python
class SplitContainer(QWidget):
    activePaneChanged = pyqtSignal(object)  # Emits active TerminalPane
    paneCountChanged = pyqtSignal(int)       # Emits pane count
```

### Acceptance Criteria

- [ ] SplitContainer class copied to new project
- [ ] No modifications needed
- [ ] All methods preserved
- [ ] Signal definitions unchanged

---

## Task B.2: Create NetworkTerminalPane

### Objective

Adapt `TerminalPane` to work with `NetworkConfig` and `NetworkWorker`.

### File

`ui/dialogs/terminal_dialog.py`

### Changes Summary

| Section | Action |
|---------|--------|
| Imports | Replace SerialWorker with NetworkWorker imports |
| `__init__` | Change config type, remove baud rate attributes |
| Connection | Use `create_network_worker()` factory |
| Context Menu | Remove baud rate menu, add protocol options |
| Status | Change display format |

### Code Structure

```python
#!/usr/bin/env python3
"""
Network Terminal Dialog
Adapted from Serial Terminal for TCP/UDP connections
"""

from typing import Optional
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *

from ui.windows.terminal_formatter import TerminalStreamFormatter
from ui.resources import resource_manager
from core.network_config import NetworkConfig
from core.network_worker import NetworkWorker, create_network_worker


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

        # Line buffering (KEEP from original)
        self.line_buffer = ""
        self.buffer_timer = QTimer()
        self.buffer_timer.setSingleShot(True)
        self.buffer_timer.timeout.connect(self._flush_buffer)

        # Display settings (KEEP from original)
        self.encoding = 'utf-8'
        self.hex_display_mode = False
        self.local_echo_enabled = True

        # Help display (KEEP from original)
        self.help_displayed = False
        self.auto_scroll_state_before_help = True

        # REMOVED: Baud rate detection attributes
        # - encoding_error_count
        # - encoding_error_window
        # - encoding_error_threshold
        # - suggested_baud_rates
        # - baud_rate_suggestion_shown
        # - consecutive_errors

        self._setup_ui()
        self._setup_context_menu()
```

### Methods to COPY UNCHANGED from TerminalPane

```python
# Copy these methods exactly as they are:
def _setup_ui(self): ...
def _setup_context_menu(self): ...
def eventFilter(self, obj, event): ...
def _handle_key_press(self, event): ...
def _send_raw_data(self, data: str): ...
def _echo_local_data(self, data: str): ...
def _flush_buffer(self): ...
def _toggle_auto_scroll(self, enabled: bool): ...
def _toggle_hex_mode(self, enabled: bool): ...
def _toggle_local_echo(self, enabled: bool): ...
def _scroll_to_bottom(self): ...
def _clear_terminal(self): ...
def _set_font_size(self, size: int): ...
def _increase_font_size(self): ...
def _decrease_font_size(self): ...
def _reset_font_size(self): ...
def _show_help(self): ...
def _dismiss_help(self): ...
def _create_font_size_menu(self, menu: QMenu): ...
def checkbox_icon(self, checked: bool) -> QIcon: ...
def _format_bytes(self, bytes_count: int) -> str: ...
```

### Methods to MODIFY

#### connect() - Use NetworkWorker

```python
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
```

#### _on_data_received() - Remove baud rate handling

```python
def _on_data_received(self, data: bytes):
    """Handle received data - simplified from serial version"""
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
                self.terminal, formatted_data, "incoming", show_timestamp=True
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
        for line in lines[:-1]:
            if line:
                self.formatter.append_data(
                    self.terminal, line, "incoming", show_timestamp=True
                )

        # Keep incomplete line
        self.line_buffer = lines[-1]

        # Start buffer timer
        if self.line_buffer:
            self.buffer_timer.stop()
            self.buffer_timer.start(1000)

    except Exception as e:
        self.formatter.append_status(
            self.terminal, f"Data processing error: {e}", "error"
        )
```

#### get_status_info() - Network format

```python
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
```

#### _create_terminal_menu() - Remove serial options, add network options

```python
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

    # REMOVED: Baud rate submenu
    # REMOVED: COM port submenu

    # Connection settings submenu (NEW)
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
    """Create connection settings submenu"""
    # Show current connection info
    info = menu.addAction(f"{self.config.protocol} {self.config.mode.title()}")
    info.setEnabled(False)

    menu.addSeparator()

    # Protocol info
    host_action = menu.addAction(f"Host: {self.config.host}")
    host_action.setEnabled(False)

    port_action = menu.addAction(f"Port: {self.config.port}")
    port_action.setEnabled(False)
```

### Methods to REMOVE

```python
# DELETE these methods entirely:
def _create_baud_rate_menu(self, menu: QMenu): ...
def _create_com_port_menu(self, menu: QMenu): ...
def _create_com_port_display_text(self, port): ...
def _set_baud_rate(self, baud_rate: int): ...
def _complete_baud_rate_change(self, ...): ...
def _set_com_port(self, new_port: str): ...
def _complete_com_port_change(self, ...): ...
def _handle_encoding_error(self): ...
def _show_baud_rate_suggestion(self, error_rate: float): ...
def reset_baud_rate_detection(self): ...
def _handle_excessive_errors(self): ...
def _is_data_garbled(self, text_data: str) -> bool: ...
def _show_virtual_port_manager(self): ...
```

### Acceptance Criteria

- [ ] NetworkTerminalPane created with NetworkConfig
- [ ] Uses create_network_worker() for connections
- [ ] Baud rate logic completely removed
- [ ] Context menu adapted for network
- [ ] Status bar shows network format
- [ ] All display features work (hex, colors, scroll)

---

## Task B.3: Create Connection Dialog

### Objective

Create a dialog for configuring new network connections.

### File

`ui/dialogs/connection_dialog.py` (NEW FILE)

### Code

```python
#!/usr/bin/env python3
"""
Connection Dialog - Configure network connections
"""

from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *

from core.network_config import NetworkConfig


class QuickConnectDialog(QDialog):
    """
    Quick connect dialog for network connections.

    Allows user to configure:
    - Protocol (TCP/UDP)
    - Mode (Client/Server)
    - Host address
    - Port number
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
    """

    configSelected = pyqtSignal(object)  # Emits NetworkConfig

    def __init__(self, connection_manager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Connection History")
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
```

### Acceptance Criteria

- [ ] QuickConnectDialog shows protocol/mode/host/port options
- [ ] Server mode disables host field
- [ ] TCP/UDP toggle shows appropriate options
- [ ] get_config() returns valid NetworkConfig
- [ ] ConnectionHistoryDialog shows favorites and history

---

## Task B.4: Create NetworkMonitorWindow

### Objective

Adapt `SerialMonitorWindow` to `NetworkMonitorWindow` for the main application window.

### File

`ui/dialogs/terminal_dialog.py` (append after NetworkTerminalPane)

### Changes from SerialMonitorWindow

| Original | Change |
|----------|--------|
| Window title | "Network Terminal" |
| New connection | Opens QuickConnectDialog |
| get_connected_ports() | REMOVE - not applicable |
| Virtual Port Manager | REMOVE - not applicable |
| Status bar format | Network format |

### Code Structure

```python
class NetworkMonitorWindow(QMainWindow):
    """
    Main application window for Network Terminal.

    Adapted from SerialMonitorWindow with:
    - Network connection dialogs
    - Removed serial-specific features
    - Same split-pane and tab management
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Terminal")

        # Initialize managers
        from core.core import ResponsiveWindowManager
        window_config = ResponsiveWindowManager.calculate_main_window_config()

        self.setGeometry(
            window_config.x, window_config.y,
            window_config.width, window_config.height
        )
        self.setMinimumSize(window_config.min_width, window_config.min_height)

        self._setup_ui()
        self._setup_shortcuts()
        self._setup_status_bar()

    def _setup_ui(self):
        """Setup main UI components"""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Ribbon toolbar
        from ui.components.ribbon_toolbar import RibbonToolbar
        self.ribbon = RibbonToolbar()
        self.ribbon.newClicked.connect(self._on_new_connection)
        self.ribbon.refreshClicked.connect(self._on_refresh)
        self.ribbon.connectClicked.connect(self._on_connect_toggle)
        self.ribbon.clearClicked.connect(self._on_clear)
        self.ribbon.settingsClicked.connect(self._on_settings)
        layout.addWidget(self.ribbon)

        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self._close_tab)
        self.tab_widget.currentChanged.connect(self._on_tab_changed)
        layout.addWidget(self.tab_widget)

        # Create initial tab
        self._create_new_tab()

    def _on_new_connection(self):
        """Handle new connection button"""
        dialog = QuickConnectDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            config = dialog.get_config()
            self._add_pane_with_config(config)

    def _add_pane_with_config(self, config: NetworkConfig):
        """Add a new pane with the given config"""
        # Get current container or create new tab
        container = self._get_active_container()
        if container:
            pane = NetworkTerminalPane(
                config,
                main_window=self,
                container=container
            )
            container.add_pane(pane)
            pane.connect()

    # ... (Copy remaining methods from SerialMonitorWindow, adapting as needed)
```

### Methods to Copy/Adapt

Copy these methods with minimal changes:
- `_setup_shortcuts()` - keyboard shortcuts
- `_setup_status_bar()` - status bar setup
- `_create_new_tab()` - tab creation
- `_close_tab()` - tab closing
- `_on_tab_changed()` - tab switch handling
- `_get_active_container()` - get current split container
- `_get_active_pane()` - get current terminal pane
- `_update_status_bar()` - status bar updates
- `_on_clear()` - clear terminal
- `_on_connect_toggle()` - connect/disconnect
- `_navigate_panes()` - pane navigation
- `closeEvent()` - cleanup on close

### Methods to Remove

- `get_connected_ports()` - not applicable for network
- Any references to virtual port manager

### Acceptance Criteria

- [ ] Window opens with correct title
- [ ] New connection shows QuickConnectDialog
- [ ] Tabs work correctly
- [ ] Split panes work correctly
- [ ] Status bar shows network info
- [ ] Keyboard shortcuts work
- [ ] Clean shutdown on close

---

## Stream Deliverables

After completing Stream B, the following should be ready:

| Item | File | Status |
|------|------|--------|
| SplitContainer (copied) | `ui/dialogs/terminal_dialog.py` | |
| NetworkTerminalPane | `ui/dialogs/terminal_dialog.py` | |
| QuickConnectDialog | `ui/dialogs/connection_dialog.py` | |
| ConnectionHistoryDialog | `ui/dialogs/connection_dialog.py` | |
| NetworkMonitorWindow | `ui/dialogs/terminal_dialog.py` | |

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Architecture Analysis](../reference/architecture-analysis.md) - Original code structure
- [Interface Definitions](../reference/interface-definitions.md) - Signal contracts
- [Stream A: Core Network](./stream-a-core-network.md) - NetworkWorker classes
- [Stream C: Config Layer](./stream-c-config-data.md) - NetworkConfig class

---

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Stream A](./stream-a-core-network.md) | [Stream C →](./stream-c-config-data.md)
