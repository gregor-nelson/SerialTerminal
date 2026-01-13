# Architecture Analysis: Serial Terminal Codebase

[← Back to Master PRD](../PRD_NetworkTerminal.md)

---

## Overview

This document provides a detailed breakdown of the Serial Terminal codebase to guide Network Terminal development. Understanding the existing architecture enables maximum code reuse and consistent design patterns.

---

## Codebase Statistics

| File | Lines | Purpose | Reusability |
|------|-------|---------|-------------|
| `core/core.py` | 1,747 | Workers, scanners, utilities | Partial |
| `ui/dialogs/terminal_dialog.py` | 2,598 | Main UI components | High |
| `ui/windows/terminal_formatter.py` | 344 | Display formatting | 100% |
| `ui/components/ribbon_toolbar.py` | ~150 | Toolbar component | 100% |
| `ui/common/icons.py` | ~100 | Icon definitions | 100% |
| `ui/resources.py` | ~200 | Resource management | 100% |
| `core/serial_config.py` | ~150 | Serial configuration | Template |
| `constants.py` | ~50 | App constants | Adapt |

---

## Core Components

### 1. core/core.py

This file contains the core functionality organized into several major sections:

#### Data Classes (Lines 48-143)

```python
@dataclass
class WindowConfig:
    """Configuration for window sizing and layout"""
    width: int
    height: int
    x: int
    y: int
    is_small_screen: bool
    min_width: int = 800
    min_height: int = 600

@dataclass
class SerialPortInfo:
    """Information about a detected serial port"""
    port_name: str
    device_name: str
    port_type: str
    registry_key: str
    description: str = ""
    is_moxa: bool = False
    # ... additional fields
```

**Reusability**: `WindowConfig` is 100% reusable. `SerialPortInfo` will be replaced by `NetworkConfig`.

#### SettingsManager (Lines 614-628)

```python
class SettingsManager:
    """Manages application settings using QSettings"""

    def __init__(self):
        self.settings = QSettings("SerialSplit", "Hub4com")

    def get_show_launch_dialog(self) -> bool:
        return self.settings.value("ui/show_launch_dialog", True, type=bool)

    def set_show_launch_dialog(self, show_dialog: bool):
        self.settings.setValue("ui/show_launch_dialog", show_dialog)
        self.settings.sync()
```

**Reusability**: Copy with namespace change to "NetworkTerminal".

#### ResponsiveWindowManager (Lines 630-760)

```python
class ResponsiveWindowManager:
    """Manages responsive window sizing and layout decisions"""

    SMALL_SCREEN_WIDTH_THRESHOLD = 1024
    SMALL_SCREEN_HEIGHT_THRESHOLD = 768
    # ... threshold constants

    @classmethod
    def calculate_main_window_config(cls) -> WindowConfig:
        """Calculate optimal window configuration"""
        # Screen detection and sizing logic
```

**Reusability**: 100% reusable without modification.

#### PortScanner (Lines 762-1032)

```python
class PortScanner(QThread):
    """Thread for scanning Windows registry for serial ports"""
    scan_completed = pyqtSignal(list)
    scan_progress = pyqtSignal(str)
    # ... signals for progressive loading
```

**Reusability**: NOT reusable - replaced by network socket discovery (if needed).

#### SerialPortMonitor (Lines 1151-1599)

```python
class SerialPortMonitor(QThread):
    """Serial port monitoring class for real-time statistics"""
    stats_updated = pyqtSignal(dict)
    data_received = pyqtSignal(bytes)
    error_occurred = pyqtSignal(str)
```

**Reusability**: NOT reusable - serial-specific monitoring.

---

### 2. ui/dialogs/terminal_dialog.py

The main UI file organized into three major classes:

#### SerialWorker (Lines 55-169)

```python
class SerialWorker(QThread):
    """Background thread for serial communication"""

    dataReceived = pyqtSignal(bytes)
    errorOccurred = pyqtSignal(str)
    connectionStateChanged = pyqtSignal(bool)

    def __init__(self, config: SerialConfig):
        super().__init__()
        self.config = config
        self.serial_port = None
        self.running = False
        self.write_queue = queue.Queue()
```

**Reusability**: REPLACE with `NetworkWorker` - but keep same signal interface!

**Critical Interface**:
```python
# These signals MUST be preserved in NetworkWorker
dataReceived = pyqtSignal(bytes)      # Data from connection
errorOccurred = pyqtSignal(str)       # Error messages
connectionStateChanged = pyqtSignal(bool)  # Connected/disconnected
```

#### TerminalPane (Lines 172-1399)

```python
class TerminalPane(QWidget):
    """Individual terminal display with formatter integration"""

    focusChanged = pyqtSignal(bool)
    splitRequested = pyqtSignal(object, str)
    closeRequested = pyqtSignal(object)

    def __init__(self, config: SerialConfig, parent=None,
                 main_window=None, container=None):
        # ... initialization
```

**Reusability**: HIGH - adapt for NetworkConfig, remove baud rate logic.

**Key Methods to Copy**:
- `_setup_ui()` - Terminal widget setup
- `_setup_context_menu()` - Right-click menu
- `eventFilter()` - Keyboard handling
- `_handle_key_press()` - Input processing
- `_flush_buffer()` - Line buffering
- `_format_bytes()` - Byte count formatting
- `_show_help()` / `_dismiss_help()` - Help display
- All font size methods
- All scroll methods

**Key Methods to Remove**:
- `_create_baud_rate_menu()` - Serial-specific
- `_create_com_port_menu()` - Serial-specific
- `_set_baud_rate()` - Serial-specific
- `_handle_encoding_error()` - Baud rate detection
- `_show_baud_rate_suggestion()` - Serial-specific
- `_show_virtual_port_manager()` - Serial-specific

**Key Methods to Modify**:
- `connect()` - Use NetworkWorker
- `_on_data_received()` - Remove baud rate handling
- `get_status_info()` - Network format
- `_create_terminal_menu()` - Remove serial options

#### SplitContainer (Search in file)

```python
class SplitContainer(QWidget):
    """Container for split terminal panes"""

    activePaneChanged = pyqtSignal(object)
    paneCountChanged = pyqtSignal(int)

    def __init__(self, initial_pane, parent=None):
        # Manages recursive splitter layout
```

**Reusability**: 100% reusable - handles split pane layout generically.

#### SerialMonitorWindow (Search in file)

```python
class SerialMonitorWindow(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serial Monitor")
        # ... window setup
```

**Reusability**: HIGH - adapt to `NetworkMonitorWindow`.

**Key Methods to Copy**:
- `_setup_ui()` - Main layout
- `_setup_shortcuts()` - Keyboard shortcuts
- `_setup_status_bar()` - Status bar
- `_create_new_tab()` - Tab management
- `_close_tab()` - Tab closing
- `_navigate_panes()` - Pane navigation
- `closeEvent()` - Cleanup

**Key Methods to Remove**:
- `get_connected_ports()` - Serial-specific

**Key Methods to Modify**:
- `_on_new_connection()` - Use QuickConnectDialog
- Window title

---

### 3. ui/windows/terminal_formatter.py

Complete file analysis:

```python
class TerminalStreamFormatter:
    """Formats terminal stream data with color-coded output"""

    def __init__(self):
        self.auto_scroll_enabled = True

        # Terminal color scheme
        self.colors = {
            'incoming': TerminalColors.INCOMING,
            'outgoing': TerminalColors.OUTGOING,
            'timestamp': TerminalColors.TIMESTAMP,
            # ... more colors
        }

        # NMEA message colors (27 types)
        self.nmea_colors = {
            'GGA': '#8bb5d9',
            'GLL': '#7ba8cc',
            # ... more NMEA types
        }
```

**Reusability**: 100% reusable - handles all display formatting.

**Key Methods**:
- `append_data()` - Add data with formatting
- `append_separator()` - Add visual separator
- `append_status()` - Add status message
- `clear()` - Clear terminal
- `format_connection_start()` - Connection message
- `format_connection_end()` - Disconnection message
- `_detect_nmea_message_type()` - NMEA color coding
- Auto-scroll management methods

---

## Component Relationships

```
┌─────────────────────────────────────────────────────────────────┐
│                     SerialMonitorWindow                          │
│                    (Main Application Window)                     │
├─────────────────────────────────────────────────────────────────┤
│ ┌─────────────┐  ┌──────────────────────────────────────────┐  │
│ │ RibbonToolbar│  │              Tab Widget                  │  │
│ │             │  ├──────────────────────────────────────────┤  │
│ │ • New       │  │  Tab 1: SplitContainer                   │  │
│ │ • Connect   │  │  ┌─────────────┬─────────────┐           │  │
│ │ • Clear     │  │  │ TerminalPane│ TerminalPane│           │  │
│ │ • Settings  │  │  │ ┌─────────┐ │ ┌─────────┐ │           │  │
│ │             │  │  │ │Terminal │ │ │Terminal │ │           │  │
│ └─────────────┘  │  │ │Formatter│ │ │Formatter│ │           │  │
│                  │  │ └─────────┘ │ └─────────┘ │           │  │
│                  │  │      │      │      │      │           │  │
│                  │  │ SerialWorker│ SerialWorker│           │  │
│                  │  └─────────────┴─────────────┘           │  │
│                  └──────────────────────────────────────────┘  │
├─────────────────────────────────────────────────────────────────┤
│                         Status Bar                               │
└─────────────────────────────────────────────────────────────────┘
```

---

## Signal Flow

### Data Reception Flow

```
SerialWorker.run()
    │
    ▼ (reads from port)
serial_port.read()
    │
    ▼ (emits signal)
dataReceived.emit(bytes)
    │
    ▼ (Qt queued connection - thread safe)
TerminalPane._on_data_received(bytes)
    │
    ▼ (decodes and buffers)
self.line_buffer += text_data
    │
    ▼ (formats output)
TerminalStreamFormatter.append_data()
    │
    ▼ (updates display)
QTextEdit.insertText()
```

### Connection State Flow

```
User clicks "Connect"
    │
    ▼
TerminalPane.connect()
    │
    ▼
SerialWorker.__init__(config)
SerialWorker.start()
    │
    ▼ (in thread)
serial.Serial(port=..., baudrate=...)
    │
    ▼ (success)
connectionStateChanged.emit(True)
    │
    ▼
TerminalPane._on_connection_state_changed(True)
    │
    ▼
formatter.format_connection_start()
```

---

## Threading Model

### Current Model (Serial)

```
┌─────────────────┐
│   Main Thread   │
│   (Qt GUI)      │
│                 │
│ • Event loop    │
│ • UI updates    │
│ • Signal slots  │
└────────┬────────┘
         │
         │ Qt Signals (QueuedConnection)
         │
┌────────▼────────┐
│ SerialWorker    │
│ (QThread)       │
│                 │
│ • Read loop     │
│ • Write queue   │
│ • Error handling│
└─────────────────┘
```

### Network Model (Same Pattern)

```
┌─────────────────┐
│   Main Thread   │
│   (Qt GUI)      │
│                 │
│ • Event loop    │
│ • UI updates    │
│ • Signal slots  │
└────────┬────────┘
         │
         │ Qt Signals (QueuedConnection)
         │
┌────────▼────────┐
│ NetworkWorker   │
│ (QThread)       │
│                 │
│ • Socket ops    │
│ • Write queue   │
│ • Error handling│
└─────────────────┘
```

---

## Error Handling Patterns

### Serial Error Handling

```python
def run(self):
    try:
        self.serial_port = serial.Serial(...)
        self.connectionStateChanged.emit(True)

        while self.running:
            # ... communication loop

    except serial.SerialException as e:
        self.errorOccurred.emit(str(e))
    finally:
        if self.serial_port:
            self.serial_port.close()
        self.connectionStateChanged.emit(False)
```

### Network Error Handling (Follow Same Pattern)

```python
def run(self):
    try:
        self.socket = socket.socket(...)
        self.socket.connect((host, port))
        self.connectionStateChanged.emit(True)

        while self.running:
            # ... communication loop

    except socket.timeout:
        self.errorOccurred.emit("Connection timeout")
    except ConnectionRefusedError:
        self.errorOccurred.emit("Connection refused")
    except OSError as e:
        self.errorOccurred.emit(f"Network error: {e}")
    finally:
        if self.socket:
            self.socket.close()
        self.connectionStateChanged.emit(False)
```

---

## Configuration Pattern

### Serial Configuration

```python
@dataclass
class SerialConfig:
    port: str
    baudrate: int = 115200
    databits: int = 8
    parity: str = 'N'
    stopbits: float = 1

    def get_display_string(self) -> str:
        return f"{self.baudrate} {self.databits}{self.parity}{int(self.stopbits)}"
```

### Network Configuration (Mirror Pattern)

```python
@dataclass
class NetworkConfig:
    host: str
    port: int
    protocol: str = 'TCP'
    mode: str = 'client'

    def get_display_string(self) -> str:
        return f"{self.protocol} | {self.host}:{self.port}"
```

---

## Key Design Principles

### 1. Signal-Based Communication

All cross-thread communication uses Qt signals with `QueuedConnection`:

```python
self.worker.dataReceived.connect(
    self._on_data_received,
    Qt.ConnectionType.QueuedConnection
)
```

### 2. Separation of Concerns

- **Worker**: Handles I/O in background thread
- **Pane**: Manages UI and user interaction
- **Formatter**: Handles all text formatting
- **Config**: Data transfer objects

### 3. Resource Cleanup

Every component has explicit cleanup:

```python
def cleanup(self):
    """Single point of cleanup"""
    self.is_connected = False

    if self.worker:
        self.worker.blockSignals(True)
        self.worker.stop()
        self.worker.wait(5000)
        self.worker = None

    self.line_buffer = ""
```

### 4. Defensive Programming

```python
def _on_data_received(self, data: bytes):
    # Guard against signals after cleanup
    if not self.worker:
        return

    try:
        # ... processing
    except Exception as e:
        self.formatter.append_status(
            self.terminal, f"Error: {e}", "error"
        )
```

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Reusability Matrix](./reusability-matrix.md)
- [Interface Definitions](./interface-definitions.md)
- [File Structure](./file-structure.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md)
