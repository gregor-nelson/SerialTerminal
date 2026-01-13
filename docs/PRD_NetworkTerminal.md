# Product Requirements Document: Network Terminal Application

## TCP/UDP Terminal - Derived from Serial Terminal Codebase

**Document Version:** 1.0
**Date:** 2026-01-13
**Status:** Draft
**Parent Project:** SerialTerminal

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Current Architecture Analysis](#2-current-architecture-analysis)
3. [Reusability Assessment](#3-reusability-assessment)
4. [Network Terminal Requirements](#4-network-terminal-requirements)
5. [Implementation Phases](#5-implementation-phases)
6. [Interface Definitions](#6-interface-definitions)
7. [File Structure](#7-file-structure)
8. [Testing Requirements](#8-testing-requirements)
9. [Risk Assessment](#9-risk-assessment)

---

## 1. Executive Summary

### 1.1 Objective

Create a standalone **Network Terminal Application** for TCP and UDP connections that mirrors the design, layout, and feature set of the existing Serial Terminal application. The goal is to maximize code reuse while creating an independent, maintainable application.

### 1.2 Key Goals

- **85%+ code reuse** from existing Serial Terminal
- **Identical UI/UX** to Serial Terminal (Fusion theme, split panes, tabs)
- **Support both TCP and UDP** protocols with client/server modes
- **Standalone application** (no dependency on Serial Terminal at runtime)
- **Parallel development** capability via modular task breakdown

### 1.3 Scope

| In Scope | Out of Scope |
|----------|--------------|
| TCP Client connections | Serial port support |
| TCP Server (listen mode) | Virtual port management (COM0COM) |
| UDP Sender/Receiver | Baud rate detection |
| UDP Multicast support | Hardware flow control |
| Connection history/favorites | Moxa device integration |
| Split-pane terminal display | Windows registry scanning |
| NMEA message color-coding | |
| Hex display mode | |

---

## 2. Current Architecture Analysis

### 2.1 Project Structure Overview

```
SerialTerminal/
├── main.py                          # Entry point (95% reusable)
├── constants.py                     # Colors, app info (90% reusable)
├── build.py                         # Build script (80% reusable)
│
├── core/
│   ├── core.py                      # 1,747 lines - Core logic
│   │   ├── SerialWorker             # REPLACE → NetworkWorker
│   │   ├── SerialPortMonitor        # REPLACE → NetworkMonitor
│   │   ├── PortScanner              # REPLACE → ConnectionManager
│   │   ├── SettingsManager          # REUSE as-is
│   │   ├── ResponsiveWindowManager  # REUSE as-is
│   │   └── AdvancedStatistics       # ADAPT for network metrics
│   │
│   ├── serial_config.py             # REPLACE → network_config.py
│   └── com0com.py                   # NOT NEEDED
│
├── ui/
│   ├── dialogs/
│   │   ├── terminal_dialog.py       # 2,598 lines - Main UI
│   │   │   ├── SerialWorker         # REPLACE → NetworkWorker
│   │   │   ├── TerminalPane         # ADAPT (change config type)
│   │   │   ├── SplitContainer       # REUSE as-is
│   │   │   └── SerialMonitorWindow  # RENAME → NetworkMonitorWindow
│   │   │
│   │   └── virtual_port_dialog.py   # NOT NEEDED
│   │
│   ├── windows/
│   │   └── terminal_formatter.py    # 344 lines - REUSE 100%
│   │
│   ├── components/
│   │   └── ribbon_toolbar.py        # REUSE with minor text changes
│   │
│   ├── common/
│   │   └── icons.py                 # REUSE 100%
│   │
│   └── resources.py                 # REUSE 100%
│
└── assets/                          # REUSE 100%
    ├── fonts/
    └── icons/
```

### 2.2 Key Classes Analysis

#### 2.2.1 SerialWorker (terminal_dialog.py:55-169)

**Current Implementation:**
```python
class SerialWorker(QThread):
    dataReceived = pyqtSignal(bytes)
    errorOccurred = pyqtSignal(str)
    connectionStateChanged = pyqtSignal(bool)

    def __init__(self, config: SerialConfig):
        self.config = config
        self.serial_port: Optional[serial.Serial] = None
        self.write_queue = queue.Queue()

    def run(self):
        # Open serial port
        # Read/write loop with queue

    def stop(self):
        # Graceful shutdown

    def write(self, data: bytes):
        # Queue data for transmission
```

**Required Changes for Network:**
- Replace `serial.Serial` with `socket.socket`
- Add protocol selection (TCP/UDP)
- Add client/server mode support
- Different connection establishment logic
- Socket-specific error handling

#### 2.2.2 TerminalPane (terminal_dialog.py:172-1399)

**Current Implementation:**
- Manages one serial connection
- Handles data display via `TerminalStreamFormatter`
- Context menu for settings
- Baud rate switching (NOT NEEDED for network)
- COM port switching → becomes Host:Port switching

**Required Changes:**
- Change `SerialConfig` → `NetworkConfig`
- Replace COM port menu with host:port selection
- Remove baud rate detection logic
- Add protocol-specific options (TCP keepalive, UDP broadcast)

#### 2.2.3 TerminalStreamFormatter (terminal_formatter.py:1-345)

**100% Reusable** - This class is completely protocol-agnostic:
- Color-coded data streams (RX green, TX blue)
- NMEA message detection and coloring
- Timestamp formatting
- Hex display mode
- Auto-scroll management

#### 2.2.4 SplitContainer (terminal_dialog.py - referenced)

**100% Reusable** - Manages split pane layout independent of connection type.

### 2.3 Data Structures

#### Current: SerialConfig (serial_config.py)

```python
@dataclass
class SerialConfig:
    port: str
    baudrate: int = 115200
    databits: int = 8
    parity: str = 'N'
    stopbits: float = 1.0

    def get_display_string(self) -> str:
        return f"{self.baudrate} {self.databits}{self.parity}{self.stopbits}"
```

#### Required: NetworkConfig

```python
@dataclass
class NetworkConfig:
    host: str
    port: int
    protocol: str = 'TCP'  # 'TCP' or 'UDP'
    mode: str = 'client'   # 'client' or 'server'

    # TCP-specific
    keepalive: bool = True
    keepalive_interval: int = 60

    # UDP-specific
    broadcast: bool = False
    multicast_group: Optional[str] = None
    multicast_ttl: int = 1

    # Common
    buffer_size: int = 4096
    timeout: float = 5.0

    def get_display_string(self) -> str:
        mode_str = "Server" if self.mode == 'server' else "Client"
        return f"{self.protocol} {mode_str} - {self.host}:{self.port}"
```

---

## 3. Reusability Assessment

### 3.1 Reusability Matrix

| Component | File | Lines | Reuse % | Action |
|-----------|------|-------|---------|--------|
| `TerminalStreamFormatter` | terminal_formatter.py | 344 | **100%** | Copy unchanged |
| `SplitContainer` | terminal_dialog.py | ~200 | **100%** | Extract and copy |
| `RibbonToolbar` | ribbon_toolbar.py | ~150 | **95%** | Copy, change button labels |
| `Icons` | icons.py | ~100 | **100%** | Copy unchanged |
| `ResourceManager` | resources.py | ~200 | **100%** | Copy unchanged |
| `TerminalColors` | constants.py | 26 | **100%** | Copy unchanged |
| `ResponsiveWindowManager` | core.py | ~130 | **100%** | Extract and copy |
| `SettingsManager` | core.py | ~20 | **90%** | Copy, change app name |
| `TerminalPane` | terminal_dialog.py | ~1200 | **70%** | Adapt for network |
| `SerialMonitorWindow` | terminal_dialog.py | ~800 | **80%** | Rename, adapt menus |
| `SerialWorker` | terminal_dialog.py | ~115 | **30%** | Replace with NetworkWorker |
| `SerialConfig` | serial_config.py | 21 | **20%** | Replace with NetworkConfig |
| `PortScanner` | core.py | ~270 | **0%** | Replace with ConnectionManager |

### 3.2 Estimated Development Effort

| Category | Lines to Write | Lines to Adapt | Lines to Copy |
|----------|---------------|----------------|---------------|
| Core Network Logic | ~400 | ~200 | ~150 |
| UI Adaptation | ~100 | ~300 | ~2000 |
| Configuration | ~80 | ~20 | ~0 |
| Entry Point | ~20 | ~50 | ~100 |
| **Total** | **~600** | **~570** | **~2250** |

**Summary:** ~600 new lines, ~570 adapted lines, ~2250 copied lines = ~3420 total

---

## 4. Network Terminal Requirements

### 4.1 Functional Requirements

#### 4.1.1 TCP Client Mode

| ID | Requirement | Priority |
|----|-------------|----------|
| TCP-C-01 | Connect to remote host:port | P0 |
| TCP-C-02 | Automatic reconnection on disconnect | P1 |
| TCP-C-03 | Configurable connection timeout | P1 |
| TCP-C-04 | TCP keepalive support | P1 |
| TCP-C-05 | Display connection state in status bar | P0 |

#### 4.1.2 TCP Server Mode

| ID | Requirement | Priority |
|----|-------------|----------|
| TCP-S-01 | Listen on specified port | P0 |
| TCP-S-02 | Accept single client connection | P0 |
| TCP-S-03 | Display client address on connect | P0 |
| TCP-S-04 | Graceful client disconnect handling | P1 |
| TCP-S-05 | Optional: Accept multiple clients (future) | P2 |

#### 4.1.3 UDP Mode

| ID | Requirement | Priority |
|----|-------------|----------|
| UDP-01 | Send datagrams to host:port | P0 |
| UDP-02 | Receive datagrams on bound port | P0 |
| UDP-03 | UDP broadcast support | P1 |
| UDP-04 | UDP multicast support | P2 |
| UDP-05 | Display source address for received data | P1 |

#### 4.1.4 UI Requirements

| ID | Requirement | Priority |
|----|-------------|----------|
| UI-01 | Split-pane terminal display | P0 |
| UI-02 | Tab-based multiple connections | P0 |
| UI-03 | Hex display mode toggle | P0 |
| UI-04 | Auto-scroll with manual override | P0 |
| UI-05 | NMEA message color-coding | P0 |
| UI-06 | Connection favorites/history | P1 |
| UI-07 | Right-click context menu | P0 |
| UI-08 | Ribbon toolbar with 5 actions | P0 |
| UI-09 | Status bar with RX/TX statistics | P0 |

### 4.2 Non-Functional Requirements

| ID | Requirement | Target |
|----|-------------|--------|
| NFR-01 | Startup time | < 2 seconds |
| NFR-02 | Memory usage (idle) | < 100 MB |
| NFR-03 | CPU usage (idle) | < 1% |
| NFR-04 | Terminal buffer | 10,000 lines |
| NFR-05 | Supported OS | Windows 10/11 |
| NFR-06 | Python version | 3.8+ |

---

## 5. Implementation Phases

### Overview: Parallel Development Streams

The implementation is organized into **independent work streams** that can be developed in parallel by multiple developers. Each stream has clear interfaces and minimal dependencies.

```
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 0: Project Setup                       │
│                    (Single developer, 1 day)                    │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│   STREAM A    │    │   STREAM B    │    │   STREAM C    │
│  Core Network │    │   UI Layer    │    │  Config/Data  │
│   (Dev 1)     │    │   (Dev 2)     │    │   (Dev 3)     │
└───────────────┘    └───────────────┘    └───────────────┘
        │                     │                     │
        └─────────────────────┼─────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 4: Integration                         │
│                    (All developers, 2 days)                     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 5: Testing & Polish                    │
│                    (All developers, 2 days)                     │
└─────────────────────────────────────────────────────────────────┘
```

---

### PHASE 0: Project Setup (Prerequisite)

**Owner:** Any single developer
**Dependencies:** None
**Deliverables:** Project skeleton with copied reusable components

#### Task 0.1: Create Project Structure

```bash
NetworkTerminal/
├── main.py                    # Copy from SerialTerminal, modify app name
├── constants.py               # Copy from SerialTerminal
├── build.py                   # Copy from SerialTerminal, modify output name
├── requirements.txt           # New file (PyQt6 only, no pyserial)
│
├── core/
│   ├── __init__.py
│   ├── network_config.py      # NEW - placeholder
│   ├── network_worker.py      # NEW - placeholder
│   └── core.py                # Extract: SettingsManager, ResponsiveWindowManager
│
├── ui/
│   ├── __init__.py
│   ├── dialogs/
│   │   ├── __init__.py
│   │   └── terminal_dialog.py # Placeholder - will be adapted
│   │
│   ├── windows/
│   │   ├── __init__.py
│   │   └── terminal_formatter.py  # COPY unchanged
│   │
│   ├── components/
│   │   ├── __init__.py
│   │   └── ribbon_toolbar.py  # COPY unchanged
│   │
│   ├── common/
│   │   ├── __init__.py
│   │   └── icons.py           # COPY unchanged
│   │
│   └── resources.py           # COPY unchanged
│
├── assets/                    # COPY entire directory
│   ├── fonts/
│   └── icons/
│
└── docs/
    └── PRD_NetworkTerminal.md # This document
```

#### Task 0.2: Copy Reusable Files (Direct Copy)

| Source | Destination | Notes |
|--------|-------------|-------|
| `SerialTerminal/constants.py` | `NetworkTerminal/constants.py` | Change AppInfo.NAME |
| `SerialTerminal/ui/windows/terminal_formatter.py` | `NetworkTerminal/ui/windows/terminal_formatter.py` | No changes |
| `SerialTerminal/ui/components/ribbon_toolbar.py` | `NetworkTerminal/ui/components/ribbon_toolbar.py` | No changes |
| `SerialTerminal/ui/common/icons.py` | `NetworkTerminal/ui/common/icons.py` | No changes |
| `SerialTerminal/ui/resources.py` | `NetworkTerminal/ui/resources.py` | No changes |
| `SerialTerminal/assets/*` | `NetworkTerminal/assets/*` | Copy entire directory |

#### Task 0.3: Create requirements.txt

```
PyQt6>=6.4.0
```

#### Task 0.4: Verify Skeleton Runs

Create minimal `main.py` that imports all copied modules and shows a blank window.

**Acceptance Criteria:**
- [ ] Project structure created
- [ ] All files copied
- [ ] `python main.py` shows blank PyQt6 window
- [ ] No import errors

---

### STREAM A: Core Network Layer (Parallel)

**Owner:** Developer 1
**Dependencies:** Phase 0 complete
**Files to Create:**
- `core/network_worker.py`
- `core/network_monitor.py` (optional)

#### Task A.1: Create NetworkWorker Base Class

**File:** `core/network_worker.py`

```python
"""
Network Worker - Background thread for TCP/UDP communication
Mirrors SerialWorker interface for drop-in replacement
"""

from PyQt6.QtCore import QThread, pyqtSignal
from typing import Optional
import socket
import queue
import threading

class NetworkWorker(QThread):
    """Background thread for network communication"""

    # Signals - IDENTICAL to SerialWorker for compatibility
    dataReceived = pyqtSignal(bytes)
    errorOccurred = pyqtSignal(str)
    connectionStateChanged = pyqtSignal(bool)

    # Additional signals for network-specific events
    clientConnected = pyqtSignal(str)      # Client address (server mode)
    clientDisconnected = pyqtSignal(str)   # Client address (server mode)

    def __init__(self, config: 'NetworkConfig'):
        super().__init__()
        self.config = config
        self.socket: Optional[socket.socket] = None
        self.client_socket: Optional[socket.socket] = None  # For server mode
        self.running = False
        self.write_queue = queue.Queue()
        self._stop_event = threading.Event()

    def run(self):
        """Main thread loop - protocol-specific implementation"""
        raise NotImplementedError("Subclasses must implement run()")

    def stop(self):
        """Stop the worker thread safely"""
        self._stop_event.set()
        self.running = False

        # Close sockets to interrupt blocking operations
        for sock in [self.client_socket, self.socket]:
            if sock:
                try:
                    sock.shutdown(socket.SHUT_RDWR)
                except Exception:
                    pass
                try:
                    sock.close()
                except Exception:
                    pass

        # Wait for thread to finish
        if self.isRunning():
            if not self.wait(5000):
                print(f"Warning: Network worker did not stop cleanly")

    def write(self, data: bytes):
        """Queue data to be written"""
        if self.running:
            self.write_queue.put(data)
```

#### Task A.2: Implement TCPClientWorker

**File:** `core/network_worker.py` (append)

```python
class TCPClientWorker(NetworkWorker):
    """TCP Client connection handler"""

    def run(self):
        try:
            # Create socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(self.config.timeout)

            # Configure keepalive if enabled
            if self.config.keepalive:
                self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
                # Platform-specific keepalive settings (Windows)
                if hasattr(socket, 'SIO_KEEPALIVE_VALS'):
                    self.socket.ioctl(socket.SIO_KEEPALIVE_VALS, (
                        1,  # Enable
                        self.config.keepalive_interval * 1000,  # Idle time (ms)
                        1000  # Interval (ms)
                    ))

            # Connect
            self.socket.connect((self.config.host, self.config.port))
            self.socket.setblocking(False)

            self.running = True
            self.connectionStateChanged.emit(True)

            # Main loop
            while self.running and not self._stop_event.is_set():
                # Read available data
                try:
                    data = self.socket.recv(self.config.buffer_size)
                    if data:
                        self.dataReceived.emit(data)
                    elif data == b'':
                        # Connection closed by remote
                        self.errorOccurred.emit("Connection closed by remote host")
                        break
                except BlockingIOError:
                    pass  # No data available
                except socket.timeout:
                    pass  # Timeout, continue loop

                # Write queued data
                try:
                    while not self.write_queue.empty() and self.running:
                        data = self.write_queue.get_nowait()
                        self.socket.sendall(data)
                except queue.Empty:
                    pass
                except socket.error as e:
                    self.errorOccurred.emit(f"Send error: {e}")

                # Small sleep to prevent CPU spinning
                self._stop_event.wait(0.01)

        except socket.timeout:
            self.errorOccurred.emit(f"Connection timeout to {self.config.host}:{self.config.port}")
        except ConnectionRefusedError:
            self.errorOccurred.emit(f"Connection refused by {self.config.host}:{self.config.port}")
        except socket.gaierror as e:
            self.errorOccurred.emit(f"DNS resolution failed: {e}")
        except Exception as e:
            self.errorOccurred.emit(f"Connection error: {e}")
        finally:
            if self.socket:
                try:
                    self.socket.close()
                except Exception:
                    pass
            self.connectionStateChanged.emit(False)
```

#### Task A.3: Implement TCPServerWorker

**File:** `core/network_worker.py` (append)

```python
class TCPServerWorker(NetworkWorker):
    """TCP Server (listen) mode handler"""

    def run(self):
        try:
            # Create server socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.socket.settimeout(1.0)  # For accept() timeout

            # Bind and listen
            self.socket.bind((self.config.host, self.config.port))
            self.socket.listen(1)

            self.running = True
            self.connectionStateChanged.emit(True)

            # Wait for client connection
            while self.running and not self._stop_event.is_set():
                try:
                    self.client_socket, client_addr = self.socket.accept()
                    self.client_socket.setblocking(False)
                    self.clientConnected.emit(f"{client_addr[0]}:{client_addr[1]}")

                    # Handle client
                    self._handle_client()

                    self.clientDisconnected.emit(f"{client_addr[0]}:{client_addr[1]}")

                except socket.timeout:
                    continue  # Keep waiting for connections

        except OSError as e:
            if "address already in use" in str(e).lower():
                self.errorOccurred.emit(f"Port {self.config.port} is already in use")
            else:
                self.errorOccurred.emit(f"Server error: {e}")
        except Exception as e:
            self.errorOccurred.emit(f"Server error: {e}")
        finally:
            for sock in [self.client_socket, self.socket]:
                if sock:
                    try:
                        sock.close()
                    except Exception:
                        pass
            self.connectionStateChanged.emit(False)

    def _handle_client(self):
        """Handle connected client communication"""
        while self.running and not self._stop_event.is_set():
            # Read from client
            try:
                data = self.client_socket.recv(self.config.buffer_size)
                if data:
                    self.dataReceived.emit(data)
                elif data == b'':
                    break  # Client disconnected
            except BlockingIOError:
                pass
            except socket.error:
                break

            # Write to client
            try:
                while not self.write_queue.empty() and self.running:
                    data = self.write_queue.get_nowait()
                    self.client_socket.sendall(data)
            except queue.Empty:
                pass
            except socket.error:
                break

            self._stop_event.wait(0.01)
```

#### Task A.4: Implement UDPWorker

**File:** `core/network_worker.py` (append)

```python
class UDPWorker(NetworkWorker):
    """UDP sender/receiver handler"""

    def run(self):
        try:
            # Create UDP socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self.socket.settimeout(0.1)

            # Enable broadcast if configured
            if self.config.broadcast:
                self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)

            # Bind for receiving (server mode) or any port (client mode)
            if self.config.mode == 'server':
                self.socket.bind((self.config.host, self.config.port))
            else:
                self.socket.bind(('', 0))  # Any available port

            # Join multicast group if configured
            if self.config.multicast_group:
                import struct
                mreq = struct.pack('4sl',
                    socket.inet_aton(self.config.multicast_group),
                    socket.INADDR_ANY)
                self.socket.setsockopt(socket.IPPROTO_IP,
                    socket.IP_ADD_MEMBERSHIP, mreq)

            self.running = True
            self.connectionStateChanged.emit(True)

            # Main loop
            while self.running and not self._stop_event.is_set():
                # Receive datagrams
                try:
                    data, addr = self.socket.recvfrom(self.config.buffer_size)
                    if data:
                        # Include source address in signal for UDP
                        self.dataReceived.emit(data)
                except socket.timeout:
                    pass

                # Send queued datagrams
                try:
                    while not self.write_queue.empty() and self.running:
                        data = self.write_queue.get_nowait()
                        target = (self.config.host, self.config.port)
                        self.socket.sendto(data, target)
                except queue.Empty:
                    pass
                except socket.error as e:
                    self.errorOccurred.emit(f"Send error: {e}")

                self._stop_event.wait(0.01)

        except Exception as e:
            self.errorOccurred.emit(f"UDP error: {e}")
        finally:
            if self.socket:
                try:
                    self.socket.close()
                except Exception:
                    pass
            self.connectionStateChanged.emit(False)
```

#### Task A.5: Create NetworkWorker Factory

**File:** `core/network_worker.py` (append)

```python
def create_network_worker(config: 'NetworkConfig') -> NetworkWorker:
    """Factory function to create appropriate worker based on config"""
    if config.protocol == 'TCP':
        if config.mode == 'server':
            return TCPServerWorker(config)
        else:
            return TCPClientWorker(config)
    elif config.protocol == 'UDP':
        return UDPWorker(config)
    else:
        raise ValueError(f"Unknown protocol: {config.protocol}")
```

**Acceptance Criteria for Stream A:**
- [ ] `NetworkWorker` base class with identical signals to `SerialWorker`
- [ ] `TCPClientWorker` connects to remote host
- [ ] `TCPServerWorker` accepts incoming connections
- [ ] `UDPWorker` sends/receives datagrams
- [ ] All workers use `write_queue` pattern
- [ ] Graceful shutdown via `stop()` method
- [ ] Unit tests for each worker type

---

### STREAM B: UI Layer (Parallel)

**Owner:** Developer 2
**Dependencies:** Phase 0 complete
**Files to Modify:**
- `ui/dialogs/terminal_dialog.py`
- `ui/components/ribbon_toolbar.py` (minor)

#### Task B.1: Extract and Copy SplitContainer

Extract `SplitContainer` class from Serial Terminal's `terminal_dialog.py` and copy to Network Terminal. This class is 100% reusable.

**Location in source:** `SerialTerminal/ui/dialogs/terminal_dialog.py` (search for `class SplitContainer`)

**No modifications needed** - copy as-is.

#### Task B.2: Adapt TerminalPane for Network

Create `NetworkTerminalPane` by adapting `TerminalPane`:

**Key Changes:**

```python
class NetworkTerminalPane(QWidget):
    """Individual terminal display for network connections"""

    # SAME signals as TerminalPane
    focusChanged = pyqtSignal(bool)
    splitRequested = pyqtSignal(object, str)
    closeRequested = pyqtSignal(object)

    def __init__(self, config: NetworkConfig, parent=None, main_window=None, container=None):
        # Change: SerialConfig → NetworkConfig
        self.config = config

        # Change: SerialWorker → NetworkWorker
        self.network_worker: Optional[NetworkWorker] = None

        # REMOVE: Baud rate detection attributes
        # - self.encoding_error_count
        # - self.suggested_baud_rates
        # - self.baud_rate_suggestion_shown
        # - etc.

        # KEEP: All display-related attributes
        # - self.formatter
        # - self.hex_display_mode
        # - self.local_echo_enabled
        # - self.line_buffer
        # - etc.
```

**Methods to REMOVE:**
- `_create_baud_rate_menu()`
- `_set_baud_rate()`
- `_complete_baud_rate_change()`
- `_handle_encoding_error()`
- `_show_baud_rate_suggestion()`
- `reset_baud_rate_detection()`
- `_handle_excessive_errors()`
- `_is_data_garbled()`

**Methods to MODIFY:**
- `_create_terminal_menu()` - Remove baud rate, add protocol options
- `_create_com_port_menu()` → `_create_connection_menu()`
- `connect()` - Use `create_network_worker()`
- `get_status_info()` - Change display format

**Methods to KEEP (unchanged):**
- `_setup_ui()`
- `_setup_context_menu()`
- `_on_data_received()` (remove baud rate error handling)
- `_on_error()`
- `_toggle_auto_scroll()`
- `_toggle_hex_mode()`
- `_toggle_local_echo()`
- `_clear_terminal()`
- `_set_font_size()`
- `_show_help()`
- All formatter integration methods

#### Task B.3: Create Connection Dialog

**New File:** `ui/dialogs/connection_dialog.py`

```python
"""
Connection Dialog - Quick connect to network endpoint
"""

from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from core.network_config import NetworkConfig


class QuickConnectDialog(QDialog):
    """Quick connect dialog for network connections"""

    def __init__(self, parent=None, initial_config: NetworkConfig = None):
        super().__init__(parent)
        self.setWindowTitle("Connect")
        self.config = initial_config or NetworkConfig(host="127.0.0.1", port=5000)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        # Protocol selection
        protocol_group = QGroupBox("Protocol")
        protocol_layout = QHBoxLayout(protocol_group)
        self.tcp_radio = QRadioButton("TCP")
        self.udp_radio = QRadioButton("UDP")
        self.tcp_radio.setChecked(self.config.protocol == 'TCP')
        self.udp_radio.setChecked(self.config.protocol == 'UDP')
        protocol_layout.addWidget(self.tcp_radio)
        protocol_layout.addWidget(self.udp_radio)
        layout.addWidget(protocol_group)

        # Mode selection
        mode_group = QGroupBox("Mode")
        mode_layout = QHBoxLayout(mode_group)
        self.client_radio = QRadioButton("Client")
        self.server_radio = QRadioButton("Server")
        self.client_radio.setChecked(self.config.mode == 'client')
        self.server_radio.setChecked(self.config.mode == 'server')
        mode_layout.addWidget(self.client_radio)
        mode_layout.addWidget(self.server_radio)
        layout.addWidget(mode_group)

        # Host input
        host_layout = QHBoxLayout()
        host_layout.addWidget(QLabel("Host:"))
        self.host_input = QLineEdit(self.config.host)
        self.host_input.setPlaceholderText("hostname or IP address")
        host_layout.addWidget(self.host_input)
        layout.addLayout(host_layout)

        # Port input
        port_layout = QHBoxLayout()
        port_layout.addWidget(QLabel("Port:"))
        self.port_input = QSpinBox()
        self.port_input.setRange(1, 65535)
        self.port_input.setValue(self.config.port)
        port_layout.addWidget(self.port_input)
        layout.addLayout(port_layout)

        # Buttons
        button_layout = QHBoxLayout()
        connect_btn = QPushButton("Connect")
        connect_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(connect_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        # Update host field enabled state based on mode
        self.server_radio.toggled.connect(self._on_mode_changed)
        self._on_mode_changed()

    def _on_mode_changed(self):
        """Update UI based on client/server mode"""
        is_server = self.server_radio.isChecked()
        self.host_input.setEnabled(not is_server)
        if is_server:
            self.host_input.setText("0.0.0.0")
        elif self.host_input.text() == "0.0.0.0":
            self.host_input.setText("127.0.0.1")

    def get_config(self) -> NetworkConfig:
        """Get the configured NetworkConfig"""
        return NetworkConfig(
            host=self.host_input.text(),
            port=self.port_input.value(),
            protocol='TCP' if self.tcp_radio.isChecked() else 'UDP',
            mode='server' if self.server_radio.isChecked() else 'client'
        )
```

#### Task B.4: Adapt Main Window

Adapt `SerialMonitorWindow` → `NetworkMonitorWindow`:

**Key Changes:**

```python
class NetworkMonitorWindow(QMainWindow):
    """Main application window for Network Terminal"""

    def __init__(self):
        # Change window title
        self.setWindowTitle("Network Terminal")

        # SAME: Tab widget, split container, status bar
        # SAME: Ribbon toolbar (button actions adapted)

    def _create_new_connection(self):
        """Show connection dialog and create new pane"""
        dialog = QuickConnectDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            config = dialog.get_config()
            self._add_pane_with_config(config)

    # REMOVE: get_connected_ports() - not applicable
    # REMOVE: _show_virtual_port_manager() - not applicable

    # KEEP: All split/tab management methods
    # KEEP: Status bar update methods
    # KEEP: Keyboard shortcut handlers
```

**Acceptance Criteria for Stream B:**
- [ ] `NetworkTerminalPane` displays data from `NetworkWorker`
- [ ] Context menu adapted for network options
- [ ] `QuickConnectDialog` allows protocol/mode/host/port selection
- [ ] `NetworkMonitorWindow` manages tabs and split panes
- [ ] All keyboard shortcuts work
- [ ] Status bar shows RX/TX/connection state

---

### STREAM C: Configuration & Data Layer (Parallel)

**Owner:** Developer 3
**Dependencies:** Phase 0 complete
**Files to Create:**
- `core/network_config.py`
- `core/connection_manager.py`

#### Task C.1: Create NetworkConfig

**File:** `core/network_config.py`

```python
"""
Network Configuration Data Model
"""

from dataclasses import dataclass, field
from typing import Optional, List
import json
from pathlib import Path


@dataclass
class NetworkConfig:
    """Network connection configuration"""

    # Required fields
    host: str
    port: int

    # Protocol settings
    protocol: str = 'TCP'      # 'TCP' or 'UDP'
    mode: str = 'client'       # 'client' or 'server'

    # TCP-specific options
    keepalive: bool = True
    keepalive_interval: int = 60  # seconds
    nodelay: bool = True          # TCP_NODELAY (disable Nagle)

    # UDP-specific options
    broadcast: bool = False
    multicast_group: Optional[str] = None
    multicast_ttl: int = 1

    # Common options
    buffer_size: int = 4096
    timeout: float = 5.0
    reconnect_on_disconnect: bool = False
    reconnect_delay: float = 3.0

    # Display name (for favorites)
    name: Optional[str] = None

    def get_display_string(self) -> str:
        """Get display string for status bar"""
        mode_str = "Listen" if self.mode == 'server' else ""
        if self.protocol == 'UDP':
            mode_str = "Broadcast" if self.broadcast else ""

        parts = [self.protocol]
        if mode_str:
            parts.append(mode_str)
        parts.append(f"{self.host}:{self.port}")

        return " | ".join(parts)

    def get_connection_id(self) -> str:
        """Get unique identifier for this connection"""
        return f"{self.protocol}_{self.mode}_{self.host}_{self.port}"

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization"""
        return {
            'host': self.host,
            'port': self.port,
            'protocol': self.protocol,
            'mode': self.mode,
            'keepalive': self.keepalive,
            'keepalive_interval': self.keepalive_interval,
            'nodelay': self.nodelay,
            'broadcast': self.broadcast,
            'multicast_group': self.multicast_group,
            'multicast_ttl': self.multicast_ttl,
            'buffer_size': self.buffer_size,
            'timeout': self.timeout,
            'reconnect_on_disconnect': self.reconnect_on_disconnect,
            'reconnect_delay': self.reconnect_delay,
            'name': self.name
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'NetworkConfig':
        """Create from dictionary"""
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def __post_init__(self):
        """Validate configuration after initialization"""
        if self.protocol not in ('TCP', 'UDP'):
            raise ValueError(f"Invalid protocol: {self.protocol}")
        if self.mode not in ('client', 'server'):
            raise ValueError(f"Invalid mode: {self.mode}")
        if not 1 <= self.port <= 65535:
            raise ValueError(f"Invalid port: {self.port}")
```

#### Task C.2: Create Connection Manager

**File:** `core/connection_manager.py`

```python
"""
Connection Manager - Handles favorites and connection history
"""

from dataclasses import dataclass, field
from typing import List, Optional
from datetime import datetime
import json
from pathlib import Path

from PyQt6.QtCore import QSettings

from .network_config import NetworkConfig


@dataclass
class ConnectionHistory:
    """Record of a past connection"""
    config: NetworkConfig
    last_connected: datetime
    connect_count: int = 1

    def to_dict(self) -> dict:
        return {
            'config': self.config.to_dict(),
            'last_connected': self.last_connected.isoformat(),
            'connect_count': self.connect_count
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ConnectionHistory':
        return cls(
            config=NetworkConfig.from_dict(data['config']),
            last_connected=datetime.fromisoformat(data['last_connected']),
            connect_count=data.get('connect_count', 1)
        )


class ConnectionManager:
    """Manages connection favorites and history"""

    MAX_HISTORY = 20

    def __init__(self):
        self.settings = QSettings("NetworkTerminal", "ConnectionManager")
        self._favorites: List[NetworkConfig] = []
        self._history: List[ConnectionHistory] = []
        self._load()

    def _load(self):
        """Load favorites and history from settings"""
        # Load favorites
        favorites_json = self.settings.value("favorites", "[]")
        try:
            favorites_data = json.loads(favorites_json)
            self._favorites = [NetworkConfig.from_dict(f) for f in favorites_data]
        except (json.JSONDecodeError, TypeError):
            self._favorites = []

        # Load history
        history_json = self.settings.value("history", "[]")
        try:
            history_data = json.loads(history_json)
            self._history = [ConnectionHistory.from_dict(h) for h in history_data]
        except (json.JSONDecodeError, TypeError):
            self._history = []

    def _save(self):
        """Save favorites and history to settings"""
        favorites_json = json.dumps([f.to_dict() for f in self._favorites])
        self.settings.setValue("favorites", favorites_json)

        history_json = json.dumps([h.to_dict() for h in self._history])
        self.settings.setValue("history", history_json)

        self.settings.sync()

    # Favorites management

    def get_favorites(self) -> List[NetworkConfig]:
        """Get all favorite connections"""
        return self._favorites.copy()

    def add_favorite(self, config: NetworkConfig):
        """Add a connection to favorites"""
        # Check for duplicates
        for fav in self._favorites:
            if fav.get_connection_id() == config.get_connection_id():
                return  # Already exists

        self._favorites.append(config)
        self._save()

    def remove_favorite(self, config: NetworkConfig):
        """Remove a connection from favorites"""
        self._favorites = [f for f in self._favorites
                         if f.get_connection_id() != config.get_connection_id()]
        self._save()

    def is_favorite(self, config: NetworkConfig) -> bool:
        """Check if a connection is in favorites"""
        return any(f.get_connection_id() == config.get_connection_id()
                  for f in self._favorites)

    # History management

    def get_history(self) -> List[ConnectionHistory]:
        """Get connection history, most recent first"""
        return sorted(self._history,
                     key=lambda h: h.last_connected,
                     reverse=True)

    def record_connection(self, config: NetworkConfig):
        """Record a successful connection"""
        conn_id = config.get_connection_id()

        # Update existing history entry or create new one
        for hist in self._history:
            if hist.config.get_connection_id() == conn_id:
                hist.last_connected = datetime.now()
                hist.connect_count += 1
                self._save()
                return

        # New entry
        self._history.append(ConnectionHistory(
            config=config,
            last_connected=datetime.now()
        ))

        # Trim to max size
        if len(self._history) > self.MAX_HISTORY:
            self._history = sorted(self._history,
                                  key=lambda h: h.last_connected,
                                  reverse=True)[:self.MAX_HISTORY]

        self._save()

    def clear_history(self):
        """Clear all connection history"""
        self._history = []
        self._save()

    # Default configurations

    @staticmethod
    def get_default_configs() -> List[NetworkConfig]:
        """Get list of common default configurations"""
        return [
            NetworkConfig(host="127.0.0.1", port=5000, name="Localhost TCP"),
            NetworkConfig(host="127.0.0.1", port=5000, protocol='UDP', name="Localhost UDP"),
            NetworkConfig(host="0.0.0.0", port=5000, mode='server', name="TCP Server"),
        ]
```

#### Task C.3: Adapt SettingsManager

**File:** `core/core.py` (extracted/modified)

```python
"""
Core utilities extracted from Serial Terminal
"""

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication


class SettingsManager:
    """Manages application settings using QSettings"""

    def __init__(self):
        # Change organization/app name for Network Terminal
        self.settings = QSettings("NetworkTerminal", "Settings")

    def get_show_launch_dialog(self) -> bool:
        """Get whether to show launch dialog on startup"""
        return self.settings.value("ui/show_launch_dialog", True, type=bool)

    def set_show_launch_dialog(self, show_dialog: bool):
        """Set whether to show launch dialog on startup"""
        self.settings.setValue("ui/show_launch_dialog", show_dialog)
        self.settings.sync()

    def get_last_config(self) -> dict:
        """Get last used connection configuration"""
        import json
        config_json = self.settings.value("connection/last_config", "{}")
        try:
            return json.loads(config_json)
        except json.JSONDecodeError:
            return {}

    def set_last_config(self, config: dict):
        """Save last used connection configuration"""
        import json
        self.settings.setValue("connection/last_config", json.dumps(config))
        self.settings.sync()


# Copy ResponsiveWindowManager unchanged from Serial Terminal
class ResponsiveWindowManager:
    """Manages responsive window sizing and layout decisions"""
    # ... (copy entire class from SerialTerminal/core/core.py)
```

**Acceptance Criteria for Stream C:**
- [ ] `NetworkConfig` dataclass with all fields
- [ ] Serialization/deserialization works
- [ ] `ConnectionManager` persists favorites
- [ ] `ConnectionManager` maintains history
- [ ] `SettingsManager` uses correct app name
- [ ] Unit tests for config validation

---

### PHASE 4: Integration (Sequential)

**Owner:** All developers
**Dependencies:** Streams A, B, C complete

#### Task 4.1: Wire Up Components

Connect all components:

1. **main.py** imports and initializes:
   - `NetworkMonitorWindow`
   - `ConnectionManager`
   - `SettingsManager`

2. **NetworkMonitorWindow** uses:
   - `NetworkTerminalPane` for each connection
   - `QuickConnectDialog` for new connections
   - `ConnectionManager` for favorites/history

3. **NetworkTerminalPane** uses:
   - `create_network_worker()` to create appropriate worker
   - `TerminalStreamFormatter` for display
   - `NetworkConfig` for settings

#### Task 4.2: Integration Testing

Test complete workflows:

| Test Case | Steps | Expected Result |
|-----------|-------|-----------------|
| TCP Client Connect | 1. New → Enter host:port → Connect | Status shows "Connected" |
| TCP Client Data | 2. Send data from remote | Data appears in terminal |
| TCP Server Accept | 1. New → Server mode → Start | Status shows "Listening" |
| UDP Send/Receive | 1. New UDP → Send data | Data transmitted |
| Split Pane | 1. Right-click → Split | Two panes appear |
| Tab Management | 1. Ctrl+N → New tab | New tab created |
| Hex Mode | 1. Right-click → Hex | Data shows as hex |
| Favorites | 1. Connect → Add to favorites | Appears in favorites menu |

#### Task 4.3: Fix Integration Issues

Address any issues discovered during integration testing.

---

### PHASE 5: Testing & Polish (Sequential)

**Owner:** All developers
**Dependencies:** Phase 4 complete

#### Task 5.1: Manual Testing Matrix

| Feature | Windows 10 | Windows 11 |
|---------|------------|------------|
| TCP Client | | |
| TCP Server | | |
| UDP Sender | | |
| UDP Receiver | | |
| Split Panes | | |
| Tabs | | |
| Favorites | | |
| History | | |
| Hex Mode | | |
| NMEA Colors | | |

#### Task 5.2: Performance Testing

- Startup time < 2 seconds
- Memory usage < 100 MB idle
- Handle 10 MB/s data rate without lag

#### Task 5.3: Documentation

- Update constants.py with correct app name/version
- Create README.md for Network Terminal
- Document keyboard shortcuts

#### Task 5.4: Build & Package

- Test PyInstaller build
- Verify single-file executable works
- Test on clean Windows installation

---

## 6. Interface Definitions

### 6.1 NetworkWorker Interface (Must Match SerialWorker)

```python
class NetworkWorker(QThread):
    """Interface contract for network workers"""

    # Signals (MUST match SerialWorker for compatibility)
    dataReceived = pyqtSignal(bytes)
    errorOccurred = pyqtSignal(str)
    connectionStateChanged = pyqtSignal(bool)

    def __init__(self, config: NetworkConfig): ...
    def run(self) -> None: ...
    def stop(self) -> None: ...
    def write(self, data: bytes) -> None: ...
```

### 6.2 NetworkConfig Interface (Replaces SerialConfig)

```python
@dataclass
class NetworkConfig:
    """Interface contract for network configuration"""

    host: str
    port: int
    protocol: str  # 'TCP' or 'UDP'
    mode: str      # 'client' or 'server'

    def get_display_string(self) -> str: ...
    def get_connection_id(self) -> str: ...
    def to_dict(self) -> dict: ...
    @classmethod
    def from_dict(cls, data: dict) -> 'NetworkConfig': ...
```

### 6.3 TerminalPane Interface (Preserved)

```python
class NetworkTerminalPane(QWidget):
    """Interface preserved from TerminalPane"""

    # Signals
    focusChanged = pyqtSignal(bool)
    splitRequested = pyqtSignal(object, str)
    closeRequested = pyqtSignal(object)

    def __init__(self, config: NetworkConfig, ...): ...
    def connect(self) -> None: ...
    def disconnect(self) -> None: ...
    def cleanup(self) -> None: ...
    def send_data(self, data: str) -> None: ...
    def get_status_info(self) -> str: ...
```

---

## 7. File Structure

### 7.1 Final Project Structure

```
NetworkTerminal/
├── main.py                           # Entry point
├── constants.py                      # App info, terminal colors
├── build.py                          # PyInstaller build script
├── requirements.txt                  # PyQt6
│
├── core/
│   ├── __init__.py
│   ├── core.py                       # SettingsManager, ResponsiveWindowManager
│   ├── network_config.py             # NetworkConfig dataclass
│   ├── network_worker.py             # NetworkWorker, TCP/UDP implementations
│   └── connection_manager.py         # Favorites, history management
│
├── ui/
│   ├── __init__.py
│   ├── dialogs/
│   │   ├── __init__.py
│   │   ├── terminal_dialog.py        # NetworkTerminalPane, NetworkMonitorWindow
│   │   └── connection_dialog.py      # QuickConnectDialog
│   │
│   ├── windows/
│   │   ├── __init__.py
│   │   └── terminal_formatter.py     # TerminalStreamFormatter (unchanged)
│   │
│   ├── components/
│   │   ├── __init__.py
│   │   └── ribbon_toolbar.py         # RibbonToolbar (unchanged)
│   │
│   ├── common/
│   │   ├── __init__.py
│   │   └── icons.py                  # Icons (unchanged)
│   │
│   └── resources.py                  # ResourceManager (unchanged)
│
├── assets/
│   ├── fonts/
│   │   ├── JetBrainsMono/
│   │   └── Poppins/
│   └── icons/
│
├── tests/
│   ├── __init__.py
│   ├── test_network_config.py
│   ├── test_network_worker.py
│   └── test_connection_manager.py
│
└── docs/
    ├── PRD_NetworkTerminal.md        # This document
    └── README.md                     # User documentation
```

### 7.2 Lines of Code Estimate

| Directory | New Lines | Adapted Lines | Copied Lines | Total |
|-----------|-----------|---------------|--------------|-------|
| core/ | 450 | 100 | 150 | 700 |
| ui/dialogs/ | 200 | 400 | 1500 | 2100 |
| ui/windows/ | 0 | 0 | 344 | 344 |
| ui/components/ | 0 | 10 | 140 | 150 |
| ui/common/ | 0 | 0 | 100 | 100 |
| main.py | 20 | 30 | 50 | 100 |
| **Total** | **670** | **540** | **2284** | **3494** |

---

## 8. Testing Requirements

### 8.1 Unit Tests

| Component | Test File | Coverage Target |
|-----------|-----------|-----------------|
| NetworkConfig | test_network_config.py | 100% |
| NetworkWorker | test_network_worker.py | 80% |
| ConnectionManager | test_connection_manager.py | 90% |

### 8.2 Integration Tests

| Test | Description |
|------|-------------|
| TCP Loopback | Client connects to local server, exchanges data |
| UDP Loopback | Sender/receiver on localhost |
| Multi-pane | Open 4 panes, each to different endpoint |
| Favorites | Add, remove, reload favorites |
| History | Connect, close app, reopen, check history |

### 8.3 Manual Test Cases

| ID | Test Case | Steps | Expected |
|----|-----------|-------|----------|
| MT-01 | TCP Connect | New → TCP → localhost:5000 → Connect | Shows "Connected" |
| MT-02 | TCP Disconnect | Connect → Disconnect | Shows "Disconnected" |
| MT-03 | TCP Server | New → Server → Listen 5000 | Shows "Listening" |
| MT-04 | UDP Send | New → UDP → Send "test" | No error |
| MT-05 | Split Pane | Right-click → Split Vertical | Two panes |
| MT-06 | New Tab | Ctrl+N | New tab appears |
| MT-07 | Close Tab | Ctrl+W | Tab closes |
| MT-08 | Hex Mode | Right-click → Hex Display | Data as hex |
| MT-09 | Font Size | Right-click → Font → 14pt | Font changes |
| MT-10 | Copy Text | Select → Ctrl+C | Text copied |

---

## 9. Risk Assessment

### 9.1 Technical Risks

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Socket blocking issues | Medium | High | Use non-blocking sockets, timeouts |
| Thread cleanup failures | Low | Medium | Follow SerialWorker cleanup pattern |
| UDP packet loss | High | Low | Document as expected UDP behavior |
| Multicast compatibility | Medium | Low | Make multicast optional feature |

### 9.2 Schedule Risks

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Integration issues | Medium | Medium | Clear interface definitions |
| Parallel development conflicts | Low | Low | Isolated file ownership |
| Testing delays | Medium | Medium | Automated tests first |

### 9.3 Dependencies

| Dependency | Version | Risk |
|------------|---------|------|
| PyQt6 | 6.4+ | Low - stable, well-documented |
| Python | 3.8+ | Low - standard library sockets |
| Windows | 10/11 | Low - target platform |

---

## Appendix A: Quick Reference - What to Copy vs. Create

### Files to COPY UNCHANGED

```
constants.py              → Copy, change AppInfo only
ui/windows/terminal_formatter.py
ui/components/ribbon_toolbar.py
ui/common/icons.py
ui/resources.py
assets/*
```

### Files to CREATE NEW

```
core/network_config.py
core/network_worker.py
core/connection_manager.py
ui/dialogs/connection_dialog.py
requirements.txt
```

### Files to ADAPT (Copy and Modify)

```
main.py                   → Change imports, app name
core/core.py              → Extract SettingsManager, ResponsiveWindowManager
ui/dialogs/terminal_dialog.py → Major adaptation
build.py                  → Change output name
```

---

## Appendix B: Code Snippets for Common Patterns

### B.1 Worker Signal Connection (Same as Serial)

```python
# In NetworkTerminalPane.connect()
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

### B.2 Display String Format

```python
# SerialConfig.get_display_string()
return f"{self.baudrate} {self.databits}{self.parity}{self.stopbits}"
# Example: "115200 8N1"

# NetworkConfig.get_display_string()
return f"{self.protocol} | {self.host}:{self.port}"
# Example: "TCP | 192.168.1.100:5000"
```

### B.3 Status Bar Format

```python
# Serial: "COM3: Connected | 115200 8N1 | RX: 1.2KB | TX: 256B"
# Network: "TCP 192.168.1.100:5000: Connected | RX: 1.2KB | TX: 256B"
```

---

## Document History

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | 2026-01-13 | Claude | Initial draft |

---

*End of Document*
