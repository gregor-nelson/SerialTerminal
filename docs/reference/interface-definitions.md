# Interface Definitions

[← Back to Master PRD](../PRD_NetworkTerminal.md)

---

## Overview

This document defines the API contracts between Network Terminal components. Adhering to these interfaces ensures compatibility between streams developed in parallel.

---

## Critical Interfaces

### NetworkWorker Signals (MUST Match SerialWorker)

The `NetworkWorker` class **MUST** emit these signals to be compatible with the UI layer:

```python
from PyQt6.QtCore import QThread, pyqtSignal

class NetworkWorker(QThread):
    """
    Base class for network workers.

    CRITICAL: These signals MUST match SerialWorker exactly.
    The UI layer depends on this interface.
    """

    # === REQUIRED SIGNALS (Match SerialWorker) ===
    dataReceived = pyqtSignal(bytes)
    """
    Emitted when data is received from the network.

    Args:
        data (bytes): Raw bytes received

    Usage:
        worker.dataReceived.connect(pane._on_data_received)
    """

    errorOccurred = pyqtSignal(str)
    """
    Emitted when an error occurs.

    Args:
        message (str): Human-readable error message

    Usage:
        worker.errorOccurred.connect(pane._on_error)
    """

    connectionStateChanged = pyqtSignal(bool)
    """
    Emitted when connection state changes.

    Args:
        connected (bool): True if connected, False if disconnected

    Usage:
        worker.connectionStateChanged.connect(pane._on_connection_state_changed)

    Notes:
        - Emit True AFTER successful connection
        - Emit False BEFORE cleanup in finally block
        - For server mode: True means "listening"
    """
```

### Additional NetworkWorker Signals (Optional)

```python
class NetworkWorker(QThread):
    # ... required signals above ...

    # === OPTIONAL SIGNALS (Network-specific) ===
    clientConnected = pyqtSignal(str)
    """
    Emitted when a client connects (server mode only).

    Args:
        address (str): Client address as "host:port"
    """

    clientDisconnected = pyqtSignal(str)
    """
    Emitted when a client disconnects (server mode only).

    Args:
        address (str): Client address as "host:port"
    """
```

---

## NetworkConfig API

### Class Definition

```python
from dataclasses import dataclass
from typing import Optional

@dataclass
class NetworkConfig:
    """
    Network connection configuration.

    This class replaces SerialConfig for network connections.
    """

    # === Required Fields ===
    host: str
    """Target hostname or IP address (e.g., "192.168.1.100", "localhost")"""

    port: int
    """Target port number (1-65535)"""

    # === Protocol Settings ===
    protocol: str = 'TCP'
    """Protocol type: 'TCP' or 'UDP'"""

    mode: str = 'client'
    """Connection mode: 'client' or 'server'"""

    # === TCP Options ===
    keepalive: bool = True
    """Enable TCP keepalive"""

    nodelay: bool = True
    """Enable TCP_NODELAY (disable Nagle algorithm)"""

    # === UDP Options ===
    broadcast: bool = False
    """Enable broadcast for UDP"""

    multicast_group: Optional[str] = None
    """Multicast group address (e.g., "224.0.0.1")"""

    # === Common Options ===
    buffer_size: int = 4096
    """Receive buffer size in bytes"""

    timeout: float = 5.0
    """Connection timeout in seconds"""
```

### Required Methods

```python
class NetworkConfig:
    def get_display_string(self) -> str:
        """
        Get display string for status bar.

        Returns:
            Human-readable connection string

        Examples:
            "TCP | 192.168.1.100:5000"
            "TCP Server | 0.0.0.0:5000"
            "UDP Broadcast | 255.255.255.255:5000"
        """
        pass

    def get_connection_id(self) -> str:
        """
        Get unique identifier for this connection.

        Returns:
            Unique string like "TCP_client_192.168.1.100_5000"
        """
        pass

    def to_dict(self) -> dict:
        """
        Convert to dictionary for serialization.

        Returns:
            Dictionary of all fields
        """
        pass

    @classmethod
    def from_dict(cls, data: dict) -> 'NetworkConfig':
        """
        Create from dictionary.

        Args:
            data: Dictionary with field values

        Returns:
            NetworkConfig instance
        """
        pass
```

---

## ConnectionManager API

### Class Definition

```python
class ConnectionManager:
    """
    Manages connection favorites and history.

    Uses QSettings for persistent storage.
    """

    MAX_HISTORY = 20
    """Maximum history entries to keep"""
```

### Required Methods

```python
class ConnectionManager:
    def get_favorites(self) -> List[NetworkConfig]:
        """
        Get all favorite connections.

        Returns:
            List of NetworkConfig objects
        """
        pass

    def add_favorite(self, config: NetworkConfig) -> bool:
        """
        Add connection to favorites.

        Args:
            config: Connection to add

        Returns:
            True if added, False if duplicate
        """
        pass

    def remove_favorite(self, config: NetworkConfig) -> bool:
        """
        Remove connection from favorites.

        Args:
            config: Connection to remove

        Returns:
            True if removed, False if not found
        """
        pass

    def is_favorite(self, config: NetworkConfig) -> bool:
        """
        Check if connection is in favorites.

        Args:
            config: Connection to check

        Returns:
            True if in favorites
        """
        pass

    def get_history(self, limit: Optional[int] = None) -> List[ConnectionHistory]:
        """
        Get connection history.

        Args:
            limit: Maximum entries to return

        Returns:
            List sorted by most recent first
        """
        pass

    def record_connection(self, config: NetworkConfig):
        """
        Record a successful connection.

        Args:
            config: Connection that was made
        """
        pass
```

---

## Factory Function API

```python
def create_network_worker(config: NetworkConfig) -> NetworkWorker:
    """
    Factory function to create appropriate network worker.

    Args:
        config: NetworkConfig with protocol, mode, host, port

    Returns:
        TCPClientWorker, TCPServerWorker, or UDPWorker

    Raises:
        ValueError: If protocol or mode is invalid

    Examples:
        # TCP Client
        config = NetworkConfig(host="127.0.0.1", port=5000,
                              protocol="TCP", mode="client")
        worker = create_network_worker(config)
        assert isinstance(worker, TCPClientWorker)

        # TCP Server
        config = NetworkConfig(host="0.0.0.0", port=5000,
                              protocol="TCP", mode="server")
        worker = create_network_worker(config)
        assert isinstance(worker, TCPServerWorker)

        # UDP
        config = NetworkConfig(host="127.0.0.1", port=5000,
                              protocol="UDP", mode="client")
        worker = create_network_worker(config)
        assert isinstance(worker, UDPWorker)
    """
    pass
```

---

## UI Component Interfaces

### NetworkTerminalPane

```python
class NetworkTerminalPane(QWidget):
    """
    Terminal pane for network connections.

    Signals (must match TerminalPane):
    """

    # === Signals ===
    focusChanged = pyqtSignal(bool)
    """Emitted when focus changes. Args: has_focus (bool)"""

    splitRequested = pyqtSignal(object, str)
    """Emitted when split requested. Args: source_pane, direction ('vertical'/'horizontal')"""

    closeRequested = pyqtSignal(object)
    """Emitted when close requested. Args: source_pane"""
```

### Required Methods

```python
class NetworkTerminalPane:
    def __init__(self, config: NetworkConfig, parent=None,
                 main_window=None, container=None):
        """
        Initialize terminal pane.

        Args:
            config: NetworkConfig for this connection
            parent: Parent widget
            main_window: Reference to main window
            container: Reference to SplitContainer
        """
        pass

    def connect(self):
        """
        Connect using network worker.

        Creates worker, connects signals, starts thread.
        """
        pass

    def disconnect(self):
        """
        Disconnect from network.

        Stops worker, cleans up resources.
        """
        pass

    def cleanup(self):
        """
        Single point of cleanup.

        Called by disconnect() and when pane is closed.
        """
        pass

    def send_data(self, data: str):
        """
        Send data over network.

        Args:
            data: String data to send
        """
        pass

    def get_status_info(self) -> str:
        """
        Get status for status bar.

        Returns:
            String like "TCP 127.0.0.1:5000: Connected | RX: 1KB | TX: 256B"
        """
        pass
```

---

## QuickConnectDialog API

```python
class QuickConnectDialog(QDialog):
    """
    Dialog for configuring new connections.
    """

    def __init__(self, parent=None, initial_config: NetworkConfig = None):
        """
        Initialize dialog.

        Args:
            parent: Parent widget
            initial_config: Pre-fill with existing config (optional)
        """
        pass

    def get_config(self) -> NetworkConfig:
        """
        Get configured NetworkConfig.

        Returns:
            NetworkConfig with user-specified settings

        Notes:
            Only call after dialog.exec() returns Accepted
        """
        pass
```

---

## ConnectionHistoryDialog API

```python
class ConnectionHistoryDialog(QDialog):
    """
    Dialog showing favorites and history.
    """

    # === Signals ===
    configSelected = pyqtSignal(object)
    """Emitted when user selects a config. Args: NetworkConfig"""

    def __init__(self, connection_manager: ConnectionManager, parent=None):
        """
        Initialize dialog.

        Args:
            connection_manager: ConnectionManager instance
            parent: Parent widget
        """
        pass
```

---

## Signal Connection Examples

### Correct Usage

```python
# In NetworkTerminalPane.connect()
def connect(self):
    self.network_worker = create_network_worker(self.config)

    # Use QueuedConnection for thread safety
    self.network_worker.dataReceived.connect(
        self._on_data_received,
        Qt.ConnectionType.QueuedConnection
    )
    self.network_worker.errorOccurred.connect(
        self._on_error,
        Qt.ConnectionType.QueuedConnection
    )
    self.network_worker.connectionStateChanged.connect(
        self._on_connection_state_changed,
        Qt.ConnectionType.QueuedConnection
    )

    self.network_worker.start()
```

### Signal Handler Signatures

```python
def _on_data_received(self, data: bytes):
    """Handle received data."""
    pass

def _on_error(self, error_msg: str):
    """Handle error."""
    pass

def _on_connection_state_changed(self, connected: bool):
    """Handle connection state change."""
    pass
```

---

## Validation Checklist

Use this checklist to verify interface compliance:

### NetworkWorker

- [ ] `dataReceived` signal emits `bytes`
- [ ] `errorOccurred` signal emits `str`
- [ ] `connectionStateChanged` signal emits `bool`
- [ ] `stop()` method blocks until thread exits
- [ ] `write(data: bytes)` method queues data

### NetworkConfig

- [ ] `__post_init__` validates all fields
- [ ] `get_display_string()` returns formatted string
- [ ] `to_dict()` / `from_dict()` round-trip correctly
- [ ] `get_connection_id()` returns unique string

### Factory Function

- [ ] Returns `TCPClientWorker` for TCP client
- [ ] Returns `TCPServerWorker` for TCP server
- [ ] Returns `UDPWorker` for UDP
- [ ] Raises `ValueError` for invalid protocol/mode

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Architecture Analysis](./architecture-analysis.md)
- [Stream A: Core Network](../phases/stream-a-core-network.md)
- [Stream C: Config/Data](../phases/stream-c-config-data.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md)
