# Testing Requirements

[← Back to Master PRD](../PRD_NetworkTerminal.md)

---

## Overview

This document specifies all testing requirements for the Network Terminal application, including unit tests, integration tests, and manual test cases.

---

## Test Organization

```
tests/
├── __init__.py
├── test_network_config.py        # NetworkConfig unit tests
├── test_network_worker.py        # NetworkWorker unit tests
├── test_connection_manager.py    # ConnectionManager unit tests
├── test_integration.py           # Integration tests
└── conftest.py                   # Pytest fixtures (optional)
```

---

## Unit Tests

### test_network_config.py

#### Test Cases

```python
"""Unit tests for NetworkConfig"""

import pytest
from core.network_config import NetworkConfig


class TestNetworkConfigCreation:
    """Test NetworkConfig creation and validation"""

    def test_basic_creation(self):
        """Test basic config creation with required fields"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        assert config.host == "127.0.0.1"
        assert config.port == 5000
        assert config.protocol == "TCP"
        assert config.mode == "client"

    def test_tcp_server_creation(self):
        """Test TCP server config"""
        config = NetworkConfig(
            host="0.0.0.0",
            port=5000,
            protocol="TCP",
            mode="server"
        )
        assert config.mode == "server"

    def test_udp_creation(self):
        """Test UDP config"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=5000,
            protocol="UDP"
        )
        assert config.protocol == "UDP"

    def test_all_options(self):
        """Test config with all options"""
        config = NetworkConfig(
            host="192.168.1.100",
            port=8080,
            protocol="TCP",
            mode="client",
            keepalive=True,
            nodelay=False,
            timeout=10.0,
            buffer_size=8192
        )
        assert config.keepalive is True
        assert config.nodelay is False
        assert config.timeout == 10.0
        assert config.buffer_size == 8192


class TestNetworkConfigValidation:
    """Test NetworkConfig validation"""

    def test_invalid_protocol(self):
        """Invalid protocol should raise ValueError"""
        with pytest.raises(ValueError, match="Invalid protocol"):
            NetworkConfig(host="127.0.0.1", port=5000, protocol="HTTP")

    def test_invalid_mode(self):
        """Invalid mode should raise ValueError"""
        with pytest.raises(ValueError, match="Invalid mode"):
            NetworkConfig(host="127.0.0.1", port=5000, mode="broadcast")

    def test_invalid_port_zero(self):
        """Port 0 should raise ValueError"""
        with pytest.raises(ValueError, match="Invalid port"):
            NetworkConfig(host="127.0.0.1", port=0)

    def test_invalid_port_negative(self):
        """Negative port should raise ValueError"""
        with pytest.raises(ValueError, match="Invalid port"):
            NetworkConfig(host="127.0.0.1", port=-1)

    def test_invalid_port_high(self):
        """Port > 65535 should raise ValueError"""
        with pytest.raises(ValueError, match="Invalid port"):
            NetworkConfig(host="127.0.0.1", port=70000)

    def test_empty_host(self):
        """Empty host should raise ValueError"""
        with pytest.raises(ValueError, match="Host cannot be empty"):
            NetworkConfig(host="", port=5000)

    def test_valid_multicast_group(self):
        """Valid multicast group should succeed"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=5000,
            protocol="UDP",
            multicast_group="224.0.0.1"
        )
        assert config.multicast_group == "224.0.0.1"

    def test_invalid_multicast_group(self):
        """Invalid multicast group should raise ValueError"""
        with pytest.raises(ValueError, match="Multicast address must be in range"):
            NetworkConfig(
                host="127.0.0.1",
                port=5000,
                multicast_group="192.168.1.1"  # Not multicast range
            )


class TestNetworkConfigDisplayString:
    """Test get_display_string() method"""

    def test_tcp_client_display(self):
        """TCP client display string"""
        config = NetworkConfig(host="192.168.1.100", port=5000)
        display = config.get_display_string()
        assert "TCP" in display
        assert "192.168.1.100:5000" in display

    def test_tcp_server_display(self):
        """TCP server display string"""
        config = NetworkConfig(host="0.0.0.0", port=5000, mode="server")
        display = config.get_display_string()
        assert "Server" in display

    def test_udp_broadcast_display(self):
        """UDP broadcast display string"""
        config = NetworkConfig(
            host="255.255.255.255",
            port=5000,
            protocol="UDP",
            broadcast=True
        )
        display = config.get_display_string()
        assert "Broadcast" in display


class TestNetworkConfigSerialization:
    """Test serialization/deserialization"""

    def test_to_dict(self):
        """to_dict() returns correct dictionary"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        data = config.to_dict()
        assert data["host"] == "127.0.0.1"
        assert data["port"] == 5000
        assert data["protocol"] == "TCP"

    def test_from_dict(self):
        """from_dict() creates correct config"""
        data = {"host": "127.0.0.1", "port": 5000, "protocol": "UDP"}
        config = NetworkConfig.from_dict(data)
        assert config.host == "127.0.0.1"
        assert config.port == 5000
        assert config.protocol == "UDP"

    def test_to_json_from_json_roundtrip(self):
        """JSON roundtrip preserves data"""
        original = NetworkConfig(
            host="192.168.1.100",
            port=8080,
            protocol="TCP",
            mode="server",
            timeout=15.0
        )
        json_str = original.to_json()
        restored = NetworkConfig.from_json(json_str)

        assert restored.host == original.host
        assert restored.port == original.port
        assert restored.protocol == original.protocol
        assert restored.mode == original.mode
        assert restored.timeout == original.timeout

    def test_copy_with_changes(self):
        """copy() creates modified copy"""
        original = NetworkConfig(host="127.0.0.1", port=5000)
        copy = original.copy(port=6000, protocol="UDP")

        assert original.port == 5000
        assert original.protocol == "TCP"
        assert copy.port == 6000
        assert copy.protocol == "UDP"
        assert copy.host == original.host


class TestNetworkConfigConnectionId:
    """Test get_connection_id() method"""

    def test_unique_ids(self):
        """Different configs have different IDs"""
        config1 = NetworkConfig(host="127.0.0.1", port=5000)
        config2 = NetworkConfig(host="127.0.0.1", port=5001)
        config3 = NetworkConfig(host="127.0.0.1", port=5000, protocol="UDP")

        assert config1.get_connection_id() != config2.get_connection_id()
        assert config1.get_connection_id() != config3.get_connection_id()

    def test_same_config_same_id(self):
        """Same config values produce same ID"""
        config1 = NetworkConfig(host="127.0.0.1", port=5000)
        config2 = NetworkConfig(host="127.0.0.1", port=5000)

        assert config1.get_connection_id() == config2.get_connection_id()
```

---

### test_network_worker.py

#### Test Cases

```python
"""Unit tests for NetworkWorker classes"""

import pytest
import socket
import threading
import time
from unittest.mock import Mock, patch

from PyQt6.QtCore import QCoreApplication
from core.network_config import NetworkConfig
from core.network_worker import (
    NetworkWorker,
    TCPClientWorker,
    TCPServerWorker,
    UDPWorker,
    create_network_worker
)


@pytest.fixture
def app():
    """Create Qt application for event loop"""
    app = QCoreApplication.instance()
    if not app:
        app = QCoreApplication([])
    return app


@pytest.fixture
def echo_server():
    """Create a simple TCP echo server for testing"""
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server_socket.bind(('127.0.0.1', 0))  # Random available port
    server_socket.listen(1)
    port = server_socket.getsockname()[1]

    def run_server():
        try:
            conn, addr = server_socket.accept()
            conn.settimeout(5)
            while True:
                data = conn.recv(1024)
                if not data:
                    break
                conn.sendall(data)
            conn.close()
        except:
            pass
        finally:
            server_socket.close()

    thread = threading.Thread(target=run_server, daemon=True)
    thread.start()

    yield port

    server_socket.close()


class TestFactoryFunction:
    """Test create_network_worker() factory"""

    def test_tcp_client(self):
        """Factory returns TCPClientWorker for TCP client"""
        config = NetworkConfig(
            host="127.0.0.1", port=5000,
            protocol="TCP", mode="client"
        )
        worker = create_network_worker(config)
        assert isinstance(worker, TCPClientWorker)

    def test_tcp_server(self):
        """Factory returns TCPServerWorker for TCP server"""
        config = NetworkConfig(
            host="0.0.0.0", port=5000,
            protocol="TCP", mode="server"
        )
        worker = create_network_worker(config)
        assert isinstance(worker, TCPServerWorker)

    def test_udp(self):
        """Factory returns UDPWorker for UDP"""
        config = NetworkConfig(
            host="127.0.0.1", port=5000,
            protocol="UDP", mode="client"
        )
        worker = create_network_worker(config)
        assert isinstance(worker, UDPWorker)

    def test_invalid_protocol(self):
        """Factory raises ValueError for invalid protocol"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        config.protocol = "HTTP"  # Bypass validation for test
        with pytest.raises(ValueError, match="Invalid protocol"):
            create_network_worker(config)

    def test_invalid_mode(self):
        """Factory raises ValueError for invalid mode"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        config.mode = "broadcast"  # Bypass validation for test
        with pytest.raises(ValueError, match="Invalid TCP mode"):
            create_network_worker(config)


class TestNetworkWorkerBase:
    """Test NetworkWorker base class"""

    def test_stop_sets_flags(self):
        """stop() sets running flag and stop event"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        worker = TCPClientWorker(config)

        worker.running = True
        worker.stop()

        assert worker.running is False
        assert worker._stop_event.is_set()

    def test_write_queues_data(self):
        """write() adds data to queue when running"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        worker = TCPClientWorker(config)
        worker.running = True

        worker.write(b"test data")

        assert not worker.write_queue.empty()
        assert worker.write_queue.get() == b"test data"

    def test_write_ignores_when_not_running(self):
        """write() ignores data when not running"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        worker = TCPClientWorker(config)
        worker.running = False

        worker.write(b"test data")

        assert worker.write_queue.empty()


class TestTCPClientWorker:
    """Test TCPClientWorker"""

    def test_connects_to_server(self, app, echo_server):
        """TCPClientWorker connects to server"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=echo_server,
            protocol="TCP",
            mode="client"
        )
        worker = TCPClientWorker(config)

        # Track signals
        connected = []
        worker.connectionStateChanged.connect(lambda c: connected.append(c))

        # Start worker
        worker.start()
        time.sleep(0.5)  # Wait for connection

        assert True in connected

        # Cleanup
        worker.stop()
        worker.wait(2000)

    def test_receives_data(self, app, echo_server):
        """TCPClientWorker receives data"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=echo_server,
            protocol="TCP",
            mode="client"
        )
        worker = TCPClientWorker(config)

        # Track received data
        received = []
        worker.dataReceived.connect(lambda d: received.append(d))

        # Connect and send
        worker.start()
        time.sleep(0.3)
        worker.write(b"hello")
        time.sleep(0.3)

        assert len(received) > 0
        assert b"hello" in received

        # Cleanup
        worker.stop()
        worker.wait(2000)

    def test_connection_refused_error(self, app):
        """TCPClientWorker handles connection refused"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=59999,  # No server listening
            protocol="TCP",
            mode="client",
            timeout=1.0
        )
        worker = TCPClientWorker(config)

        # Track errors
        errors = []
        worker.errorOccurred.connect(lambda e: errors.append(e))

        worker.start()
        time.sleep(2)

        assert len(errors) > 0
        assert any("refused" in e.lower() for e in errors)

        worker.stop()
        worker.wait(2000)


class TestTCPServerWorker:
    """Test TCPServerWorker"""

    def test_starts_listening(self, app):
        """TCPServerWorker starts listening"""
        config = NetworkConfig(
            host="0.0.0.0",
            port=0,  # Random port
            protocol="TCP",
            mode="server"
        )
        worker = TCPServerWorker(config)

        connected = []
        worker.connectionStateChanged.connect(lambda c: connected.append(c))

        worker.start()
        time.sleep(0.3)

        assert True in connected

        worker.stop()
        worker.wait(2000)

    def test_port_in_use_error(self, app):
        """TCPServerWorker handles port in use"""
        # First server
        server1_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server1_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        server1_socket.bind(('127.0.0.1', 0))
        server1_socket.listen(1)
        port = server1_socket.getsockname()[1]

        # Try second server on same port
        config = NetworkConfig(
            host="127.0.0.1",
            port=port,
            protocol="TCP",
            mode="server"
        )
        worker = TCPServerWorker(config)

        errors = []
        worker.errorOccurred.connect(lambda e: errors.append(e))

        worker.start()
        time.sleep(0.5)

        assert len(errors) > 0
        assert any("in use" in e.lower() for e in errors)

        worker.stop()
        worker.wait(2000)
        server1_socket.close()


class TestUDPWorker:
    """Test UDPWorker"""

    def test_sends_datagram(self, app):
        """UDPWorker sends datagrams"""
        # Receiver
        receiver = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        receiver.bind(('127.0.0.1', 0))
        receiver.settimeout(2)
        port = receiver.getsockname()[1]

        # Sender
        config = NetworkConfig(
            host="127.0.0.1",
            port=port,
            protocol="UDP",
            mode="client"
        )
        worker = UDPWorker(config)
        worker.start()
        time.sleep(0.2)
        worker.write(b"test udp")
        time.sleep(0.2)

        # Check received
        try:
            data, addr = receiver.recvfrom(1024)
            assert data == b"test udp"
        finally:
            worker.stop()
            worker.wait(2000)
            receiver.close()

    def test_receives_datagram(self, app):
        """UDPWorker receives datagrams"""
        config = NetworkConfig(
            host="0.0.0.0",
            port=0,  # Random port
            protocol="UDP",
            mode="server"
        )
        worker = UDPWorker(config)

        received = []
        worker.dataReceived.connect(lambda d: received.append(d))

        worker.start()
        time.sleep(0.2)

        # Get the bound port
        port = worker.socket.getsockname()[1]

        # Send test datagram
        sender = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sender.sendto(b"test data", ('127.0.0.1', port))
        sender.close()

        time.sleep(0.2)

        assert len(received) > 0
        assert b"test data" in received

        worker.stop()
        worker.wait(2000)
```

---

### test_connection_manager.py

#### Test Cases

```python
"""Unit tests for ConnectionManager"""

import pytest
from datetime import datetime
from unittest.mock import Mock, patch

from PyQt6.QtCore import QCoreApplication, QSettings
from core.network_config import NetworkConfig
from core.connection_manager import ConnectionManager, ConnectionHistory


@pytest.fixture
def app():
    """Create Qt application for QSettings"""
    app = QCoreApplication.instance()
    if not app:
        app = QCoreApplication([])
    return app


@pytest.fixture
def clean_settings(app):
    """Clean settings before each test"""
    settings = QSettings("NetworkTerminal", "ConnectionManager")
    settings.clear()
    settings.sync()
    yield
    settings.clear()
    settings.sync()


class TestConnectionManagerFavorites:
    """Test favorites functionality"""

    def test_add_favorite(self, clean_settings):
        """add_favorite() adds config to favorites"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)

        result = manager.add_favorite(config)

        assert result is True
        assert len(manager.get_favorites()) == 1
        assert manager.get_favorites()[0].host == "127.0.0.1"

    def test_add_duplicate_favorite(self, clean_settings):
        """add_favorite() returns False for duplicate"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)

        manager.add_favorite(config)
        result = manager.add_favorite(config)

        assert result is False
        assert len(manager.get_favorites()) == 1

    def test_remove_favorite(self, clean_settings):
        """remove_favorite() removes config"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)
        manager.add_favorite(config)

        result = manager.remove_favorite(config)

        assert result is True
        assert len(manager.get_favorites()) == 0

    def test_remove_nonexistent_favorite(self, clean_settings):
        """remove_favorite() returns False if not found"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)

        result = manager.remove_favorite(config)

        assert result is False

    def test_is_favorite(self, clean_settings):
        """is_favorite() returns correct status"""
        manager = ConnectionManager()
        config1 = NetworkConfig(host="127.0.0.1", port=5000)
        config2 = NetworkConfig(host="127.0.0.1", port=5001)
        manager.add_favorite(config1)

        assert manager.is_favorite(config1) is True
        assert manager.is_favorite(config2) is False

    def test_favorites_persist(self, clean_settings):
        """Favorites persist across instances"""
        manager1 = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)
        manager1.add_favorite(config)

        manager2 = ConnectionManager()
        assert len(manager2.get_favorites()) == 1
        assert manager2.get_favorites()[0].host == "127.0.0.1"


class TestConnectionManagerHistory:
    """Test history functionality"""

    def test_record_connection(self, clean_settings):
        """record_connection() adds to history"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)

        manager.record_connection(config)

        history = manager.get_history()
        assert len(history) == 1
        assert history[0].config.host == "127.0.0.1"
        assert history[0].connect_count == 1

    def test_record_connection_increments_count(self, clean_settings):
        """record_connection() increments count for same config"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)

        manager.record_connection(config)
        manager.record_connection(config)

        history = manager.get_history()
        assert len(history) == 1
        assert history[0].connect_count == 2

    def test_history_sorted_by_recent(self, clean_settings):
        """get_history() returns most recent first"""
        manager = ConnectionManager()
        config1 = NetworkConfig(host="127.0.0.1", port=5000)
        config2 = NetworkConfig(host="127.0.0.1", port=5001)

        manager.record_connection(config1)
        manager.record_connection(config2)

        history = manager.get_history()
        assert history[0].config.port == 5001  # Most recent

    def test_history_limit(self, clean_settings):
        """get_history() respects limit parameter"""
        manager = ConnectionManager()
        for i in range(10):
            config = NetworkConfig(host="127.0.0.1", port=5000 + i)
            manager.record_connection(config)

        history = manager.get_history(limit=5)
        assert len(history) == 5

    def test_history_max_size(self, clean_settings):
        """History is limited to MAX_HISTORY"""
        manager = ConnectionManager()
        for i in range(30):  # More than MAX_HISTORY
            config = NetworkConfig(host="127.0.0.1", port=5000 + i)
            manager.record_connection(config)

        history = manager.get_history()
        assert len(history) <= ConnectionManager.MAX_HISTORY

    def test_clear_history(self, clean_settings):
        """clear_history() removes all history"""
        manager = ConnectionManager()
        config = NetworkConfig(host="127.0.0.1", port=5000)
        manager.record_connection(config)

        manager.clear_history()

        assert len(manager.get_history()) == 0


class TestConnectionManagerSearch:
    """Test search functionality"""

    def test_search_by_host(self, clean_settings):
        """search() finds by host"""
        manager = ConnectionManager()
        config1 = NetworkConfig(host="192.168.1.100", port=5000)
        config2 = NetworkConfig(host="10.0.0.1", port=5000)
        manager.add_favorite(config1)
        manager.add_favorite(config2)

        results = manager.search("192.168")

        assert len(results) == 1
        assert results[0].host == "192.168.1.100"

    def test_search_by_port(self, clean_settings):
        """search() finds by port"""
        manager = ConnectionManager()
        config1 = NetworkConfig(host="127.0.0.1", port=5000)
        config2 = NetworkConfig(host="127.0.0.1", port=8080)
        manager.add_favorite(config1)
        manager.add_favorite(config2)

        results = manager.search("8080")

        assert len(results) == 1
        assert results[0].port == 8080

    def test_search_case_insensitive(self, clean_settings):
        """search() is case insensitive"""
        manager = ConnectionManager()
        config = NetworkConfig(host="MyServer.local", port=5000)
        manager.add_favorite(config)

        results = manager.search("myserver")

        assert len(results) == 1


class TestConnectionHistory:
    """Test ConnectionHistory dataclass"""

    def test_to_dict(self):
        """to_dict() returns correct dictionary"""
        config = NetworkConfig(host="127.0.0.1", port=5000)
        history = ConnectionHistory(
            config=config,
            last_connected=datetime(2024, 1, 1, 12, 0, 0),
            connect_count=5
        )

        data = history.to_dict()

        assert data["connect_count"] == 5
        assert "config" in data
        assert "last_connected" in data

    def test_from_dict(self):
        """from_dict() creates correct instance"""
        data = {
            "config": {"host": "127.0.0.1", "port": 5000},
            "last_connected": "2024-01-01T12:00:00",
            "connect_count": 5
        }

        history = ConnectionHistory.from_dict(data)

        assert history.config.host == "127.0.0.1"
        assert history.connect_count == 5
```

---

## Integration Tests

### test_integration.py

```python
"""Integration tests for Network Terminal"""

import pytest
import socket
import threading
import time

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

from core.network_config import NetworkConfig
from core.network_worker import create_network_worker
from ui.dialogs.terminal_dialog import NetworkTerminalPane


@pytest.fixture
def app():
    """Create Qt application"""
    app = QApplication.instance()
    if not app:
        app = QApplication([])
    return app


@pytest.fixture
def echo_server():
    """TCP echo server"""
    # ... same as in test_network_worker.py


class TestConnectionFlow:
    """Test complete connection flow"""

    def test_dialog_to_pane_flow(self, app):
        """Config from dialog creates working pane"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=5000,
            protocol="TCP",
            mode="client"
        )

        pane = NetworkTerminalPane(config)

        assert pane.config.host == "127.0.0.1"
        assert pane.config.port == 5000
        assert pane.network_worker is None  # Not connected yet

    def test_connect_disconnect_cycle(self, app, echo_server):
        """Pane can connect and disconnect cleanly"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=echo_server,
            protocol="TCP",
            mode="client"
        )

        pane = NetworkTerminalPane(config)

        # Connect
        pane.connect()
        time.sleep(0.5)
        assert pane.is_connected is True

        # Disconnect
        pane.disconnect()
        time.sleep(0.3)
        assert pane.is_connected is False
        assert pane.network_worker is None


class TestDataFlow:
    """Test data flow through components"""

    def test_data_reaches_terminal(self, app, echo_server):
        """Data from worker displays in terminal"""
        config = NetworkConfig(
            host="127.0.0.1",
            port=echo_server,
            protocol="TCP",
            mode="client"
        )

        pane = NetworkTerminalPane(config)
        pane.connect()
        time.sleep(0.3)

        # Send data (will echo back)
        pane.send_data("test message")
        time.sleep(0.3)

        # Check terminal content
        content = pane.terminal.toPlainText()
        assert "test message" in content or pane.rx_bytes > 0

        pane.disconnect()
```

---

## Manual Test Procedures

See [Phase 5: Testing & Polish](../phases/phase-5-testing.md) for complete manual test procedures.

---

## Test Coverage Requirements

| Component | Minimum Coverage |
|-----------|------------------|
| `NetworkConfig` | 90% |
| `NetworkWorker` | 80% |
| `ConnectionManager` | 85% |
| Overall | 75% |

---

## Running Tests

```bash
# Install test dependencies
pip install pytest pytest-cov pytest-qt

# Run all tests
pytest tests/

# Run with coverage
pytest tests/ --cov=core --cov=ui --cov-report=html

# Run specific test file
pytest tests/test_network_config.py

# Run specific test class
pytest tests/test_network_config.py::TestNetworkConfigValidation

# Run with verbose output
pytest tests/ -v
```

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Phase 5: Testing & Polish](../phases/phase-5-testing.md)
- [Interface Definitions](./interface-definitions.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md)
