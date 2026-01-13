# Stream A: Core Network Layer

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Phase 0](./phase-0-project-setup.md)

---

## Overview

| Attribute | Value |
|-----------|-------|
| **Stream** | A - Core Network |
| **Owner** | Developer 1 |
| **Dependencies** | Phase 0 complete |
| **Parallel With** | Stream B, Stream C |
| **Estimated Effort** | 5 tasks, ~400 new lines |
| **Status** | Not Started |

---

## Purpose

Implement the `NetworkWorker` classes that handle TCP and UDP communication. These classes must match the `SerialWorker` signal interface to enable drop-in replacement in the UI layer.

---

## Critical Interface Contract

The `NetworkWorker` classes **MUST** emit these signals to be compatible with the UI:

```python
# These signals MUST match SerialWorker exactly
dataReceived = pyqtSignal(bytes)           # Emitted when data arrives
errorOccurred = pyqtSignal(str)            # Emitted on errors
connectionStateChanged = pyqtSignal(bool)  # True=connected, False=disconnected
```

See [Interface Definitions](../reference/interface-definitions.md) for complete contract.

---

## Task Checklist

- [ ] [Task A.1: Create NetworkWorker Base Class](#task-a1-create-networkworker-base-class)
- [ ] [Task A.2: Implement TCPClientWorker](#task-a2-implement-tcpclientworker)
- [ ] [Task A.3: Implement TCPServerWorker](#task-a3-implement-tcpserverworker)
- [ ] [Task A.4: Implement UDPWorker](#task-a4-implement-udpworker)
- [ ] [Task A.5: Create Factory Function](#task-a5-create-factory-function)

---

## Task A.1: Create NetworkWorker Base Class

### Objective

Create the abstract base class that defines the interface for all network workers.

### File

`core/network_worker.py`

### Code

```python
#!/usr/bin/env python3
"""
Network Worker - Background thread for TCP/UDP communication
Mirrors SerialWorker interface for drop-in replacement
"""

from PyQt6.QtCore import QThread, pyqtSignal
from typing import Optional, TYPE_CHECKING
import socket
import queue
import threading

if TYPE_CHECKING:
    from .network_config import NetworkConfig


class NetworkWorker(QThread):
    """
    Base class for network communication workers.

    IMPORTANT: Signal interface MUST match SerialWorker for UI compatibility.

    Usage:
        worker = TCPClientWorker(config)
        worker.dataReceived.connect(on_data)
        worker.errorOccurred.connect(on_error)
        worker.connectionStateChanged.connect(on_state_change)
        worker.start()

        # Send data
        worker.write(b"Hello")

        # Stop
        worker.stop()
    """

    # === REQUIRED SIGNALS (Must match SerialWorker) ===
    dataReceived = pyqtSignal(bytes)
    errorOccurred = pyqtSignal(str)
    connectionStateChanged = pyqtSignal(bool)

    # === ADDITIONAL SIGNALS (Network-specific) ===
    clientConnected = pyqtSignal(str)      # Client address (server mode only)
    clientDisconnected = pyqtSignal(str)   # Client address (server mode only)

    def __init__(self, config: 'NetworkConfig'):
        """
        Initialize the network worker.

        Args:
            config: NetworkConfig instance with connection parameters
        """
        super().__init__()
        self.config = config
        self.socket: Optional[socket.socket] = None
        self.client_socket: Optional[socket.socket] = None  # For server mode
        self.running = False
        self.write_queue: queue.Queue[bytes] = queue.Queue()
        self._stop_event = threading.Event()

    def run(self):
        """
        Main thread loop - must be implemented by subclasses.

        Subclasses must:
        1. Create and configure socket
        2. Connect/bind as appropriate
        3. Emit connectionStateChanged(True) on success
        4. Loop: read data, emit dataReceived; write from queue
        5. Handle errors, emit errorOccurred
        6. Emit connectionStateChanged(False) on exit
        """
        raise NotImplementedError("Subclasses must implement run()")

    def stop(self):
        """
        Stop the worker thread safely with graceful shutdown.
        Blocks until thread exits or timeout.
        """
        # Signal thread to stop
        self._stop_event.set()
        self.running = False

        # Close sockets to interrupt blocking operations
        for sock in [self.client_socket, self.socket]:
            if sock:
                try:
                    sock.shutdown(socket.SHUT_RDWR)
                except OSError:
                    pass  # Socket may already be closed
                try:
                    sock.close()
                except OSError:
                    pass

        # Wait for thread to finish (5 second timeout)
        if self.isRunning():
            if not self.wait(5000):
                print(f"Warning: Network worker did not stop cleanly for "
                      f"{self.config.host}:{self.config.port}")

    def write(self, data: bytes):
        """
        Queue data to be written to the socket.

        Args:
            data: Bytes to send
        """
        if self.running:
            self.write_queue.put(data)

    def _drain_write_queue(self) -> int:
        """
        Get count of pending writes (for shutdown warnings).

        Returns:
            Number of items remaining in write queue
        """
        return self.write_queue.qsize()
```

### Acceptance Criteria

- [ ] Base class created with all required signals
- [ ] Signal names match `SerialWorker` exactly
- [ ] `stop()` method handles socket cleanup
- [ ] `write()` method queues data
- [ ] Abstract `run()` raises NotImplementedError

---

## Task A.2: Implement TCPClientWorker

### Objective

Implement TCP client that connects to a remote server.

### File

`core/network_worker.py` (append to file)

### Code

```python
class TCPClientWorker(NetworkWorker):
    """
    TCP Client connection handler.

    Connects to a remote server and maintains bidirectional communication.
    """

    def run(self):
        """Main thread loop for TCP client."""
        try:
            # Create TCP socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(self.config.timeout)

            # Configure TCP options
            if self.config.nodelay:
                self.socket.setsockopt(
                    socket.IPPROTO_TCP, socket.TCP_NODELAY, 1
                )

            # Configure keepalive if enabled
            if self.config.keepalive:
                self.socket.setsockopt(
                    socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1
                )
                # Windows-specific keepalive configuration
                try:
                    # keepalive_idle, keepalive_interval, keepalive_count
                    self.socket.ioctl(
                        socket.SIO_KEEPALIVE_VALS,
                        (1, self.config.keepalive_interval * 1000, 1000)
                    )
                except (AttributeError, OSError):
                    pass  # Not on Windows or not supported

            # Connect to server
            self.socket.connect((self.config.host, self.config.port))

            # Switch to non-blocking for main loop
            self.socket.setblocking(False)

            self.running = True
            self.connectionStateChanged.emit(True)

            # Main communication loop
            while self.running and not self._stop_event.is_set():
                # === READ ===
                try:
                    data = self.socket.recv(self.config.buffer_size)
                    if data:
                        self.dataReceived.emit(data)
                    elif data == b'':
                        # Empty bytes = connection closed by remote
                        self.errorOccurred.emit("Connection closed by remote host")
                        break
                except BlockingIOError:
                    pass  # No data available, continue
                except socket.timeout:
                    pass  # Timeout, continue
                except ConnectionResetError:
                    self.errorOccurred.emit("Connection reset by remote host")
                    break

                # === WRITE ===
                try:
                    while not self.write_queue.empty() and self.running:
                        data = self.write_queue.get_nowait()
                        self.socket.sendall(data)
                except queue.Empty:
                    pass
                except (socket.error, OSError) as e:
                    self.errorOccurred.emit(f"Send error: {e}")
                    # Continue running - send errors may be recoverable

                # Small sleep to prevent CPU spinning
                self._stop_event.wait(0.01)

        except socket.timeout:
            self.errorOccurred.emit(
                f"Connection timeout to {self.config.host}:{self.config.port}"
            )
        except ConnectionRefusedError:
            self.errorOccurred.emit(
                f"Connection refused by {self.config.host}:{self.config.port}"
            )
        except socket.gaierror as e:
            self.errorOccurred.emit(f"DNS resolution failed for {self.config.host}: {e}")
        except OSError as e:
            self.errorOccurred.emit(f"Network error: {e}")
        except Exception as e:
            self.errorOccurred.emit(f"Unexpected error: {e}")
        finally:
            # Cleanup
            if self.socket:
                try:
                    self.socket.close()
                except OSError:
                    pass
            self.connectionStateChanged.emit(False)

            # Warn about unsent data
            pending = self._drain_write_queue()
            if pending > 0:
                print(f"Warning: {pending} items not sent - connection closed")
```

### Test Cases

```python
# Test 1: Connect to echo server
config = NetworkConfig(host="127.0.0.1", port=7, protocol="TCP", mode="client")
worker = TCPClientWorker(config)

# Test 2: Connection refused handling
config = NetworkConfig(host="127.0.0.1", port=59999, protocol="TCP", mode="client")
# Should emit errorOccurred with "Connection refused"

# Test 3: DNS failure handling
config = NetworkConfig(host="invalid.host.example", port=80, protocol="TCP", mode="client")
# Should emit errorOccurred with "DNS resolution failed"
```

### Acceptance Criteria

- [ ] Connects to remote TCP server
- [ ] Emits `connectionStateChanged(True)` on successful connect
- [ ] Emits `dataReceived` when data arrives
- [ ] Sends queued data via `write()`
- [ ] Handles connection refused gracefully
- [ ] Handles DNS failures gracefully
- [ ] Handles remote disconnect gracefully
- [ ] Emits `connectionStateChanged(False)` on disconnect
- [ ] TCP keepalive works when enabled

---

## Task A.3: Implement TCPServerWorker

### Objective

Implement TCP server that listens for incoming connections.

### File

`core/network_worker.py` (append to file)

### Code

```python
class TCPServerWorker(NetworkWorker):
    """
    TCP Server (listen mode) handler.

    Listens for incoming connections and handles one client at a time.
    """

    def run(self):
        """Main thread loop for TCP server."""
        try:
            # Create server socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.socket.settimeout(1.0)  # Timeout for accept() to allow stop checking

            # Bind and listen
            bind_address = self.config.host if self.config.host else "0.0.0.0"
            self.socket.bind((bind_address, self.config.port))
            self.socket.listen(1)  # Single client queue

            self.running = True
            self.connectionStateChanged.emit(True)  # Listening is "connected"

            # Accept loop - wait for clients
            while self.running and not self._stop_event.is_set():
                try:
                    self.client_socket, client_addr = self.socket.accept()
                    client_str = f"{client_addr[0]}:{client_addr[1]}"

                    # Configure client socket
                    self.client_socket.setblocking(False)
                    if self.config.nodelay:
                        self.client_socket.setsockopt(
                            socket.IPPROTO_TCP, socket.TCP_NODELAY, 1
                        )

                    self.clientConnected.emit(client_str)

                    # Handle this client until disconnect
                    self._handle_client(client_str)

                    self.clientDisconnected.emit(client_str)

                    # Close client socket
                    try:
                        self.client_socket.close()
                    except OSError:
                        pass
                    self.client_socket = None

                except socket.timeout:
                    continue  # Keep waiting for connections
                except OSError as e:
                    if self.running:  # Only report if not shutting down
                        self.errorOccurred.emit(f"Accept error: {e}")

        except OSError as e:
            error_str = str(e).lower()
            if "address already in use" in error_str:
                self.errorOccurred.emit(
                    f"Port {self.config.port} is already in use"
                )
            elif "permission denied" in error_str:
                self.errorOccurred.emit(
                    f"Permission denied for port {self.config.port} "
                    f"(try a port > 1024)"
                )
            else:
                self.errorOccurred.emit(f"Server error: {e}")
        except Exception as e:
            self.errorOccurred.emit(f"Unexpected server error: {e}")
        finally:
            # Cleanup
            for sock in [self.client_socket, self.socket]:
                if sock:
                    try:
                        sock.close()
                    except OSError:
                        pass
            self.connectionStateChanged.emit(False)

    def _handle_client(self, client_str: str):
        """
        Handle communication with connected client.

        Args:
            client_str: Client address string for logging
        """
        while self.running and not self._stop_event.is_set():
            # === READ from client ===
            try:
                data = self.client_socket.recv(self.config.buffer_size)
                if data:
                    self.dataReceived.emit(data)
                elif data == b'':
                    # Client disconnected
                    break
            except BlockingIOError:
                pass  # No data available
            except (ConnectionResetError, ConnectionAbortedError):
                break  # Client disconnected abruptly
            except socket.error:
                break  # Other socket error

            # === WRITE to client ===
            try:
                while not self.write_queue.empty() and self.running:
                    data = self.write_queue.get_nowait()
                    self.client_socket.sendall(data)
            except queue.Empty:
                pass
            except (socket.error, OSError) as e:
                self.errorOccurred.emit(f"Send to client error: {e}")
                break

            # Small sleep to prevent CPU spinning
            self._stop_event.wait(0.01)
```

### Test Cases

```python
# Test 1: Start server and accept connection
config = NetworkConfig(host="0.0.0.0", port=5000, protocol="TCP", mode="server")
worker = TCPServerWorker(config)
# Connect with: nc localhost 5000

# Test 2: Port already in use
# Start two servers on same port - second should emit error

# Test 3: Client disconnect handling
# Connect, then close client - should emit clientDisconnected
```

### Acceptance Criteria

- [ ] Binds to specified port
- [ ] Emits `connectionStateChanged(True)` when listening
- [ ] Accepts incoming connections
- [ ] Emits `clientConnected` with client address
- [ ] Receives data from client, emits `dataReceived`
- [ ] Sends queued data to client
- [ ] Handles client disconnect, emits `clientDisconnected`
- [ ] Handles "port in use" error gracefully
- [ ] Handles permission denied for low ports
- [ ] Continues listening after client disconnects

---

## Task A.4: Implement UDPWorker

### Objective

Implement UDP sender/receiver for connectionless communication.

### File

`core/network_worker.py` (append to file)

### Code

```python
class UDPWorker(NetworkWorker):
    """
    UDP sender/receiver handler.

    Handles connectionless UDP communication with optional broadcast
    and multicast support.
    """

    def run(self):
        """Main thread loop for UDP."""
        try:
            # Create UDP socket
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self.socket.settimeout(0.1)  # Short timeout for recv

            # Enable broadcast if configured
            if self.config.broadcast:
                self.socket.setsockopt(
                    socket.SOL_SOCKET, socket.SO_BROADCAST, 1
                )

            # Bind for receiving
            if self.config.mode == 'server':
                # Server mode: bind to specific port to receive
                bind_addr = self.config.host if self.config.host else "0.0.0.0"
                self.socket.bind((bind_addr, self.config.port))
            else:
                # Client mode: bind to any available port
                self.socket.bind(('', 0))

            # Join multicast group if configured
            if self.config.multicast_group:
                try:
                    import struct
                    mreq = struct.pack(
                        '4sl',
                        socket.inet_aton(self.config.multicast_group),
                        socket.INADDR_ANY
                    )
                    self.socket.setsockopt(
                        socket.IPPROTO_IP,
                        socket.IP_ADD_MEMBERSHIP,
                        mreq
                    )
                    # Set multicast TTL
                    self.socket.setsockopt(
                        socket.IPPROTO_IP,
                        socket.IP_MULTICAST_TTL,
                        self.config.multicast_ttl
                    )
                except OSError as e:
                    self.errorOccurred.emit(f"Multicast setup failed: {e}")

            self.running = True
            self.connectionStateChanged.emit(True)

            # Track last receive address for display
            last_recv_addr = None

            # Main communication loop
            while self.running and not self._stop_event.is_set():
                # === RECEIVE datagrams ===
                try:
                    data, addr = self.socket.recvfrom(self.config.buffer_size)
                    if data:
                        self.dataReceived.emit(data)
                        # Optionally track source address
                        if addr != last_recv_addr:
                            last_recv_addr = addr
                except socket.timeout:
                    pass  # No data, continue
                except OSError as e:
                    if self.running:
                        self.errorOccurred.emit(f"Receive error: {e}")

                # === SEND datagrams ===
                try:
                    while not self.write_queue.empty() and self.running:
                        data = self.write_queue.get_nowait()
                        target = (self.config.host, self.config.port)
                        self.socket.sendto(data, target)
                except queue.Empty:
                    pass
                except OSError as e:
                    self.errorOccurred.emit(f"Send error: {e}")

                # Small sleep to prevent CPU spinning
                self._stop_event.wait(0.01)

        except OSError as e:
            error_str = str(e).lower()
            if "address already in use" in error_str:
                self.errorOccurred.emit(
                    f"UDP port {self.config.port} is already in use"
                )
            else:
                self.errorOccurred.emit(f"UDP error: {e}")
        except Exception as e:
            self.errorOccurred.emit(f"Unexpected UDP error: {e}")
        finally:
            if self.socket:
                try:
                    self.socket.close()
                except OSError:
                    pass
            self.connectionStateChanged.emit(False)
```

### Test Cases

```python
# Test 1: UDP echo
config = NetworkConfig(host="127.0.0.1", port=5000, protocol="UDP", mode="client")
worker = UDPWorker(config)

# Test 2: UDP server (receiver)
config = NetworkConfig(host="0.0.0.0", port=5000, protocol="UDP", mode="server")

# Test 3: Broadcast
config = NetworkConfig(host="255.255.255.255", port=5000, protocol="UDP",
                       mode="client", broadcast=True)
```

### Acceptance Criteria

- [ ] Creates UDP socket
- [ ] Binds to port in server mode
- [ ] Sends datagrams to target in client mode
- [ ] Receives datagrams, emits `dataReceived`
- [ ] Broadcast mode works
- [ ] Multicast join works (optional)
- [ ] Handles port-in-use error
- [ ] Emits `connectionStateChanged` appropriately

---

## Task A.5: Create Factory Function

### Objective

Create a factory function that returns the appropriate worker type based on configuration.

### File

`core/network_worker.py` (append to file)

### Code

```python
def create_network_worker(config: 'NetworkConfig') -> NetworkWorker:
    """
    Factory function to create appropriate network worker based on configuration.

    Args:
        config: NetworkConfig specifying protocol, mode, host, port

    Returns:
        Appropriate NetworkWorker subclass instance

    Raises:
        ValueError: If protocol or mode is invalid

    Example:
        config = NetworkConfig(host="127.0.0.1", port=5000,
                              protocol="TCP", mode="client")
        worker = create_network_worker(config)
        worker.start()
    """
    if config.protocol == 'TCP':
        if config.mode == 'server':
            return TCPServerWorker(config)
        elif config.mode == 'client':
            return TCPClientWorker(config)
        else:
            raise ValueError(f"Invalid TCP mode: {config.mode}")
    elif config.protocol == 'UDP':
        return UDPWorker(config)
    else:
        raise ValueError(f"Invalid protocol: {config.protocol}")


# Export public API
__all__ = [
    'NetworkWorker',
    'TCPClientWorker',
    'TCPServerWorker',
    'UDPWorker',
    'create_network_worker'
]
```

### Acceptance Criteria

- [ ] Factory function returns correct worker type
- [ ] Raises `ValueError` for invalid protocol
- [ ] Raises `ValueError` for invalid mode
- [ ] `__all__` exports all public classes

---

## Stream Deliverables

After completing Stream A, the following should be ready:

| Item | File | Status |
|------|------|--------|
| NetworkWorker base class | `core/network_worker.py` | |
| TCPClientWorker | `core/network_worker.py` | |
| TCPServerWorker | `core/network_worker.py` | |
| UDPWorker | `core/network_worker.py` | |
| Factory function | `core/network_worker.py` | |
| Unit tests | `tests/test_network_worker.py` | |

---

## Testing Notes

### Manual Testing with netcat

```bash
# TCP Server test
python -c "
from core.network_config import NetworkConfig
from core.network_worker import create_network_worker
config = NetworkConfig(host='0.0.0.0', port=5000, protocol='TCP', mode='server')
worker = create_network_worker(config)
worker.dataReceived.connect(lambda d: print(f'Received: {d}'))
worker.start()
input('Press Enter to stop...')
worker.stop()
"

# Connect with: nc localhost 5000
```

### Automated Testing

See [Testing Requirements](../reference/testing-requirements.md) for unit test specifications.

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Interface Definitions](../reference/interface-definitions.md) - Signal contracts
- [Stream C: Config Layer](./stream-c-config-data.md) - NetworkConfig class
- [Phase 4: Integration](./phase-4-integration.md) - Wiring up components

---

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Phase 0](./phase-0-project-setup.md) | [Stream B →](./stream-b-ui-layer.md)
