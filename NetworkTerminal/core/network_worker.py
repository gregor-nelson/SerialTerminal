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
                # Platform-specific keepalive configuration
                try:
                    # Windows-specific keepalive configuration
                    # keepalive_idle, keepalive_interval, keepalive_count
                    self.socket.ioctl(
                        socket.SIO_KEEPALIVE_VALS,
                        (1, self.config.keepalive_interval * 1000, 1000)
                    )
                except (AttributeError, OSError):
                    # Not on Windows - try Linux/macOS options
                    try:
                        self.socket.setsockopt(
                            socket.IPPROTO_TCP, socket.TCP_KEEPIDLE,
                            self.config.keepalive_interval
                        )
                        self.socket.setsockopt(
                            socket.IPPROTO_TCP, socket.TCP_KEEPINTVL,
                            self.config.keepalive_interval
                        )
                        self.socket.setsockopt(
                            socket.IPPROTO_TCP, socket.TCP_KEEPCNT, 3
                        )
                    except (AttributeError, OSError):
                        pass  # Not supported on this platform

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
            self.running = False
            self.connectionStateChanged.emit(False)

            # Warn about unsent data
            pending = self._drain_write_queue()
            if pending > 0:
                print(f"Warning: {pending} items not sent - connection closed")


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
            self.running = False
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
            self.running = False
            self.connectionStateChanged.emit(False)


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
