# Kickoff Prompt: Stream A - Core Network Layer

## Context

You are implementing Stream A of the Network Terminal project. Phase 0 is complete - the project skeleton exists at `/home/user/SerialTerminal/NetworkTerminal/`.

## Your Task

Implement the NetworkWorker classes that handle TCP and UDP communication. These must match the SerialWorker signal interface for drop-in UI compatibility.

## Branch

Work on branch: `dev/network-terminal` (or `claude/setup-network-terminal-SRbHj`)

## Documentation

**Detailed spec**: `/home/user/SerialTerminal/docs/phases/stream-a-core-network.md` - READ THIS FIRST

## Overview

| Attribute | Value |
|-----------|-------|
| **Stream** | A - Core Network |
| **File** | `NetworkTerminal/core/network_worker.py` |
| **Tasks** | 5 |
| **New Lines** | ~400 |
| **Dependencies** | NetworkConfig (Stream C) - use TYPE_CHECKING import |

## Critical Interface Contract

Your workers MUST emit these signals (same as SerialWorker):
```python
dataReceived = pyqtSignal(bytes)           # Emitted when data arrives
errorOccurred = pyqtSignal(str)            # Emitted on errors
connectionStateChanged = pyqtSignal(bool)  # True=connected, False=disconnected
```

## Tasks

### Task A.1: Create NetworkWorker Base Class
- Abstract base class inheriting from QThread
- Define all signals
- Implement `stop()` and `write()` methods
- Abstract `run()` method

### Task A.2: Implement TCPClientWorker
- Connect to remote TCP server
- Non-blocking read/write loop
- Handle connection errors (refused, timeout, DNS failure)
- Support TCP_NODELAY and keepalive options

### Task A.3: Implement TCPServerWorker
- Listen for incoming connections
- Accept one client at a time
- Handle "port in use" errors
- Emit clientConnected/clientDisconnected signals

### Task A.4: Implement UDPWorker
- Connectionless datagram send/receive
- Support broadcast mode
- Support multicast (optional)

### Task A.5: Create Factory Function
```python
def create_network_worker(config: 'NetworkConfig') -> NetworkWorker:
    """Factory function to create appropriate worker based on config"""
```

## Key Implementation Notes

1. **Use `TYPE_CHECKING` for NetworkConfig import** (avoids circular import):
```python
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from .network_config import NetworkConfig
```

2. **Socket timeout handling** - Use short timeouts and `_stop_event.wait()` for responsive shutdown

3. **Error handling** - Catch and emit specific errors (ConnectionRefused, timeout, DNS failure)

## Testing

Test manually with netcat:
```bash
# TCP Server test
nc localhost 5000

# TCP Client test (after starting a server)
nc -l 5000
```

## Acceptance Criteria

- [ ] NetworkWorker base class with required signals
- [ ] TCPClientWorker connects, sends, receives, handles errors
- [ ] TCPServerWorker listens, accepts, handles clients
- [ ] UDPWorker sends/receives datagrams
- [ ] Factory function returns correct worker type
- [ ] All classes use consistent error handling

## Files to Read First

```bash
cat /home/user/SerialTerminal/docs/phases/stream-a-core-network.md
cat /home/user/SerialTerminal/NetworkTerminal/core/network_worker.py  # Current placeholder
```

Commit your changes when complete.
