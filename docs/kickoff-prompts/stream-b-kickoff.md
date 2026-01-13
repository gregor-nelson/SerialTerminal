# Kickoff Prompt: Stream B - UI Layer

## Context

You are implementing Stream B of the Network Terminal project. Phase 0 is complete - the project skeleton exists at `/home/user/SerialTerminal/NetworkTerminal/`.

**Important**: Stream B depends on Stream A (NetworkWorker) and Stream C (NetworkConfig). Ensure those are complete first, or use placeholder imports.

## Your Task

Adapt the Serial Terminal UI components for network connections:
1. Extract SplitContainer (copy unchanged)
2. Create NetworkTerminalPane (adapt TerminalPane)
3. Create connection dialogs
4. Create NetworkMonitorWindow (main window)

## Branch

Work on branch: `dev/network-terminal` (or `claude/setup-network-terminal-SRbHj`)

## Documentation

**Detailed spec**: `/home/user/SerialTerminal/docs/phases/stream-b-ui-layer.md` - READ THIS FIRST

## Overview

| Attribute | Value |
|-----------|-------|
| **Stream** | B - UI Layer |
| **Files** | `NetworkTerminal/ui/dialogs/terminal_dialog.py`, `connection_dialog.py` |
| **Tasks** | 4 |
| **New Lines** | ~600 (adapted from ~1400 source) |
| **Dependencies** | Stream A, Stream C |

## Key Source File

**All UI adaptations reference**: `SerialTerminal/ui/dialogs/terminal_dialog.py` (2,598 lines)

Key classes in source:
- `SerialWorker` (lines 55-169) → Replace with NetworkWorker import
- `TerminalPane` (lines 172-1399) → Adapt to NetworkTerminalPane
- `SplitContainer` → Copy unchanged
- `SerialMonitorWindow` → Adapt to NetworkMonitorWindow

## Tasks

### Task B.1: Extract SplitContainer

**Copy unchanged** from SerialTerminal. This class manages recursive split pane layouts.

Location in source: Search for `class SplitContainer` in `terminal_dialog.py`

### Task B.2: Create NetworkTerminalPane

Adapt `TerminalPane` with these changes:

| Change | Action |
|--------|--------|
| Config type | Use `NetworkConfig` instead of `SerialConfig` |
| Worker | Use `create_network_worker()` factory |
| Baud rate logic | **REMOVE** entirely |
| COM port logic | **REMOVE** entirely |
| Status format | Change to "TCP 192.168.1.1:5000: Connected" |
| Context menu | Remove baud/port menus, add connection settings |

**Methods to COPY unchanged:**
- `_setup_ui()`, `eventFilter()`, `_handle_key_press()`
- `_flush_buffer()`, `_toggle_auto_scroll()`, `_toggle_hex_mode()`
- `_clear_terminal()`, font size methods, `_show_help()`

**Methods to REMOVE:**
- `_create_baud_rate_menu()`, `_create_com_port_menu()`
- `_set_baud_rate()`, `_set_com_port()`
- `_handle_encoding_error()`, baud rate detection methods

### Task B.3: Create Connection Dialogs

**File**: `NetworkTerminal/ui/dialogs/connection_dialog.py` (NEW)

Create:
1. `QuickConnectDialog` - Protocol/mode/host/port configuration
2. `ConnectionHistoryDialog` - Favorites and history list

### Task B.4: Create NetworkMonitorWindow

Adapt `SerialMonitorWindow` with:
- Window title: "Network Terminal"
- New connection: Opens `QuickConnectDialog`
- Remove: `get_connected_ports()`, virtual port manager references
- Keep: Tab management, split panes, keyboard shortcuts

## Key Implementation Notes

1. **Import NetworkWorker correctly:**
```python
from core.network_config import NetworkConfig
from core.network_worker import NetworkWorker, create_network_worker
```

2. **Remove baud rate detection** - The encoding error handling in `_on_data_received()` should be simplified (no baud rate suggestions)

3. **Status bar format change:**
```python
# Old (serial): "COM3 @ 115200: Connected | RX: 1KB | TX: 256B"
# New (network): "TCP 192.168.1.1:5000: Connected | RX: 1KB | TX: 256B"
```

## Reference Files

```bash
# Main source file to adapt
cat /home/user/SerialTerminal/ui/dialogs/terminal_dialog.py

# Stream documentation
cat /home/user/SerialTerminal/docs/phases/stream-b-ui-layer.md
```

## Acceptance Criteria

- [ ] SplitContainer copied and working
- [ ] NetworkTerminalPane renders terminal, handles data
- [ ] All baud rate/COM port logic removed
- [ ] QuickConnectDialog configures TCP/UDP connections
- [ ] NetworkMonitorWindow opens and manages tabs
- [ ] Context menu shows network-appropriate options
- [ ] Status bar shows network connection format

## Files to Read First

```bash
cat /home/user/SerialTerminal/docs/phases/stream-b-ui-layer.md
cat /home/user/SerialTerminal/ui/dialogs/terminal_dialog.py  # Source to adapt
cat /home/user/SerialTerminal/NetworkTerminal/ui/dialogs/terminal_dialog.py  # Current placeholder
```

Commit your changes when complete.
