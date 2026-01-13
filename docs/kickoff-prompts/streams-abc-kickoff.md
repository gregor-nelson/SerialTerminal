# Kickoff Prompt: Network Terminal - Streams A, B, C (Parallel Development)

## Context

Phase 0 of the Network Terminal project is complete. The project skeleton has been created with all reusable components copied from Serial Terminal. Three parallel development streams can now begin.

**Repository**: `/home/user/SerialTerminal`
**Network Terminal Location**: `/home/user/SerialTerminal/NetworkTerminal/`
**Development Branch**: `dev/network-terminal` (or create from `claude/setup-network-terminal-SRbHj`)

## Phase 0 Completion Summary

The following has been completed:
- Project directory structure created
- Reusable components copied: `terminal_formatter.py`, `ribbon_toolbar.py`, `icons.py`, `resources.py`
- Assets copied (fonts, icons)
- `constants.py` modified with "Network Terminal" branding
- `build.py` modified for NetworkTerminal output
- `requirements.txt` created (PyQt6)
- Placeholder files created for all modules

## Documentation Location

All detailed documentation is in `/home/user/SerialTerminal/docs/`:
- Master PRD: `docs/PRD_NetworkTerminal.md`
- **Stream A**: `docs/phases/stream-a-core-network.md` - Core Network Layer
- **Stream B**: `docs/phases/stream-b-ui-layer.md` - UI Layer
- **Stream C**: `docs/phases/stream-c-config-data.md` - Config & Data Layer

## Choose Your Stream

These three streams can be developed **in parallel**. Pick one to implement:

---

### Stream A: Core Network Layer (Recommended First)

| Attribute | Value |
|-----------|-------|
| **Focus** | NetworkWorker classes for TCP/UDP communication |
| **Files** | `NetworkTerminal/core/network_worker.py` |
| **Tasks** | 5 tasks, ~400 lines |
| **Dependencies** | None (can start immediately) |

**Tasks:**
1. Task A.1: Create NetworkWorker base class with signals matching SerialWorker
2. Task A.2: Implement TCPClientWorker (connect to remote server)
3. Task A.3: Implement TCPServerWorker (listen for connections)
4. Task A.4: Implement UDPWorker (connectionless datagrams)
5. Task A.5: Create `create_network_worker()` factory function

**Critical**: Worker signals MUST match SerialWorker interface:
```python
dataReceived = pyqtSignal(bytes)
errorOccurred = pyqtSignal(str)
connectionStateChanged = pyqtSignal(bool)
```

---

### Stream B: UI Layer

| Attribute | Value |
|-----------|-------|
| **Focus** | Terminal panes and connection dialogs |
| **Files** | `NetworkTerminal/ui/dialogs/terminal_dialog.py`, `connection_dialog.py` |
| **Tasks** | 4 tasks, ~600 lines |
| **Dependencies** | Uses NetworkConfig (Stream C) and NetworkWorker (Stream A) |

**Tasks:**
1. Task B.1: Extract SplitContainer from SerialTerminal (copy unchanged)
2. Task B.2: Create NetworkTerminalPane (adapt TerminalPane, remove baud rate logic)
3. Task B.3: Create QuickConnectDialog and ConnectionHistoryDialog
4. Task B.4: Create NetworkMonitorWindow (main application window)

**Key Source File**: `SerialTerminal/ui/dialogs/terminal_dialog.py` (2,598 lines)

---

### Stream C: Config & Data Layer (Recommended First)

| Attribute | Value |
|-----------|-------|
| **Focus** | NetworkConfig dataclass, ConnectionManager, utility classes |
| **Files** | `NetworkTerminal/core/network_config.py`, `connection_manager.py`, `core.py` |
| **Tasks** | 3 tasks, ~250 lines |
| **Dependencies** | None (can start immediately) |

**Tasks:**
1. Task C.1: Create NetworkConfig dataclass with validation
2. Task C.2: Create ConnectionManager for favorites/history persistence
3. Task C.3: Extract SettingsManager and ResponsiveWindowManager utilities

---

## Quick Start Commands

```bash
# Navigate to project
cd /home/user/SerialTerminal

# Ensure on correct branch
git checkout dev/network-terminal  # or claude/setup-network-terminal-SRbHj

# Verify Phase 0 structure
ls NetworkTerminal/
ls NetworkTerminal/core/
ls NetworkTerminal/ui/

# Read your stream's documentation
cat docs/phases/stream-a-core-network.md   # For Stream A
cat docs/phases/stream-b-ui-layer.md       # For Stream B
cat docs/phases/stream-c-config-data.md    # For Stream C
```

## Recommended Development Order

**Option 1: Sequential (Single Developer)**
1. Stream C (NetworkConfig) - Creates the config foundation
2. Stream A (NetworkWorker) - Uses NetworkConfig, creates communication layer
3. Stream B (UI) - Uses both NetworkConfig and NetworkWorker

**Option 2: Parallel (Multiple Sessions)**
- Stream A and Stream C can start immediately (no dependencies)
- Stream B requires Streams A and C to be complete

## Reference: Serial Terminal Source Files

When adapting code, reference these SerialTerminal files:
- `SerialTerminal/ui/dialogs/terminal_dialog.py` - Main UI classes (TerminalPane, SerialWorker, SerialMonitorWindow)
- `SerialTerminal/core/core.py` - Core utilities (SettingsManager, ResponsiveWindowManager)
- `SerialTerminal/constants.py` - Color definitions

## Acceptance Criteria (All Streams)

After completing all three streams:
- [ ] `NetworkConfig` can create/validate/serialize connection configs
- [ ] `NetworkWorker` classes (TCP/UDP) implement SerialWorker signal interface
- [ ] `ConnectionManager` persists favorites and history
- [ ] UI components render and respond to user input
- [ ] No import errors when running `python NetworkTerminal/main.py`

## After Stream Completion

Once all streams are complete, Phase 4 (Integration) will wire everything together:
- Connect NetworkMonitorWindow with NetworkWorker
- Add keyboard shortcuts and toolbar actions
- Implement connection history tracking
- Final testing and polish

---

**Read the detailed stream documentation before starting!**
