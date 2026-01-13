# Phase 4: Integration

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Stream C](./stream-c-config-data.md)

---

## Overview

| Attribute | Value |
|-----------|-------|
| **Phase** | 4 - Integration |
| **Owner** | All Developers |
| **Dependencies** | Streams A, B, C complete |
| **Estimated Effort** | 3 tasks, ~50 new + ~100 adapted lines |
| **Status** | Not Started |

---

## Purpose

Wire up all components from the parallel streams into a working application. This phase connects:
- NetworkWorker (Stream A) ↔ NetworkTerminalPane (Stream B)
- NetworkConfig (Stream C) ↔ QuickConnectDialog (Stream B)
- ConnectionManager (Stream C) ↔ NetworkMonitorWindow (Stream B)

---

## Prerequisites

Before starting Phase 4, verify the following are complete:

| Stream | Deliverable | Verification |
|--------|-------------|--------------|
| Stream A | `NetworkWorker` classes | Unit tests pass |
| Stream A | `create_network_worker()` factory | Returns correct worker types |
| Stream B | `NetworkTerminalPane` | Renders terminal UI |
| Stream B | `QuickConnectDialog` | Shows protocol/mode options |
| Stream B | `NetworkMonitorWindow` | Window opens with tabs |
| Stream C | `NetworkConfig` | Serialization works |
| Stream C | `ConnectionManager` | Favorites persist |

---

## Task Checklist

- [ ] [Task 4.1: Wire Up Connection Flow](#task-41-wire-up-connection-flow)
- [ ] [Task 4.2: Integrate History/Favorites](#task-42-integrate-historyfavorites)
- [ ] [Task 4.3: Update main.py Entry Point](#task-43-update-mainpy-entry-point)

---

## Task 4.1: Wire Up Connection Flow

### Objective

Connect the dialog → config → worker → terminal pane flow.

### Integration Points

```
┌─────────────────────┐    ┌─────────────────────┐
│ QuickConnectDialog  │───►│   NetworkConfig     │
│  (User Input)       │    │  (Configuration)    │
└─────────────────────┘    └─────────────────────┘
                                    │
                                    ▼
┌─────────────────────┐    ┌─────────────────────┐
│ NetworkTerminalPane │◄───│  create_network_    │
│  (UI Display)       │    │  worker()           │
└─────────────────────┘    └─────────────────────┘
                                    │
                                    ▼
                           ┌─────────────────────┐
                           │  NetworkWorker      │
                           │  (Connection)       │
                           └─────────────────────┘
```

### Code Changes

#### NetworkMonitorWindow._on_new_connection()

```python
def _on_new_connection(self):
    """Handle new connection button - wire up complete flow"""
    from ui.dialogs.connection_dialog import QuickConnectDialog
    from core.network_config import NetworkConfig

    dialog = QuickConnectDialog(self)
    if dialog.exec() == QDialog.DialogCode.Accepted:
        config = dialog.get_config()
        self._create_pane_with_config(config)

def _create_pane_with_config(self, config: NetworkConfig):
    """Create a new terminal pane with the given configuration"""
    # Get or create container
    container = self._get_active_container()
    if not container:
        container = self._create_new_tab()

    # Create pane with config
    pane = NetworkTerminalPane(
        config=config,
        main_window=self,
        container=container
    )

    # Add to container
    container.add_pane(pane)

    # Connect automatically
    pane.connect()

    # Record in history
    if hasattr(self, 'connection_manager'):
        self.connection_manager.record_connection(config)
```

#### NetworkTerminalPane.connect() - Verify Integration

```python
def connect(self):
    """Connect using network worker - verify integration"""
    if self.network_worker:
        return  # Already connected

    # Create worker using factory (Stream A integration)
    from core.network_worker import create_network_worker

    self.network_worker = create_network_worker(self.config)

    # Wire up signals (must match exactly)
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

    # Server-specific signals (optional)
    if hasattr(self.network_worker, 'clientConnected'):
        self.network_worker.clientConnected.connect(
            self._on_client_connected,
            Qt.ConnectionType.QueuedConnection
        )
        self.network_worker.clientDisconnected.connect(
            self._on_client_disconnected,
            Qt.ConnectionType.QueuedConnection
        )

    # Start the worker
    self.network_worker.start()
```

### Verification Steps

```bash
# Test 1: TCP Client connection
python -c "
from PyQt6.QtWidgets import QApplication
import sys
app = QApplication(sys.argv)

from core.network_config import NetworkConfig
from core.network_worker import create_network_worker

config = NetworkConfig(host='127.0.0.1', port=5000, protocol='TCP', mode='client')
worker = create_network_worker(config)

print(f'Created worker type: {type(worker).__name__}')
assert type(worker).__name__ == 'TCPClientWorker'
print('Integration test PASSED')
"

# Test 2: Worker signals emit correctly
# Start a test server: nc -l 5000
# Run application and connect - verify dataReceived signal fires
```

### Acceptance Criteria

- [ ] QuickConnectDialog returns valid NetworkConfig
- [ ] NetworkConfig passed to create_network_worker()
- [ ] Worker signals connected to pane handlers
- [ ] Connection starts when pane.connect() called
- [ ] Data flows from worker → pane → terminal display

---

## Task 4.2: Integrate History/Favorites

### Objective

Connect ConnectionManager to the UI for favorites and history.

### Code Changes

#### NetworkMonitorWindow.__init__() - Add ConnectionManager

```python
def __init__(self):
    super().__init__()
    self.setWindowTitle("Network Terminal")

    # Initialize connection manager (Stream C integration)
    from core.connection_manager import ConnectionManager
    self.connection_manager = ConnectionManager()

    # ... rest of initialization
```

#### Add History Menu Action

```python
def _setup_menu_bar(self):
    """Setup menu bar with history/favorites"""
    menu_bar = self.menuBar()

    # File menu
    file_menu = menu_bar.addMenu("File")

    new_action = file_menu.addAction("New Connection")
    new_action.setShortcut("Ctrl+N")
    new_action.triggered.connect(self._on_new_connection)

    history_action = file_menu.addAction("Connection History...")
    history_action.triggered.connect(self._show_history_dialog)

    file_menu.addSeparator()

    # Recent connections submenu
    recent_menu = file_menu.addMenu("Recent Connections")
    self._populate_recent_menu(recent_menu)

    file_menu.addSeparator()

    exit_action = file_menu.addAction("Exit")
    exit_action.setShortcut("Ctrl+Q")
    exit_action.triggered.connect(self.close)

def _show_history_dialog(self):
    """Show connection history dialog"""
    from ui.dialogs.connection_dialog import ConnectionHistoryDialog

    dialog = ConnectionHistoryDialog(self.connection_manager, self)
    dialog.configSelected.connect(self._create_pane_with_config)
    dialog.exec()

def _populate_recent_menu(self, menu: QMenu):
    """Populate recent connections menu"""
    recent = self.connection_manager.get_recent_configs(limit=10)

    if not recent:
        no_recent = menu.addAction("No recent connections")
        no_recent.setEnabled(False)
        return

    for config in recent:
        action = menu.addAction(
            f"{config.protocol} {config.host}:{config.port}"
        )
        action.triggered.connect(
            lambda checked, c=config: self._create_pane_with_config(c)
        )
```

#### Add Favorite Toggle to Terminal Pane

```python
def _create_connection_settings_menu(self, menu: QMenu):
    """Create connection settings submenu with favorite toggle"""
    # Show current connection info
    info = menu.addAction(f"{self.config.protocol} {self.config.mode.title()}")
    info.setEnabled(False)

    menu.addSeparator()

    # Host/Port info
    host_action = menu.addAction(f"Host: {self.config.host}")
    host_action.setEnabled(False)

    port_action = menu.addAction(f"Port: {self.config.port}")
    port_action.setEnabled(False)

    menu.addSeparator()

    # Favorite toggle
    if self.main_window and hasattr(self.main_window, 'connection_manager'):
        is_favorite = self.main_window.connection_manager.is_favorite(self.config)
        favorite_action = menu.addAction(
            self.checkbox_icon(is_favorite),
            "Add to Favorites" if not is_favorite else "Remove from Favorites"
        )
        favorite_action.triggered.connect(self._toggle_favorite)

def _toggle_favorite(self):
    """Toggle favorite status for current connection"""
    if not self.main_window or not hasattr(self.main_window, 'connection_manager'):
        return

    manager = self.main_window.connection_manager

    if manager.is_favorite(self.config):
        manager.remove_favorite(self.config)
        self.formatter.append_status(
            self.terminal,
            "Removed from favorites",
            "status"
        )
    else:
        manager.add_favorite(self.config)
        self.formatter.append_status(
            self.terminal,
            "Added to favorites",
            "status"
        )
```

### Acceptance Criteria

- [ ] ConnectionManager initialized in main window
- [ ] History dialog shows past connections
- [ ] Recent connections menu populated
- [ ] Favorite toggle works from pane context menu
- [ ] History persists across application restarts

---

## Task 4.3: Update main.py Entry Point

### Objective

Update the main entry point to launch the fully integrated application.

### File

`main.py`

### Code

```python
#!/usr/bin/env python3
"""
Network Terminal - Entry Point

TCP/UDP Terminal Application
"""

import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

from constants import AppInfo
from ui.resources import resource_manager


def main():
    """Main entry point"""
    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName(AppInfo.NAME)
    app.setOrganizationName(AppInfo.ORG_NAME)
    app.setOrganizationDomain(AppInfo.ORG_DOMAIN)

    # Set Fusion style for consistent look
    app.setStyle("Fusion")

    # Apply dark palette (optional - matches Serial Terminal)
    from ui.resources import apply_dark_palette
    apply_dark_palette(app)

    # Import main window (after app created for QSettings)
    from ui.dialogs.terminal_dialog import NetworkMonitorWindow

    # Create and show main window
    window = NetworkMonitorWindow()
    window.show()

    # Start event loop
    return app.exec()


def apply_dark_palette(app: QApplication):
    """Apply dark color palette to application"""
    from PyQt6.QtGui import QPalette, QColor

    palette = QPalette()

    # Dark theme colors
    palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)

    app.setPalette(palette)


if __name__ == "__main__":
    sys.exit(main())
```

### Verification Steps

```bash
cd NetworkTerminal
python main.py
```

### Expected Behavior

1. Application window opens with title "Network Terminal"
2. Ribbon toolbar shows "New", "Refresh", "Connect", "Clear", "Settings"
3. Click "New" opens QuickConnectDialog
4. Select TCP Client, enter host/port, click Connect
5. Terminal pane shows connection status
6. Data received displays with NMEA color coding

### Acceptance Criteria

- [ ] Application starts without errors
- [ ] Window uses Fusion style with dark palette
- [ ] New connection dialog opens
- [ ] Connections can be established
- [ ] Data displays in terminal
- [ ] Application shuts down cleanly

---

## Integration Testing Checklist

After completing all tasks, verify end-to-end functionality:

### TCP Client Test

```bash
# Terminal 1: Start echo server
nc -l 5000

# Terminal 2: Run Network Terminal
python main.py
# Create TCP Client connection to 127.0.0.1:5000
# Type in nc terminal, verify data appears in Network Terminal
# Type in Network Terminal, verify data appears in nc
```

### TCP Server Test

```bash
# Terminal 1: Run Network Terminal
python main.py
# Create TCP Server on port 5000

# Terminal 2: Connect client
nc localhost 5000
# Type messages, verify bidirectional communication
```

### UDP Test

```bash
# Terminal 1: Run Network Terminal
python main.py
# Create UDP connection on port 5000

# Terminal 2: Send UDP packet
echo "Hello UDP" | nc -u localhost 5000
# Verify data appears in Network Terminal
```

---

## Phase Deliverables

After completing Phase 4, the following should be ready:

| Item | Status |
|------|--------|
| Connection flow working (dialog → config → worker → pane) | |
| ConnectionManager integrated with UI | |
| Favorites/History accessible from menu | |
| main.py launches full application | |
| TCP Client connections work | |
| TCP Server connections work | |
| UDP connections work | |
| Data displays correctly | |
| Clean shutdown | |

---

## Troubleshooting

### Common Integration Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| Worker signals not firing | Signal not connected | Check connect() calls with QueuedConnection |
| Config not passing to worker | Wrong import path | Verify import from `core.network_config` |
| History not persisting | QSettings not synced | Call `settings.sync()` after changes |
| Dark theme not applying | Palette set before app | Set palette after `QApplication()` creation |

---

## Next Steps

After Phase 4 is complete, proceed to:

- [Phase 5: Testing & Polish](./phase-5-testing.md)

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Stream A: Core Network](./stream-a-core-network.md)
- [Stream B: UI Layer](./stream-b-ui-layer.md)
- [Stream C: Config/Data](./stream-c-config-data.md)
- [Interface Definitions](../reference/interface-definitions.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Stream C](./stream-c-config-data.md) | [Phase 5 →](./phase-5-testing.md)
