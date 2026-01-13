# Kickoff Prompt: Phase 4 - Integration

## Context

You are implementing Phase 4 (Integration) of the Network Terminal project. Streams A, B, and C are complete - all core components exist. Your task is to wire them together into a fully working application.

**Repository**: `/home/user/SerialTerminal`
**Network Terminal Location**: `/home/user/SerialTerminal/NetworkTerminal/`

## Completed Work

| Stream | Description | Commit | Status |
|--------|-------------|--------|--------|
| Stream A | NetworkWorker TCP/UDP classes | `fef6a26` | ✅ Complete |
| Stream B | UI Layer (NetworkTerminalPane, dialogs) | `dc409b7` | ✅ Complete |
| Stream C | Config & Data Layer | `e2f380c` | ✅ Complete |

## Your Task

Wire up all components into a working application:

1. **Task 4.1**: Verify connection flow works (partially done in Stream B)
2. **Task 4.2**: Integrate History/Favorites with UI
3. **Task 4.3**: Finalize main.py entry point

## Branch

Create a new branch from `dev/network-terminal` or work on existing integration branch.

## Documentation

**Detailed spec**: `/home/user/SerialTerminal/docs/phases/phase-4-integration.md` - READ THIS FIRST

## Overview

| Attribute | Value |
|-----------|-------|
| **Phase** | 4 - Integration |
| **Files** | `terminal_dialog.py`, `main.py` |
| **Tasks** | 3 |
| **New Lines** | ~50-100 (mostly wiring existing code) |
| **Dependencies** | Streams A, B, C complete |

---

## Task 4.1: Verify Connection Flow (Quick Check)

Stream B already wired most of this. Verify these work:

```
QuickConnectDialog → NetworkConfig → create_network_worker() → NetworkTerminalPane
```

**Verification**:
```bash
# Terminal 1: Start test server
nc -l 5000

# Terminal 2: Run Network Terminal
cd /home/user/SerialTerminal/NetworkTerminal
python main.py

# Test: Create TCP Client to 127.0.0.1:5000
# Type in nc, verify data appears in Network Terminal
```

**If broken, check**:
- `NetworkTerminalPane.connect()` calls `create_network_worker(self.config)`
- Worker signals connected with `Qt.ConnectionType.QueuedConnection`

---

## Task 4.2: Integrate History/Favorites

This is the main work. Add ConnectionManager integration to the UI.

### 4.2.1: Add ConnectionManager to NetworkMonitorWindow

**File**: `NetworkTerminal/ui/dialogs/terminal_dialog.py`

In `NetworkMonitorWindow.__init__()`, add:

```python
def __init__(self):
    super().__init__()
    self.setWindowTitle("Network Terminal")

    # Add ConnectionManager integration
    from core.connection_manager import ConnectionManager
    self.connection_manager = ConnectionManager()

    # ... rest of existing __init__ code
```

### 4.2.2: Add Menu Bar with History

Add a menu bar to NetworkMonitorWindow:

```python
def _setup_ui(self):
    """Setup main window UI"""
    # Add menu bar BEFORE ribbon toolbar
    self._setup_menu_bar()

    # ... existing ribbon and tab setup code

def _setup_menu_bar(self):
    """Setup menu bar with File menu"""
    menu_bar = self.menuBar()

    # File menu
    file_menu = menu_bar.addMenu("&File")

    new_action = file_menu.addAction("&New Connection")
    new_action.setShortcut("Ctrl+N")
    new_action.triggered.connect(self._new_connection)

    file_menu.addSeparator()

    history_action = file_menu.addAction("Connection &History...")
    history_action.triggered.connect(self._show_history_dialog)

    # Recent connections submenu
    self.recent_menu = file_menu.addMenu("&Recent Connections")
    self._populate_recent_menu()

    file_menu.addSeparator()

    exit_action = file_menu.addAction("E&xit")
    exit_action.setShortcut("Ctrl+Q")
    exit_action.triggered.connect(self.close)

def _show_history_dialog(self):
    """Show connection history dialog"""
    from ui.dialogs.connection_dialog import ConnectionHistoryDialog

    dialog = ConnectionHistoryDialog(self.connection_manager, self)
    dialog.configSelected.connect(self._create_tab)
    dialog.exec()

def _populate_recent_menu(self):
    """Populate recent connections menu"""
    self.recent_menu.clear()

    recent = self.connection_manager.get_recent_configs(limit=10)

    if not recent:
        no_recent = self.recent_menu.addAction("No recent connections")
        no_recent.setEnabled(False)
        return

    for config in recent:
        action = self.recent_menu.addAction(
            f"{config.protocol} {config.host}:{config.port}"
        )
        action.triggered.connect(
            lambda checked, c=config: self._create_tab(c)
        )
```

### 4.2.3: Record Connections in History

In `_create_tab()`, add history recording:

```python
def _create_tab(self, config: NetworkConfig):
    """Create a new tab with split container"""
    # ... existing tab creation code ...

    # Record in history
    self.connection_manager.record_connection(config)

    # Refresh recent menu
    self._populate_recent_menu()
```

### 4.2.4: Add Favorite Toggle to Terminal Pane Context Menu

In `NetworkTerminalPane._create_connection_settings_menu()`:

```python
def _create_connection_settings_menu(self, menu: QMenu):
    """Create connection settings submenu with favorite toggle"""
    # Show current connection info
    info = menu.addAction(f"{self.config.protocol} {self.config.mode.title()}")
    info.setEnabled(False)

    menu.addSeparator()

    host_action = menu.addAction(f"Host: {self.config.host}")
    host_action.setEnabled(False)

    port_action = menu.addAction(f"Port: {self.config.port}")
    port_action.setEnabled(False)

    menu.addSeparator()

    # Favorite toggle
    if self.main_window and hasattr(self.main_window, 'connection_manager'):
        is_favorite = self.main_window.connection_manager.is_favorite(self.config)
        fav_text = "Remove from Favorites" if is_favorite else "Add to Favorites"
        favorite_action = menu.addAction(self.checkbox_icon(is_favorite), fav_text)
        favorite_action.triggered.connect(self._toggle_favorite)

def _toggle_favorite(self):
    """Toggle favorite status for current connection"""
    if not self.main_window or not hasattr(self.main_window, 'connection_manager'):
        return

    manager = self.main_window.connection_manager

    if manager.is_favorite(self.config):
        manager.remove_favorite(self.config)
        self.formatter.append_status(self.terminal, "Removed from favorites", "status")
    else:
        manager.add_favorite(self.config)
        self.formatter.append_status(self.terminal, "Added to favorites", "status")
```

---

## Task 4.3: Finalize main.py

Update `NetworkTerminal/main.py` with dark palette and proper initialization:

```python
#!/usr/bin/env python3
"""
Network Terminal - TCP/UDP Terminal Application
"""

import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPalette, QColor
from PyQt6.QtCore import Qt

from constants import AppInfo


def apply_dark_palette(app: QApplication):
    """Apply dark color palette to application"""
    palette = QPalette()

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


def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    app.setApplicationName(AppInfo.NAME)
    app.setOrganizationName(AppInfo.ORG_NAME)

    # Set Fusion style for consistent look
    app.setStyle("Fusion")

    # Apply dark palette
    apply_dark_palette(app)

    # Import main window (after app created for QSettings)
    from ui.dialogs.terminal_dialog import NetworkMonitorWindow

    # Create and show main window
    window = NetworkMonitorWindow()
    window.resize(1200, 800)
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
```

---

## Files to Modify

| File | Changes |
|------|---------|
| `ui/dialogs/terminal_dialog.py` | Add ConnectionManager, menu bar, favorite toggle |
| `main.py` | Add dark palette function |

## Key Files to Read First

```bash
# Existing implementations to understand
cat /home/user/SerialTerminal/NetworkTerminal/ui/dialogs/terminal_dialog.py
cat /home/user/SerialTerminal/NetworkTerminal/core/connection_manager.py
cat /home/user/SerialTerminal/NetworkTerminal/ui/dialogs/connection_dialog.py

# Phase documentation
cat /home/user/SerialTerminal/docs/phases/phase-4-integration.md
```

---

## Testing

After implementation, test with:

```bash
# Terminal 1: Start echo server
nc -l 5000

# Terminal 2: Run application
cd /home/user/SerialTerminal/NetworkTerminal
python main.py
```

**Test Checklist**:
- [ ] Application starts with dark theme
- [ ] File menu shows New Connection, History, Recent, Exit
- [ ] Ctrl+N opens QuickConnectDialog
- [ ] TCP Client connects to nc server
- [ ] Data flows bidirectionally
- [ ] Connection appears in Recent menu after connecting
- [ ] Right-click → Connection Settings → Add to Favorites works
- [ ] History dialog shows favorites and recent connections
- [ ] Selecting from history creates new connection
- [ ] History persists after app restart

---

## Acceptance Criteria

- [ ] ConnectionManager initialized in NetworkMonitorWindow
- [ ] Menu bar with File → New, History, Recent, Exit
- [ ] Recent connections menu auto-populates
- [ ] History dialog opens and shows connections
- [ ] Favorite toggle works in pane context menu
- [ ] Connections recorded in history on connect
- [ ] Dark theme applied consistently
- [ ] All existing functionality still works

---

## Commit Message Template

```
Phase 4: Integrate ConnectionManager with UI

- Add ConnectionManager to NetworkMonitorWindow
- Add File menu with History and Recent Connections
- Add favorite toggle to terminal pane context menu
- Record connections in history on connect
- Update main.py with dark palette
```

---

## After Completion

Once Phase 4 is complete, proceed to Phase 5 (Testing & Polish):
- Manual testing of all connection types
- Performance testing
- Bug fixes
- Build and packaging

See `/home/user/SerialTerminal/docs/phases/phase-5-testing.md`
