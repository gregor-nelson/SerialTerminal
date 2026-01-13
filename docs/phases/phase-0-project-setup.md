# Phase 0: Project Setup

[← Back to Master PRD](../PRD_NetworkTerminal.md)

---

## Overview

| Attribute | Value |
|-----------|-------|
| **Phase** | 0 - Prerequisites |
| **Owner** | Any single developer |
| **Dependencies** | None |
| **Estimated Effort** | 4 tasks |
| **Status** | Not Started |

---

## Purpose

Create the NetworkTerminal project skeleton with all reusable components copied from SerialTerminal. This phase must be completed before any parallel development can begin.

---

## Task Checklist

- [ ] [Task 0.1: Create Project Structure](#task-01-create-project-structure)
- [ ] [Task 0.2: Copy Reusable Files](#task-02-copy-reusable-files)
- [ ] [Task 0.3: Create requirements.txt](#task-03-create-requirementstxt)
- [ ] [Task 0.4: Verify Skeleton Runs](#task-04-verify-skeleton-runs)

---

## Task 0.1: Create Project Structure

### Objective

Create the complete directory structure for NetworkTerminal project.

### Directory Structure

```
NetworkTerminal/
├── main.py                           # Entry point (placeholder)
├── constants.py                      # Will be copied
├── build.py                          # Will be copied
├── requirements.txt                  # Will be created
│
├── core/
│   ├── __init__.py                   # Empty init
│   ├── core.py                       # Will contain extracted classes
│   ├── network_config.py             # Placeholder for Stream C
│   ├── network_worker.py             # Placeholder for Stream A
│   └── connection_manager.py         # Placeholder for Stream C
│
├── ui/
│   ├── __init__.py
│   ├── dialogs/
│   │   ├── __init__.py
│   │   ├── terminal_dialog.py        # Placeholder for Stream B
│   │   └── connection_dialog.py      # Placeholder for Stream B
│   │
│   ├── windows/
│   │   ├── __init__.py
│   │   └── terminal_formatter.py     # Will be copied
│   │
│   ├── components/
│   │   ├── __init__.py
│   │   └── ribbon_toolbar.py         # Will be copied
│   │
│   ├── common/
│   │   ├── __init__.py
│   │   └── icons.py                  # Will be copied
│   │
│   └── resources.py                  # Will be copied
│
├── assets/                           # Will be copied entirely
│   ├── fonts/
│   │   ├── JetBrainsMono/
│   │   └── Poppins/
│   └── icons/
│
├── tests/
│   └── __init__.py
│
└── docs/                             # Documentation (optional copy)
```

### Commands

```bash
# Create all directories
mkdir -p NetworkTerminal/{core,ui/{dialogs,windows,components,common},assets,tests,docs}

# Create __init__.py files
touch NetworkTerminal/__init__.py
touch NetworkTerminal/core/__init__.py
touch NetworkTerminal/ui/__init__.py
touch NetworkTerminal/ui/dialogs/__init__.py
touch NetworkTerminal/ui/windows/__init__.py
touch NetworkTerminal/ui/components/__init__.py
touch NetworkTerminal/ui/common/__init__.py
touch NetworkTerminal/tests/__init__.py
```

### Acceptance Criteria

- [ ] All directories created
- [ ] All `__init__.py` files present
- [ ] Directory structure matches specification

---

## Task 0.2: Copy Reusable Files

### Objective

Copy all files that can be reused without modification from SerialTerminal.

### Files to Copy (No Modifications)

| Source | Destination | Size |
|--------|-------------|------|
| `SerialTerminal/ui/windows/terminal_formatter.py` | `NetworkTerminal/ui/windows/terminal_formatter.py` | 344 lines |
| `SerialTerminal/ui/components/ribbon_toolbar.py` | `NetworkTerminal/ui/components/ribbon_toolbar.py` | ~150 lines |
| `SerialTerminal/ui/common/icons.py` | `NetworkTerminal/ui/common/icons.py` | ~100 lines |
| `SerialTerminal/ui/resources.py` | `NetworkTerminal/ui/resources.py` | ~200 lines |
| `SerialTerminal/assets/*` | `NetworkTerminal/assets/*` | Directory |

### Files to Copy with Modifications

| Source | Destination | Modifications |
|--------|-------------|---------------|
| `SerialTerminal/constants.py` | `NetworkTerminal/constants.py` | Change `AppInfo.NAME` to "Network Terminal" |
| `SerialTerminal/build.py` | `NetworkTerminal/build.py` | Change output name to "NetworkTerminal" |

### Modified constants.py

```python
#!/usr/bin/env python3
"""
Network Terminal Application Constants
"""

class TerminalColors:
    """Terminal text formatting colors - reused from Serial Terminal"""
    INCOMING = "#66bb6a"      # Green for incoming data
    OUTGOING = "#4dc3ff"      # Blue for outgoing data
    TIMESTAMP = "#999999"     # Gray for timestamps
    ERROR = "#ff6b6b"         # Red for errors
    STATUS = "#ffffff"        # White for status
    WARNING = "#cccccc"       # Light gray for warnings
    DEFAULT = "#e0e0e0"       # Default terminal text


class AppInfo:
    """Application metadata"""
    NAME = "Network Terminal"           # CHANGED
    VERSION = "1.0"
    ORG_NAME = "Network Terminal"       # CHANGED
    ORG_DOMAIN = "networkterminal.local"  # CHANGED
```

### Commands

```bash
# Copy unchanged files
cp SerialTerminal/ui/windows/terminal_formatter.py NetworkTerminal/ui/windows/
cp SerialTerminal/ui/components/ribbon_toolbar.py NetworkTerminal/ui/components/
cp SerialTerminal/ui/common/icons.py NetworkTerminal/ui/common/
cp SerialTerminal/ui/resources.py NetworkTerminal/ui/

# Copy assets directory
cp -r SerialTerminal/assets/* NetworkTerminal/assets/

# Copy and modify constants.py (manual edit required)
cp SerialTerminal/constants.py NetworkTerminal/constants.py
# Edit AppInfo class as shown above

# Copy and modify build.py (manual edit required)
cp SerialTerminal/build.py NetworkTerminal/build.py
# Change output name
```

### Acceptance Criteria

- [ ] `terminal_formatter.py` copied unchanged
- [ ] `ribbon_toolbar.py` copied unchanged
- [ ] `icons.py` copied unchanged
- [ ] `resources.py` copied unchanged
- [ ] `assets/` directory copied with all fonts and icons
- [ ] `constants.py` modified with new app name
- [ ] `build.py` modified with new output name

---

## Task 0.3: Create requirements.txt

### Objective

Create a minimal requirements file for NetworkTerminal.

### File Content

```
# NetworkTerminal Requirements
# Note: No pyserial needed - network sockets are in Python stdlib

PyQt6>=6.4.0
```

### Acceptance Criteria

- [ ] `requirements.txt` created
- [ ] Only PyQt6 dependency (no pyserial)
- [ ] `pip install -r requirements.txt` succeeds

---

## Task 0.4: Verify Skeleton Runs

### Objective

Create a minimal `main.py` that proves all copied modules import correctly.

### Minimal main.py

```python
#!/usr/bin/env python3
"""
Network Terminal - Entry Point
Minimal version for Phase 0 verification
"""

import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel
from PyQt6.QtCore import Qt

# Verify all copied modules import correctly
from constants import AppInfo, TerminalColors
from ui.resources import resource_manager
from ui.windows.terminal_formatter import TerminalStreamFormatter
from ui.components.ribbon_toolbar import RibbonToolbar
from ui.common.icons import Icons


def main():
    """Main entry point for verification"""
    app = QApplication(sys.argv)
    app.setApplicationName(AppInfo.NAME)
    app.setOrganizationName(AppInfo.ORG_NAME)

    # Set Fusion style like Serial Terminal
    app.setStyle("Fusion")

    # Create a simple test window
    window = QMainWindow()
    window.setWindowTitle(f"{AppInfo.NAME} - Phase 0 Verification")
    window.setMinimumSize(800, 600)

    # Add a label to confirm it works
    label = QLabel("Phase 0 Complete!\n\nAll modules imported successfully.")
    label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    label.setFont(resource_manager.get_monospace_font(size=14))
    window.setCentralWidget(label)

    # Test that formatter initializes
    formatter = TerminalStreamFormatter()
    print(f"Formatter colors loaded: {len(formatter.colors)} types")
    print(f"NMEA colors loaded: {len(formatter.nmea_colors)} types")

    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
```

### Verification Steps

```bash
cd NetworkTerminal
python main.py
```

### Expected Output

1. Window appears with title "Network Terminal - Phase 0 Verification"
2. Console shows:
   ```
   Formatter colors loaded: 8 types
   NMEA colors loaded: 27 types
   ```
3. No import errors

### Acceptance Criteria

- [ ] `python main.py` runs without errors
- [ ] Window displays with correct title
- [ ] All imports succeed
- [ ] Formatter initializes with colors
- [ ] Fusion style applied

---

## Deliverables

After completing Phase 0, the following should be ready:

| Item | Status |
|------|--------|
| Project directory structure | |
| All `__init__.py` files | |
| `terminal_formatter.py` (copied) | |
| `ribbon_toolbar.py` (copied) | |
| `icons.py` (copied) | |
| `resources.py` (copied) | |
| `assets/` directory (copied) | |
| `constants.py` (modified) | |
| `build.py` (modified) | |
| `requirements.txt` (created) | |
| `main.py` (minimal verification) | |
| Verification test passed | |

---

## Next Steps

After Phase 0 is complete, the following streams can begin **in parallel**:

| Stream | Document | Owner |
|--------|----------|-------|
| Stream A | [Core Network Layer](./stream-a-core-network.md) | Dev 1 |
| Stream B | [UI Layer](./stream-b-ui-layer.md) | Dev 2 |
| Stream C | [Config & Data Layer](./stream-c-config-data.md) | Dev 3 |

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Architecture Analysis](../reference/architecture-analysis.md)
- [Reusability Matrix](../reference/reusability-matrix.md)
- [File Structure](../reference/file-structure.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [Next: Stream A →](./stream-a-core-network.md)
