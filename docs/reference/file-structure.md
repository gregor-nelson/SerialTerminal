# File Structure

[← Back to Master PRD](../PRD_NetworkTerminal.md)

---

## Overview

This document defines the target file structure for the Network Terminal application. Use this as a reference when creating and organizing files.

---

## Complete Project Structure

```
NetworkTerminal/
│
├── main.py                           # Application entry point
├── constants.py                      # Application constants and colors
├── build.py                          # PyInstaller build script
├── requirements.txt                  # Python dependencies
├── README.md                         # User documentation
├── LICENSE                           # License file
│
├── core/                             # Core business logic
│   ├── __init__.py
│   ├── core.py                       # Utility classes (WindowConfig, SettingsManager, etc.)
│   ├── network_config.py             # NetworkConfig dataclass
│   ├── network_worker.py             # NetworkWorker classes (TCP/UDP)
│   └── connection_manager.py         # Favorites and history persistence
│
├── ui/                               # User interface components
│   ├── __init__.py
│   ├── resources.py                  # Resource manager (fonts, themes)
│   │
│   ├── dialogs/                      # Dialog windows
│   │   ├── __init__.py
│   │   ├── terminal_dialog.py        # Main terminal UI (Pane, Container, Window)
│   │   └── connection_dialog.py      # Connection configuration dialogs
│   │
│   ├── windows/                      # Specialized window components
│   │   ├── __init__.py
│   │   └── terminal_formatter.py     # Terminal text formatting
│   │
│   ├── components/                   # Reusable UI components
│   │   ├── __init__.py
│   │   └── ribbon_toolbar.py         # Ribbon-style toolbar
│   │
│   └── common/                       # Shared UI utilities
│       ├── __init__.py
│       └── icons.py                  # SVG icon definitions
│
├── assets/                           # Static assets
│   ├── fonts/
│   │   ├── JetBrainsMono/           # Terminal monospace font
│   │   │   ├── JetBrainsMono-Regular.ttf
│   │   │   ├── JetBrainsMono-Bold.ttf
│   │   │   └── ...
│   │   │
│   │   └── Poppins/                  # UI font
│   │       ├── Poppins-Regular.ttf
│   │       ├── Poppins-Bold.ttf
│   │       └── ...
│   │
│   └── icons/                        # Application icons
│       ├── app_icon.ico              # Windows application icon
│       ├── app_icon.png              # Cross-platform icon
│       └── ...
│
├── tests/                            # Test files
│   ├── __init__.py
│   ├── test_network_config.py        # NetworkConfig tests
│   ├── test_network_worker.py        # NetworkWorker tests
│   ├── test_connection_manager.py    # ConnectionManager tests
│   └── test_integration.py           # Integration tests
│
└── docs/                             # Documentation (development only)
    ├── PRD_NetworkTerminal.md
    ├── phases/
    │   ├── phase-0-project-setup.md
    │   ├── stream-a-core-network.md
    │   ├── stream-b-ui-layer.md
    │   ├── stream-c-config-data.md
    │   ├── phase-4-integration.md
    │   └── phase-5-testing.md
    │
    └── reference/
        ├── architecture-analysis.md
        ├── reusability-matrix.md
        ├── interface-definitions.md
        ├── file-structure.md
        └── testing-requirements.md
```

---

## File Descriptions

### Root Files

| File | Purpose | Source |
|------|---------|--------|
| `main.py` | Application entry point | New (Phase 0) |
| `constants.py` | App name, colors, constants | Copied + modified |
| `build.py` | PyInstaller build configuration | Copied + modified |
| `requirements.txt` | Python dependencies | New |
| `README.md` | User documentation | New |
| `LICENSE` | License terms | New |

### core/ Directory

| File | Purpose | Source | Stream |
|------|---------|--------|--------|
| `__init__.py` | Package marker | New | Phase 0 |
| `core.py` | Utility classes | Extracted from Serial Terminal | Stream C |
| `network_config.py` | Configuration dataclass | New | Stream C |
| `network_worker.py` | Network communication | New | Stream A |
| `connection_manager.py` | Favorites/history | New | Stream C |

### ui/ Directory

| File | Purpose | Source | Stream |
|------|---------|--------|--------|
| `__init__.py` | Package marker | New | Phase 0 |
| `resources.py` | Resource manager | Copied unchanged | Phase 0 |

### ui/dialogs/ Directory

| File | Purpose | Source | Stream |
|------|---------|--------|--------|
| `__init__.py` | Package marker | New | Phase 0 |
| `terminal_dialog.py` | Main terminal components | Adapted | Stream B |
| `connection_dialog.py` | Connection dialogs | New | Stream B |

### ui/windows/ Directory

| File | Purpose | Source | Stream |
|------|---------|--------|--------|
| `__init__.py` | Package marker | New | Phase 0 |
| `terminal_formatter.py` | Text formatting | Copied unchanged | Phase 0 |

### ui/components/ Directory

| File | Purpose | Source | Stream |
|------|---------|--------|--------|
| `__init__.py` | Package marker | New | Phase 0 |
| `ribbon_toolbar.py` | Toolbar component | Copied unchanged | Phase 0 |

### ui/common/ Directory

| File | Purpose | Source | Stream |
|------|---------|--------|--------|
| `__init__.py` | Package marker | New | Phase 0 |
| `icons.py` | Icon definitions | Copied unchanged | Phase 0 |

### assets/ Directory

| Directory | Purpose | Source |
|-----------|---------|--------|
| `fonts/JetBrainsMono/` | Terminal font | Copied |
| `fonts/Poppins/` | UI font | Copied |
| `icons/` | App icons | Copied/Created |

### tests/ Directory

| File | Purpose | Owner |
|------|---------|-------|
| `__init__.py` | Package marker | Any |
| `test_network_config.py` | Config unit tests | Stream C |
| `test_network_worker.py` | Worker unit tests | Stream A |
| `test_connection_manager.py` | Manager tests | Stream C |
| `test_integration.py` | Integration tests | Phase 4 |

---

## Class Locations

### core/core.py Classes

```python
# Extracted from SerialTerminal/core/core.py
class WindowConfig:
    """Configuration for window sizing"""
    pass

class SettingsManager:
    """Application settings via QSettings"""
    pass

class ResponsiveWindowManager:
    """Responsive layout calculations"""
    pass
```

### core/network_config.py Classes

```python
# New file - Stream C
@dataclass
class NetworkConfig:
    """Network connection configuration"""
    pass

def get_default_configs() -> List[NetworkConfig]:
    """Get preset configurations"""
    pass
```

### core/network_worker.py Classes

```python
# New file - Stream A
class NetworkWorker(QThread):
    """Base class for network workers"""
    pass

class TCPClientWorker(NetworkWorker):
    """TCP client connection handler"""
    pass

class TCPServerWorker(NetworkWorker):
    """TCP server (listen mode) handler"""
    pass

class UDPWorker(NetworkWorker):
    """UDP sender/receiver handler"""
    pass

def create_network_worker(config: NetworkConfig) -> NetworkWorker:
    """Factory function"""
    pass
```

### core/connection_manager.py Classes

```python
# New file - Stream C
@dataclass
class ConnectionHistory:
    """Record of past connection"""
    pass

class ConnectionManager:
    """Favorites and history manager"""
    pass
```

### ui/dialogs/terminal_dialog.py Classes

```python
# Adapted from SerialTerminal
class NetworkTerminalPane(QWidget):
    """Terminal pane for network connections"""
    pass

class SplitContainer(QWidget):
    """Split pane container (copied unchanged)"""
    pass

class NetworkMonitorWindow(QMainWindow):
    """Main application window"""
    pass
```

### ui/dialogs/connection_dialog.py Classes

```python
# New file - Stream B
class QuickConnectDialog(QDialog):
    """Quick connect configuration dialog"""
    pass

class ConnectionHistoryDialog(QDialog):
    """Favorites and history dialog"""
    pass
```

---

## Import Structure

### main.py Imports

```python
from PyQt6.QtWidgets import QApplication
from constants import AppInfo
from ui.dialogs.terminal_dialog import NetworkMonitorWindow
```

### terminal_dialog.py Imports

```python
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *

from ui.windows.terminal_formatter import TerminalStreamFormatter
from ui.resources import resource_manager
from ui.components.ribbon_toolbar import RibbonToolbar
from ui.common.icons import Icons

from core.network_config import NetworkConfig
from core.network_worker import NetworkWorker, create_network_worker
from core.connection_manager import ConnectionManager
```

### network_worker.py Imports

```python
from PyQt6.QtCore import QThread, pyqtSignal
import socket
import queue
import threading
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from .network_config import NetworkConfig
```

---

## File Size Estimates

| File | Estimated Lines | Notes |
|------|-----------------|-------|
| `main.py` | ~80 | Entry point, dark theme |
| `constants.py` | ~50 | Colors, app info |
| `core/core.py` | ~200 | Utility classes only |
| `core/network_config.py` | ~300 | Dataclass + methods |
| `core/network_worker.py` | ~400 | All worker classes |
| `core/connection_manager.py` | ~200 | Manager class |
| `ui/dialogs/terminal_dialog.py` | ~1500 | Pane, Container, Window |
| `ui/dialogs/connection_dialog.py` | ~400 | Dialogs |
| `ui/windows/terminal_formatter.py` | ~350 | Copied unchanged |
| `ui/resources.py` | ~200 | Copied unchanged |
| `ui/components/ribbon_toolbar.py` | ~150 | Copied unchanged |
| `ui/common/icons.py` | ~100 | Copied unchanged |
| **Total** | **~3,930** | Estimated |

---

## Directory Creation Commands

```bash
# Create all directories
mkdir -p NetworkTerminal/{core,ui/{dialogs,windows,components,common},assets/{fonts,icons},tests,docs/{phases,reference}}

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

---

## File Dependencies Graph

```
main.py
├── constants.py
├── ui/resources.py
└── ui/dialogs/terminal_dialog.py
    ├── ui/windows/terminal_formatter.py
    │   ├── constants.py (TerminalColors)
    │   └── ui/resources.py (fonts)
    ├── ui/components/ribbon_toolbar.py
    ├── ui/common/icons.py
    ├── core/network_config.py
    ├── core/network_worker.py
    │   └── core/network_config.py (TYPE_CHECKING)
    └── core/connection_manager.py
        └── core/network_config.py
```

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Phase 0: Project Setup](../phases/phase-0-project-setup.md)
- [Reusability Matrix](./reusability-matrix.md)
- [Architecture Analysis](./architecture-analysis.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md)
