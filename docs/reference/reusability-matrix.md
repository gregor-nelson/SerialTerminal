# Reusability Matrix

[← Back to Master PRD](../PRD_NetworkTerminal.md)

---

## Overview

This document provides detailed copy/adapt/replace decisions for each file and class in the Serial Terminal codebase. Use this as a reference when implementing Network Terminal.

---

## Quick Reference

| Action | Files | Lines |
|--------|-------|-------|
| **Copy Unchanged** | 5 | ~850 |
| **Copy with Minor Edits** | 2 | ~250 |
| **Adapt Significantly** | 3 | ~600 |
| **Replace Entirely** | 3 | ~400 |

---

## File-by-File Analysis

### Copy Unchanged (100% Reusable)

#### `ui/windows/terminal_formatter.py` (344 lines)

| Item | Action | Notes |
|------|--------|-------|
| `TerminalStreamFormatter` class | Copy | No modifications needed |
| `colors` dict | Copy | Same color scheme |
| `nmea_colors` dict | Copy | NMEA formatting useful for network |
| `_detect_nmea_message_type()` | Copy | Detects NMEA over network too |
| `append_data()` | Copy | Generic data display |
| `append_separator()` | Copy | Generic formatting |
| `append_status()` | Copy | Generic status display |
| Auto-scroll methods | Copy | Generic UI functionality |

**Copy Command:**
```bash
cp SerialTerminal/ui/windows/terminal_formatter.py \
   NetworkTerminal/ui/windows/terminal_formatter.py
```

---

#### `ui/components/ribbon_toolbar.py` (~150 lines)

| Item | Action | Notes |
|------|--------|-------|
| `RibbonToolbar` class | Copy | Generic toolbar component |
| All signals | Copy | Same toolbar actions |
| Button styling | Copy | Consistent look |

**Copy Command:**
```bash
cp SerialTerminal/ui/components/ribbon_toolbar.py \
   NetworkTerminal/ui/components/ribbon_toolbar.py
```

---

#### `ui/common/icons.py` (~100 lines)

| Item | Action | Notes |
|------|--------|-------|
| `Icons` class | Copy | SVG icon definitions |
| All icon methods | Copy | Generic icons |

**Copy Command:**
```bash
cp SerialTerminal/ui/common/icons.py \
   NetworkTerminal/ui/common/icons.py
```

---

#### `ui/resources.py` (~200 lines)

| Item | Action | Notes |
|------|--------|-------|
| `ResourceManager` class | Copy | Font and resource loading |
| `get_monospace_font()` | Copy | Used by formatter |
| Theme functions | Copy | Dark theme support |

**Copy Command:**
```bash
cp SerialTerminal/ui/resources.py \
   NetworkTerminal/ui/resources.py
```

---

#### `assets/` directory

| Item | Action | Notes |
|------|--------|-------|
| `fonts/JetBrainsMono/` | Copy | Terminal font |
| `fonts/Poppins/` | Copy | UI font |
| `icons/` | Copy | Application icons |

**Copy Command:**
```bash
cp -r SerialTerminal/assets/* NetworkTerminal/assets/
```

---

### Copy with Minor Edits

#### `constants.py` (~50 lines)

| Item | Action | Change |
|------|--------|--------|
| `TerminalColors` class | Copy | No changes |
| `AppInfo.NAME` | Edit | "Network Terminal" |
| `AppInfo.ORG_NAME` | Edit | "Network Terminal" |
| `AppInfo.ORG_DOMAIN` | Edit | "networkterminal.local" |

**Edits Required:**
```python
# Before (Serial Terminal)
class AppInfo:
    NAME = "Serial Monitor"
    ORG_NAME = "SerialSplit"
    ORG_DOMAIN = "serialsplit.local"

# After (Network Terminal)
class AppInfo:
    NAME = "Network Terminal"
    ORG_NAME = "Network Terminal"
    ORG_DOMAIN = "networkterminal.local"
```

---

#### `build.py` (~50 lines)

| Item | Action | Change |
|------|--------|--------|
| Build script | Edit | Change output name |
| PyInstaller config | Edit | Update app name |

**Edits Required:**
```python
# Change APP_NAME
APP_NAME = "NetworkTerminal"  # Was "SerialMonitor"
```

---

### Adapt Significantly

#### `ui/dialogs/terminal_dialog.py` - TerminalPane

| Method | Action | Adaptation Required |
|--------|--------|---------------------|
| `__init__()` | Adapt | Change `SerialConfig` → `NetworkConfig`, remove baud rate attrs |
| `_setup_ui()` | Copy | No changes |
| `_setup_context_menu()` | Copy | No changes |
| `_create_terminal_menu()` | Adapt | Remove baud/port menus, add network menus |
| `_create_font_size_menu()` | Copy | No changes |
| `_create_baud_rate_menu()` | Remove | Serial-specific |
| `_create_com_port_menu()` | Remove | Serial-specific |
| `checkbox_icon()` | Copy | No changes |
| `eventFilter()` | Copy | No changes |
| `_handle_key_press()` | Copy | No changes |
| `_send_raw_data()` | Adapt | Use `network_worker` |
| `_echo_local_data()` | Copy | No changes |
| `connect()` | Adapt | Use `create_network_worker()` |
| `disconnect()` | Adapt | Use `network_worker` |
| `cleanup()` | Adapt | Use `network_worker` |
| `_on_data_received()` | Adapt | Remove baud rate detection |
| `_on_error()` | Copy | No changes |
| `_on_connection_state_changed()` | Copy | No changes |
| `send_data()` | Adapt | Use `network_worker` |
| `get_status_info()` | Adapt | Network format |
| `_format_bytes()` | Copy | No changes |
| `_toggle_auto_scroll()` | Copy | No changes |
| `_toggle_hex_mode()` | Copy | No changes |
| `_toggle_local_echo()` | Copy | No changes |
| `_set_font_size()` | Copy | No changes |
| `_increase_font_size()` | Copy | No changes |
| `_decrease_font_size()` | Copy | No changes |
| `_reset_font_size()` | Copy | No changes |
| `_clear_terminal()` | Copy | No changes |
| `_scroll_to_bottom()` | Copy | No changes |
| `_show_help()` | Adapt | Update help text |
| `_dismiss_help()` | Copy | No changes |
| `_flush_buffer()` | Copy | No changes |
| `_set_baud_rate()` | Remove | Serial-specific |
| `_set_com_port()` | Remove | Serial-specific |
| `_handle_encoding_error()` | Remove | Serial-specific |
| `_show_baud_rate_suggestion()` | Remove | Serial-specific |
| `reset_baud_rate_detection()` | Remove | Serial-specific |
| `_handle_excessive_errors()` | Remove | Serial-specific |
| `_is_data_garbled()` | Remove | Serial-specific |
| `_show_virtual_port_manager()` | Remove | Serial-specific |

**Summary:** ~30 methods to copy, ~10 to remove, ~10 to adapt.

---

#### `ui/dialogs/terminal_dialog.py` - SplitContainer

| Method | Action | Notes |
|--------|--------|-------|
| Entire class | Copy | 100% reusable |

**Copy as-is** - handles split pane layout generically.

---

#### `ui/dialogs/terminal_dialog.py` - SerialMonitorWindow

| Method | Action | Adaptation Required |
|--------|--------|---------------------|
| `__init__()` | Adapt | Change title, add ConnectionManager |
| `_setup_ui()` | Adapt | Minor changes |
| `_setup_shortcuts()` | Copy | No changes |
| `_setup_status_bar()` | Copy | No changes |
| `_on_new_connection()` | Adapt | Use QuickConnectDialog |
| `_create_new_tab()` | Adapt | Use NetworkTerminalPane |
| `_close_tab()` | Copy | No changes |
| `_on_tab_changed()` | Copy | No changes |
| `_get_active_container()` | Copy | No changes |
| `_get_active_pane()` | Copy | No changes |
| `_update_status_bar()` | Copy | No changes |
| `_update_ribbon_connection_state()` | Copy | No changes |
| `_on_clear()` | Copy | No changes |
| `_on_connect_toggle()` | Copy | No changes |
| `_on_settings()` | Adapt | Network settings |
| `_navigate_panes()` | Copy | No changes |
| `get_connected_ports()` | Remove | Serial-specific |
| `closeEvent()` | Copy | No changes |

---

#### `core/core.py` - Utility Classes

| Class | Action | Notes |
|-------|--------|-------|
| `WindowConfig` | Copy | No changes |
| `SettingsManager` | Adapt | Change namespace |
| `ResponsiveWindowManager` | Copy | No changes |
| `PortScanner` | Remove | Serial-specific |
| `SerialPortMonitor` | Remove | Serial-specific |
| `SerialPortTester` | Remove | Serial-specific |
| `PortCapabilityAnalyzer` | Remove | Serial-specific |
| `Hub4comProcess` | Remove | Serial-specific |

---

### Replace Entirely

#### `core/serial_config.py` → `core/network_config.py`

| Original | Replacement | Notes |
|----------|-------------|-------|
| `SerialConfig` | `NetworkConfig` | Different fields |
| `port` (str) | `host` (str) | Target address |
| `baudrate` (int) | `port` (int) | Port number |
| `databits` | - | Not applicable |
| `parity` | `protocol` | TCP/UDP |
| `stopbits` | `mode` | client/server |

---

#### `core/core.py` - SerialWorker → NetworkWorker

| Original Signal | Keep | Notes |
|-----------------|------|-------|
| `dataReceived` | Yes | Same interface |
| `errorOccurred` | Yes | Same interface |
| `connectionStateChanged` | Yes | Same interface |

| New Signal | Purpose |
|------------|---------|
| `clientConnected` | Server mode only |
| `clientDisconnected` | Server mode only |

---

## Summary by Stream

### Stream A Developer (Core Network)

**New Files to Create:**
- `core/network_worker.py` (replace SerialWorker)
- Signals must match!

### Stream B Developer (UI Layer)

**Files to Copy:**
- `terminal_formatter.py` (unchanged)
- `ribbon_toolbar.py` (unchanged)
- `icons.py` (unchanged)
- `resources.py` (unchanged)

**Files to Adapt:**
- `TerminalPane` → `NetworkTerminalPane`
- `SerialMonitorWindow` → `NetworkMonitorWindow`
- `SplitContainer` (copy unchanged)

**New Files to Create:**
- `connection_dialog.py` (new)

### Stream C Developer (Config/Data)

**New Files to Create:**
- `core/network_config.py` (replace SerialConfig)
- `core/connection_manager.py` (new)

**Files to Adapt:**
- `core/core.py` (extract utilities only)

---

## Reuse Checklist

### Before Starting

- [ ] Read [Architecture Analysis](./architecture-analysis.md)
- [ ] Understand signal interfaces in [Interface Definitions](./interface-definitions.md)
- [ ] Review [File Structure](./file-structure.md)

### During Implementation

- [ ] Copy unchanged files first
- [ ] Adapt files incrementally
- [ ] Test signal compatibility
- [ ] Verify no serial imports remain

### After Implementation

- [ ] All tests pass
- [ ] No serial-specific code
- [ ] Signal interface matches
- [ ] UI looks identical

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Architecture Analysis](./architecture-analysis.md)
- [Interface Definitions](./interface-definitions.md)
- [File Structure](./file-structure.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md)
