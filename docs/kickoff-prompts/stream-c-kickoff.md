# Kickoff Prompt: Stream C - Config & Data Layer

## Context

You are implementing Stream C of the Network Terminal project. Phase 0 is complete - the project skeleton exists at `/home/user/SerialTerminal/NetworkTerminal/`.

## Your Task

Create the configuration and data management layer:
1. NetworkConfig dataclass (connection parameters)
2. ConnectionManager (favorites and history persistence)
3. Utility classes (SettingsManager, ResponsiveWindowManager)

## Branch

Work on branch: `dev/network-terminal` (or `claude/setup-network-terminal-SRbHj`)

## Documentation

**Detailed spec**: `/home/user/SerialTerminal/docs/phases/stream-c-config-data.md` - READ THIS FIRST

## Overview

| Attribute | Value |
|-----------|-------|
| **Stream** | C - Config & Data |
| **Files** | `NetworkTerminal/core/network_config.py`, `connection_manager.py`, `core.py` |
| **Tasks** | 3 |
| **New Lines** | ~250 |
| **Dependencies** | None (can start immediately) |

## Tasks

### Task C.1: Create NetworkConfig Dataclass

**File**: `NetworkTerminal/core/network_config.py`

Create a dataclass with:
- Required fields: `host`, `port`
- Protocol settings: `protocol` (TCP/UDP), `mode` (client/server)
- TCP options: `keepalive`, `nodelay`, `keepalive_interval`
- UDP options: `broadcast`, `multicast_group`, `multicast_ttl`
- Common options: `buffer_size`, `timeout`, `reconnect_on_disconnect`

Key methods:
- `__post_init__()` - Validate all parameters
- `get_display_string()` - Human-readable string for status bar
- `get_connection_id()` - Unique identifier for deduplication
- `to_dict()` / `from_dict()` - Serialization
- `to_json()` / `from_json()` - JSON serialization

### Task C.2: Create ConnectionManager

**File**: `NetworkTerminal/core/connection_manager.py`

Manages:
- Favorites (add, remove, check, update)
- History (record connections, limit to MAX_HISTORY)
- Search across favorites and history
- Persistence via QSettings

### Task C.3: Extract Utility Classes

**File**: `NetworkTerminal/core/core.py`

Copy from SerialTerminal with namespace changes:
- `SettingsManager` - Application settings (change namespace to "NetworkTerminal")
- `ResponsiveWindowManager` - Window sizing calculations (copy unchanged)
- `WindowConfig` - Window configuration dataclass

## Key Implementation Notes

1. **Validation in NetworkConfig.__post_init__():**
   - Protocol must be 'TCP' or 'UDP'
   - Mode must be 'client' or 'server'
   - Port must be 1-65535
   - Multicast addresses must be in 224.0.0.0 - 239.255.255.255 range

2. **ConnectionManager persistence:**
   - Use QSettings with organization "NetworkTerminal"
   - Store as JSON strings for complex objects
   - Handle deserialization errors gracefully

## Reference Files

Read from SerialTerminal for patterns:
```bash
cat /home/user/SerialTerminal/core/core.py  # SettingsManager, ResponsiveWindowManager
```

## Acceptance Criteria

- [ ] NetworkConfig validates all parameters on creation
- [ ] NetworkConfig serialization/deserialization works
- [ ] ConnectionManager persists via QSettings
- [ ] Favorites CRUD operations work
- [ ] History recording and limiting works
- [ ] SettingsManager uses "NetworkTerminal" namespace
- [ ] ResponsiveWindowManager copied unchanged

## Files to Read First

```bash
cat /home/user/SerialTerminal/docs/phases/stream-c-config-data.md
cat /home/user/SerialTerminal/NetworkTerminal/core/network_config.py  # Current placeholder
cat /home/user/SerialTerminal/NetworkTerminal/core/connection_manager.py  # Current placeholder
cat /home/user/SerialTerminal/NetworkTerminal/core/core.py  # Current placeholder
```

Commit your changes when complete.
