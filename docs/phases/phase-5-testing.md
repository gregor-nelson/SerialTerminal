# Phase 5: Testing & Polish

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Phase 4](./phase-4-integration.md)

---

## Overview

| Attribute | Value |
|-----------|-------|
| **Phase** | 5 - Testing & Polish |
| **Owner** | All Developers |
| **Dependencies** | Phase 4 complete |
| **Estimated Effort** | 4 tasks, ~50 new + ~50 adapted lines |
| **Status** | Not Started |

---

## Purpose

Complete the application with thorough testing, performance optimization, and packaging preparation:
1. Manual testing of all connection types
2. Performance testing under load
3. Bug fixes and polish
4. Build and packaging

---

## Task Checklist

- [ ] [Task 5.1: Manual Testing Suite](#task-51-manual-testing-suite)
- [ ] [Task 5.2: Performance Testing](#task-52-performance-testing)
- [ ] [Task 5.3: Bug Fixes & Polish](#task-53-bug-fixes--polish)
- [ ] [Task 5.4: Build & Packaging](#task-54-build--packaging)

---

## Task 5.1: Manual Testing Suite

### Objective

Execute comprehensive manual tests for all connection types and UI features.

### TCP Client Tests

| Test ID | Description | Steps | Expected Result | Pass |
|---------|-------------|-------|-----------------|------|
| TC-01 | Basic connection | 1. Start `nc -l 5000`<br>2. Create TCP Client to localhost:5000<br>3. Verify connection | "Connected" status shown | [ ] |
| TC-02 | Send data | 1. Connect as TC-01<br>2. Type text in terminal<br>3. Press Enter | Data appears in nc terminal | [ ] |
| TC-03 | Receive data | 1. Connect as TC-01<br>2. Type in nc terminal | Data appears in Network Terminal with timestamp | [ ] |
| TC-04 | Connection refused | 1. Create TCP Client to localhost:59999 (no server) | "Connection refused" error shown | [ ] |
| TC-05 | DNS failure | 1. Create TCP Client to invalid.host.test:5000 | "DNS resolution failed" error | [ ] |
| TC-06 | Server disconnect | 1. Connect as TC-01<br>2. Close nc (Ctrl+C) | "Connection closed by remote" message | [ ] |
| TC-07 | Reconnect | 1. After TC-06, click Connect again | Reconnects successfully | [ ] |
| TC-08 | Large data | 1. Connect as TC-01<br>2. Send 1MB file via nc | Data displayed without crash | [ ] |
| TC-09 | Keepalive | 1. Enable keepalive in dialog<br>2. Connect and wait 2 minutes | Connection stays alive | [ ] |

### TCP Server Tests

| Test ID | Description | Steps | Expected Result | Pass |
|---------|-------------|-------|-----------------|------|
| TS-01 | Start server | 1. Create TCP Server on port 5000 | "Listening" status shown | [ ] |
| TS-02 | Accept client | 1. Start server as TS-01<br>2. `nc localhost 5000` | "Client connected" message | [ ] |
| TS-03 | Receive from client | 1. Connect client as TS-02<br>2. Type in nc | Data appears in server terminal | [ ] |
| TS-04 | Send to client | 1. Connect client as TS-02<br>2. Type in server terminal | Data appears in nc | [ ] |
| TS-05 | Client disconnect | 1. Connect client as TS-02<br>2. Close nc | "Client disconnected" message | [ ] |
| TS-06 | Accept new client | 1. After TS-05, run `nc localhost 5000` again | New client connected | [ ] |
| TS-07 | Port in use | 1. Start server on 5000<br>2. Start another on 5000 | "Port already in use" error | [ ] |
| TS-08 | Low port | 1. Start server on port 80 (non-root) | "Permission denied" error | [ ] |

### UDP Tests

| Test ID | Description | Steps | Expected Result | Pass |
|---------|-------------|-------|-----------------|------|
| TU-01 | Send UDP | 1. Create UDP client to localhost:5000<br>2. Start `nc -ul 5000`<br>3. Type in terminal | Data appears in nc | [ ] |
| TU-02 | Receive UDP | 1. Create UDP server on 5000<br>2. `echo "test" \| nc -u localhost 5000` | Data appears in terminal | [ ] |
| TU-03 | Broadcast | 1. Enable broadcast<br>2. Send to 255.255.255.255 | Data sent without error | [ ] |

### UI Tests

| Test ID | Description | Steps | Expected Result | Pass |
|---------|-------------|-------|-----------------|------|
| UI-01 | Split vertical | 1. Press Alt+Shift+- | Pane splits vertically | [ ] |
| UI-02 | Split horizontal | 1. Press Alt+Shift++ | Pane splits horizontally | [ ] |
| UI-03 | Close pane | 1. Press Ctrl+Shift+W | Pane closes | [ ] |
| UI-04 | Navigate panes | 1. Alt+Arrow keys | Focus moves between panes | [ ] |
| UI-05 | New tab | 1. Ctrl+N | New tab created | [ ] |
| UI-06 | Close tab | 1. Ctrl+W or click X | Tab closes | [ ] |
| UI-07 | Clear terminal | 1. Right-click > Clear | Terminal cleared | [ ] |
| UI-08 | Hex mode | 1. Right-click > Hex Display | Data shows as hex | [ ] |
| UI-09 | Auto-scroll | 1. Toggle auto-scroll<br>2. Receive data | Scroll behavior matches setting | [ ] |
| UI-10 | Font size | 1. Ctrl++ and Ctrl+- | Font size changes | [ ] |

### History/Favorites Tests

| Test ID | Description | Steps | Expected Result | Pass |
|---------|-------------|-------|-----------------|------|
| HF-01 | Add favorite | 1. Connect<br>2. Right-click > Add to Favorites | "Added to favorites" message | [ ] |
| HF-02 | Show favorites | 1. File > Connection History<br>2. Check Favorites tab | Favorite appears in list | [ ] |
| HF-03 | Connect from history | 1. Double-click history item | Connection established | [ ] |
| HF-04 | Persistence | 1. Add favorite<br>2. Close app<br>3. Reopen | Favorite still present | [ ] |

### Acceptance Criteria

- [ ] All TCP Client tests pass
- [ ] All TCP Server tests pass
- [ ] All UDP tests pass
- [ ] All UI tests pass
- [ ] All History/Favorites tests pass

---

## Task 5.2: Performance Testing

### Objective

Verify application handles high data rates and long-running sessions.

### Test Scenarios

#### High Data Rate Test

```bash
# Generate 1000 lines/second for 60 seconds
seq 1 60000 | while read i; do
    echo "Line $i: $(date +%H:%M:%S.%N) - Lorem ipsum dolor sit amet"
    sleep 0.001
done | nc -l 5000
```

**Metrics to Measure:**
- [ ] CPU usage stays below 50%
- [ ] Memory growth < 50MB over 60 seconds
- [ ] No dropped data
- [ ] UI remains responsive

#### Long-Running Session Test

```bash
# Run for 1 hour with periodic data
while true; do
    echo "Heartbeat: $(date)"
    sleep 10
done | nc -l 5000
```

**Metrics to Measure:**
- [ ] Memory stable after 1 hour
- [ ] No UI freezes
- [ ] Connection remains stable
- [ ] Disconnect/reconnect still works

#### Multiple Connection Test

1. Open 4 tabs with different connections
2. Send data to all simultaneously
3. Verify all terminals update correctly

**Metrics to Measure:**
- [ ] All connections handle data simultaneously
- [ ] No cross-talk between panes
- [ ] Memory usage reasonable (< 200MB for 4 connections)

### Performance Baseline

| Metric | Target | Actual |
|--------|--------|--------|
| Startup time | < 2 seconds | |
| Connection time (localhost) | < 100ms | |
| Data display latency | < 50ms | |
| Memory at idle | < 80MB | |
| Memory with 10k lines | < 150MB | |
| CPU at idle | < 1% | |
| CPU at 1000 lines/sec | < 30% | |

### Memory Profiling (Optional)

```python
# Add to main.py for memory profiling
import tracemalloc
tracemalloc.start()

# ... application code ...

# At shutdown:
snapshot = tracemalloc.take_snapshot()
top_stats = snapshot.statistics('lineno')
print("[ Top 10 memory allocations ]")
for stat in top_stats[:10]:
    print(stat)
```

### Acceptance Criteria

- [ ] High data rate test passes
- [ ] Long-running session stable
- [ ] Multiple connections work
- [ ] Performance baselines met

---

## Task 5.3: Bug Fixes & Polish

### Objective

Address issues found during testing and polish the user experience.

### Common Issues Checklist

| Issue | Category | Fix Applied | Verified |
|-------|----------|-------------|----------|
| Connection doesn't close cleanly | Threading | | [ ] |
| Memory leak on reconnect | Resource | | [ ] |
| Terminal scroll jumps | UI | | [ ] |
| Status bar not updating | Signal | | [ ] |
| Dark theme inconsistent | Styling | | [ ] |
| Keyboard shortcuts not working | Event | | [ ] |

### Polish Items

#### Error Messages

- [ ] All errors have user-friendly messages
- [ ] Technical details available in tooltip or expandable section
- [ ] Errors don't spam the terminal

#### Status Feedback

- [ ] Clear indication of connection state
- [ ] Progress shown during connection attempt
- [ ] Byte counters update in real-time

#### Visual Polish

- [ ] Consistent spacing and margins
- [ ] Icons match Serial Terminal style
- [ ] Fonts render correctly at all sizes
- [ ] Color scheme works on various monitors

#### Documentation

- [ ] Help text accurate and complete
- [ ] Keyboard shortcuts documented
- [ ] Version info displayed correctly

### Code Quality Checklist

- [ ] No commented-out code
- [ ] Consistent naming conventions
- [ ] No hardcoded values (use constants)
- [ ] Proper error handling everywhere
- [ ] Thread-safe where needed
- [ ] Resources properly cleaned up

### Acceptance Criteria

- [ ] All found bugs fixed
- [ ] Polish items complete
- [ ] Code quality checklist passes

---

## Task 5.4: Build & Packaging

### Objective

Prepare the application for distribution.

### Build Script Update

Update `build.py` for Network Terminal:

```python
#!/usr/bin/env python3
"""
Build script for Network Terminal
Creates standalone executable using PyInstaller
"""

import subprocess
import sys
import shutil
from pathlib import Path

# Application info
APP_NAME = "NetworkTerminal"
ENTRY_POINT = "main.py"
ICON_PATH = "assets/icons/app_icon.ico"  # Create if needed

def build():
    """Build the application"""
    print(f"Building {APP_NAME}...")

    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--onefile",
        "--windowed",  # No console window
        "--clean",
        # Add data files
        "--add-data", "assets;assets",
        "--add-data", "constants.py;.",
    ]

    # Add icon if exists
    if Path(ICON_PATH).exists():
        cmd.extend(["--icon", ICON_PATH])

    # Add entry point
    cmd.append(ENTRY_POINT)

    # Run PyInstaller
    result = subprocess.run(cmd, check=True)

    if result.returncode == 0:
        print(f"Build successful! Executable at: dist/{APP_NAME}")
    else:
        print("Build failed!")
        sys.exit(1)

def clean():
    """Clean build artifacts"""
    for path in ["build", "dist", f"{APP_NAME}.spec"]:
        if Path(path).exists():
            if Path(path).is_dir():
                shutil.rmtree(path)
            else:
                Path(path).unlink()
    print("Cleaned build artifacts")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "clean":
        clean()
    else:
        build()
```

### Build Commands

```bash
# Install build dependencies
pip install pyinstaller

# Build executable
python build.py

# Clean build artifacts
python build.py clean
```

### Package Contents

```
NetworkTerminal/
├── NetworkTerminal.exe          # Main executable
├── README.md                    # User documentation
├── LICENSE                      # License file
└── assets/                      # (embedded in exe)
```

### Distribution Checklist

- [ ] Executable runs on clean Windows system
- [ ] All fonts render correctly
- [ ] Icons display properly
- [ ] No missing DLL errors
- [ ] Settings persist correctly
- [ ] Application can be uninstalled cleanly

### Platform Testing

| Platform | Version | Tested | Notes |
|----------|---------|--------|-------|
| Windows 10 | 21H2+ | [ ] | Primary target |
| Windows 11 | 22H2+ | [ ] | |
| macOS | 12+ | [ ] | Optional |
| Linux | Ubuntu 22.04 | [ ] | Optional |

### Acceptance Criteria

- [ ] Build script works
- [ ] Executable runs standalone
- [ ] Distribution checklist complete
- [ ] Tested on target platforms

---

## Phase Deliverables

After completing Phase 5, the following should be ready:

| Item | Status |
|------|--------|
| Manual test suite executed | |
| All critical tests passing | |
| Performance baselines met | |
| Bug fixes applied | |
| Polish items complete | |
| Build script working | |
| Executable created | |
| Distribution tested | |

---

## Release Checklist

Before declaring version 1.0 complete:

### Functionality

- [ ] TCP Client connections work
- [ ] TCP Server connections work
- [ ] UDP connections work
- [ ] Split pane UI works
- [ ] History/Favorites work
- [ ] All keyboard shortcuts work

### Quality

- [ ] No critical bugs
- [ ] Performance acceptable
- [ ] Memory usage stable
- [ ] Clean shutdown

### Documentation

- [ ] README up to date
- [ ] Help text accurate
- [ ] Version number set

### Distribution

- [ ] Executable builds successfully
- [ ] Runs on clean system
- [ ] Settings persist

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Phase 4: Integration](./phase-4-integration.md)
- [Testing Requirements](../reference/testing-requirements.md)
- [Architecture Analysis](../reference/architecture-analysis.md)

---

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Phase 4](./phase-4-integration.md)
