# Kickoff Prompt: Phase 5 - Testing & Polish

## Context

You are implementing Phase 5 (Testing & Polish) of the Network Terminal project. Phase 4 (Integration) is complete - all components are wired together and the application is functional. Your task is to test, optimize, fix bugs, and prepare for distribution.

**Repository**: `/home/user/SerialTerminal`
**Network Terminal Location**: `/home/user/SerialTerminal/NetworkTerminal/`

## Completed Work

| Phase | Description | Commit | Status |
|-------|-------------|--------|--------|
| Stream A | NetworkWorker TCP/UDP classes | `fef6a26` | Complete |
| Stream B | UI Layer (NetworkTerminalPane, dialogs) | `dc409b7` | Complete |
| Stream C | Config & Data Layer | `e2f380c` | Complete |
| Phase 4 | Integration (ConnectionManager + UI) | `1c0cd91` | Complete |

## Your Task

Test, polish, and prepare the application for distribution:

1. **Task 5.1**: Execute manual testing suite
2. **Task 5.2**: Performance testing
3. **Task 5.3**: Bug fixes and polish
4. **Task 5.4**: Build and packaging

## Branch

Create a new branch from current HEAD or work on existing branch.

## Documentation

**Detailed spec**: `/home/user/SerialTerminal/docs/phases/phase-5-testing.md` - READ THIS FIRST

## Overview

| Attribute | Value |
|-----------|-------|
| **Phase** | 5 - Testing & Polish |
| **Files** | Various (bug fixes), `build.py` (new) |
| **Tasks** | 4 |
| **New Lines** | ~50-100 (build script + fixes) |
| **Dependencies** | Phase 4 complete |

---

## Task 5.1: Manual Testing Suite

Execute comprehensive tests. For each test, report PASS/FAIL and any issues found.

### Quick Test Commands

```bash
# Terminal 1: Start TCP server for testing
nc -l 5000

# Terminal 2: Run Network Terminal
cd /home/user/SerialTerminal/NetworkTerminal
python main.py
```

### Critical Test Cases

| Test | Steps | Expected |
|------|-------|----------|
| TCP Client Connect | Create TCP Client to localhost:5000 | "Connected" status |
| TCP Send/Receive | Type in terminal, check nc | Bidirectional data flow |
| TCP Disconnect | Close nc (Ctrl+C) | "Connection closed" message |
| TCP Reconnect | Click Connect after disconnect | Reconnects successfully |
| TCP Server | Create TCP Server on 5000, connect with nc | Accepts client |
| UDP Send | Create UDP client, send to nc -ul 5000 | Data received |
| Split Panes | Alt+Shift+- and Alt+Shift++ | Panes split correctly |
| History | File > Connection History | Shows recent connections |
| Favorites | Right-click > Add to Favorites | Added and persists |

### Test Report Template

```markdown
## Test Results - [Date]

### TCP Client Tests
- [ ] TC-01 Basic connection: PASS/FAIL
- [ ] TC-02 Send data: PASS/FAIL
- [ ] TC-03 Receive data: PASS/FAIL
- [ ] TC-04 Connection refused: PASS/FAIL
- [ ] TC-06 Server disconnect: PASS/FAIL
- [ ] TC-07 Reconnect: PASS/FAIL

### TCP Server Tests
- [ ] TS-01 Start server: PASS/FAIL
- [ ] TS-02 Accept client: PASS/FAIL
- [ ] TS-03 Receive from client: PASS/FAIL
- [ ] TS-04 Send to client: PASS/FAIL

### UI Tests
- [ ] UI-01 Split vertical: PASS/FAIL
- [ ] UI-02 Split horizontal: PASS/FAIL
- [ ] UI-05 New tab (Ctrl+N): PASS/FAIL
- [ ] UI-06 Close tab (Ctrl+W): PASS/FAIL

### History/Favorites Tests
- [ ] HF-01 Add favorite: PASS/FAIL
- [ ] HF-02 Show favorites: PASS/FAIL
- [ ] HF-04 Persistence: PASS/FAIL

### Issues Found
1. [Issue description]
2. [Issue description]
```

---

## Task 5.2: Performance Testing

### High Data Rate Test

```bash
# Generate rapid data (adjust rate as needed)
for i in $(seq 1 1000); do
    echo "Line $i: $(date +%H:%M:%S) - Test data payload"
    sleep 0.01
done | nc -l 5000
```

**Check:**
- [ ] UI remains responsive
- [ ] No visible lag in data display
- [ ] Memory usage reasonable

### Multiple Connections Test

1. Open 3-4 tabs with different connections
2. Send data to all simultaneously
3. Verify all terminals update correctly

---

## Task 5.3: Bug Fixes & Polish

Fix any issues found during testing. Common areas to check:

### Known Areas to Verify

1. **Clean Disconnect**: Verify `cleanup()` runs without errors
2. **Memory**: Check for leaks on reconnect cycles
3. **Thread Safety**: Ensure signals use `QueuedConnection`
4. **Status Updates**: Verify status bar reflects actual state
5. **Dark Theme**: Check all dialogs use dark palette

### Code Quality Checklist

- [ ] No commented-out code
- [ ] Consistent error messages
- [ ] Proper resource cleanup
- [ ] Thread-safe signal connections

---

## Task 5.4: Build & Packaging

### Create Build Script

Create `NetworkTerminal/build.py`:

```python
#!/usr/bin/env python3
"""Build script for Network Terminal"""

import subprocess
import sys
import shutil
from pathlib import Path

APP_NAME = "NetworkTerminal"
ENTRY_POINT = "main.py"

def build():
    """Build the application"""
    print(f"Building {APP_NAME}...")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--onefile",
        "--windowed",
        "--clean",
        "--add-data", "assets;assets",
        ENTRY_POINT
    ]

    subprocess.run(cmd, check=True)
    print(f"Build complete! Executable at: dist/{APP_NAME}")

def clean():
    """Clean build artifacts"""
    for path in ["build", "dist", f"{APP_NAME}.spec"]:
        p = Path(path)
        if p.exists():
            shutil.rmtree(p) if p.is_dir() else p.unlink()
    print("Cleaned build artifacts")

if __name__ == "__main__":
    clean() if len(sys.argv) > 1 and sys.argv[1] == "clean" else build()
```

### Test Build

```bash
cd /home/user/SerialTerminal/NetworkTerminal
pip install pyinstaller  # if not installed
python build.py
./dist/NetworkTerminal  # Test the executable
```

---

## Key Files Reference

```bash
# Main entry point
NetworkTerminal/main.py

# Core components
NetworkTerminal/core/network_worker.py
NetworkTerminal/core/network_config.py
NetworkTerminal/core/connection_manager.py

# UI components
NetworkTerminal/ui/dialogs/terminal_dialog.py
NetworkTerminal/ui/dialogs/connection_dialog.py
NetworkTerminal/ui/components/ribbon_toolbar.py
```

---

## Acceptance Criteria

### Testing
- [ ] All critical TCP tests pass
- [ ] All critical UDP tests pass
- [ ] All UI tests pass
- [ ] History/Favorites tests pass

### Performance
- [ ] Application starts in < 3 seconds
- [ ] UI responsive during data transfer
- [ ] Memory stable during long sessions

### Polish
- [ ] All found bugs fixed
- [ ] Error messages user-friendly
- [ ] Dark theme consistent

### Build
- [ ] Build script created and works
- [ ] Executable runs standalone
- [ ] No missing dependencies

---

## Commit Message Template

```
Phase 5: Testing, polish, and build preparation

- Execute manual test suite (X/Y tests passing)
- Fix [list bugs fixed]
- Add build.py for PyInstaller packaging
- Polish [list polish items]
```

---

## After Completion

Phase 5 marks the completion of Network Terminal v1.0. Final deliverables:

1. All tests passing
2. No critical bugs
3. Working build script
4. Standalone executable

The application should be ready for daily use alongside Serial Terminal.
