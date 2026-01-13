# Stream C: Configuration & Data Layer

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Phase 0](./phase-0-project-setup.md)

---

## Overview

| Attribute | Value |
|-----------|-------|
| **Stream** | C - Config & Data |
| **Owner** | Developer 3 |
| **Dependencies** | Phase 0 complete |
| **Parallel With** | Stream A, Stream B |
| **Estimated Effort** | 3 tasks, ~250 new lines |
| **Status** | Not Started |

---

## Purpose

Create the configuration and data management layer:
1. `NetworkConfig` dataclass - connection parameters
2. `ConnectionManager` - favorites and history persistence
3. Extract utility classes from Serial Terminal

---

## Task Checklist

- [ ] [Task C.1: Create NetworkConfig](#task-c1-create-networkconfig)
- [ ] [Task C.2: Create ConnectionManager](#task-c2-create-connectionmanager)
- [ ] [Task C.3: Extract Utility Classes](#task-c3-extract-utility-classes)

---

## Task C.1: Create NetworkConfig

### Objective

Create the `NetworkConfig` dataclass that holds all connection parameters.

### File

`core/network_config.py`

### Requirements

The `NetworkConfig` must:
1. Store all connection parameters (host, port, protocol, mode)
2. Provide a `get_display_string()` method (like SerialConfig)
3. Support serialization for persistence
4. Validate parameters on creation

### Code

```python
#!/usr/bin/env python3
"""
Network Configuration Data Model

Replaces SerialConfig for network connections.
"""

from dataclasses import dataclass, field, asdict
from typing import Optional, Dict, Any
import json


@dataclass
class NetworkConfig:
    """
    Network connection configuration.

    This class mirrors SerialConfig's interface while providing
    network-specific parameters.

    Attributes:
        host: Target hostname or IP address
        port: Target port number (1-65535)
        protocol: 'TCP' or 'UDP'
        mode: 'client' or 'server'

    Example:
        # TCP Client
        config = NetworkConfig(
            host="192.168.1.100",
            port=5000,
            protocol="TCP",
            mode="client"
        )

        # TCP Server
        config = NetworkConfig(
            host="0.0.0.0",
            port=5000,
            protocol="TCP",
            mode="server"
        )

        # UDP
        config = NetworkConfig(
            host="192.168.1.255",
            port=5000,
            protocol="UDP",
            broadcast=True
        )
    """

    # === Required Fields ===
    host: str
    port: int

    # === Protocol Settings ===
    protocol: str = 'TCP'      # 'TCP' or 'UDP'
    mode: str = 'client'       # 'client' or 'server'

    # === TCP-Specific Options ===
    keepalive: bool = True
    keepalive_interval: int = 60      # seconds
    nodelay: bool = True              # TCP_NODELAY (disable Nagle algorithm)

    # === UDP-Specific Options ===
    broadcast: bool = False
    multicast_group: Optional[str] = None
    multicast_ttl: int = 1

    # === Common Options ===
    buffer_size: int = 4096
    timeout: float = 5.0              # Connection timeout in seconds
    reconnect_on_disconnect: bool = False
    reconnect_delay: float = 3.0      # seconds between reconnect attempts

    # === Display/Storage ===
    name: Optional[str] = None        # User-friendly name for favorites

    def __post_init__(self):
        """Validate configuration after initialization."""
        # Validate protocol
        if self.protocol not in ('TCP', 'UDP'):
            raise ValueError(
                f"Invalid protocol: {self.protocol}. Must be 'TCP' or 'UDP'"
            )

        # Validate mode
        if self.mode not in ('client', 'server'):
            raise ValueError(
                f"Invalid mode: {self.mode}. Must be 'client' or 'server'"
            )

        # Validate port
        if not isinstance(self.port, int) or not 1 <= self.port <= 65535:
            raise ValueError(
                f"Invalid port: {self.port}. Must be integer 1-65535"
            )

        # Validate host
        if not self.host or not isinstance(self.host, str):
            raise ValueError("Host cannot be empty")

        # Validate multicast group if specified
        if self.multicast_group:
            parts = self.multicast_group.split('.')
            if len(parts) != 4:
                raise ValueError(
                    f"Invalid multicast group: {self.multicast_group}"
                )
            try:
                first_octet = int(parts[0])
                if not 224 <= first_octet <= 239:
                    raise ValueError(
                        f"Multicast address must be in range 224.0.0.0 - "
                        f"239.255.255.255"
                    )
            except ValueError:
                raise ValueError(
                    f"Invalid multicast group: {self.multicast_group}"
                )

    def get_display_string(self) -> str:
        """
        Get display string for status bar.

        Mirrors SerialConfig.get_display_string() interface.

        Returns:
            Human-readable connection string

        Examples:
            "TCP | 192.168.1.100:5000"
            "TCP Server | 0.0.0.0:5000"
            "UDP Broadcast | 255.255.255.255:5000"
        """
        parts = [self.protocol]

        if self.mode == 'server':
            parts.append("Server")
        elif self.protocol == 'UDP' and self.broadcast:
            parts.append("Broadcast")

        parts.append(f"{self.host}:{self.port}")

        return " | ".join(parts)

    def get_connection_id(self) -> str:
        """
        Get unique identifier for this connection configuration.

        Used for deduplication in favorites/history.

        Returns:
            Unique string identifier
        """
        return f"{self.protocol}_{self.mode}_{self.host}_{self.port}"

    def get_short_name(self) -> str:
        """
        Get a short display name for tabs/menus.

        Returns:
            Short name like "TCP 5000" or custom name if set
        """
        if self.name:
            return self.name

        if self.mode == 'server':
            return f"{self.protocol} :{self.port}"
        else:
            # Shorten localhost
            host = self.host
            if host in ('127.0.0.1', 'localhost'):
                host = 'local'
            return f"{host}:{self.port}"

    def to_dict(self) -> Dict[str, Any]:
        """
        Convert to dictionary for JSON serialization.

        Returns:
            Dictionary representation
        """
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'NetworkConfig':
        """
        Create NetworkConfig from dictionary.

        Args:
            data: Dictionary with config values

        Returns:
            NetworkConfig instance

        Raises:
            ValueError: If required fields missing or invalid
        """
        # Filter to only known fields
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered_data = {k: v for k, v in data.items() if k in known_fields}

        return cls(**filtered_data)

    def to_json(self) -> str:
        """
        Serialize to JSON string.

        Returns:
            JSON string representation
        """
        return json.dumps(self.to_dict())

    @classmethod
    def from_json(cls, json_str: str) -> 'NetworkConfig':
        """
        Deserialize from JSON string.

        Args:
            json_str: JSON string

        Returns:
            NetworkConfig instance
        """
        data = json.loads(json_str)
        return cls.from_dict(data)

    def copy(self, **changes) -> 'NetworkConfig':
        """
        Create a copy with optional field changes.

        Args:
            **changes: Fields to override

        Returns:
            New NetworkConfig instance
        """
        data = self.to_dict()
        data.update(changes)
        return NetworkConfig.from_dict(data)


# === Preset Configurations ===

def get_default_configs() -> list:
    """
    Get list of common default configurations.

    Returns:
        List of NetworkConfig presets
    """
    return [
        NetworkConfig(
            host="127.0.0.1",
            port=5000,
            protocol='TCP',
            mode='client',
            name="Localhost TCP"
        ),
        NetworkConfig(
            host="0.0.0.0",
            port=5000,
            protocol='TCP',
            mode='server',
            name="TCP Server :5000"
        ),
        NetworkConfig(
            host="127.0.0.1",
            port=5000,
            protocol='UDP',
            mode='client',
            name="Localhost UDP"
        ),
    ]


# Export public API
__all__ = ['NetworkConfig', 'get_default_configs']
```

### Test Cases

```python
# test_network_config.py

import pytest
from core.network_config import NetworkConfig

def test_basic_creation():
    config = NetworkConfig(host="127.0.0.1", port=5000)
    assert config.host == "127.0.0.1"
    assert config.port == 5000
    assert config.protocol == "TCP"
    assert config.mode == "client"

def test_display_string():
    config = NetworkConfig(host="192.168.1.100", port=5000)
    assert config.get_display_string() == "TCP | 192.168.1.100:5000"

    config = NetworkConfig(host="0.0.0.0", port=5000, mode="server")
    assert "Server" in config.get_display_string()

def test_invalid_protocol():
    with pytest.raises(ValueError):
        NetworkConfig(host="127.0.0.1", port=5000, protocol="HTTP")

def test_invalid_port():
    with pytest.raises(ValueError):
        NetworkConfig(host="127.0.0.1", port=0)
    with pytest.raises(ValueError):
        NetworkConfig(host="127.0.0.1", port=70000)

def test_serialization():
    config = NetworkConfig(host="127.0.0.1", port=5000)
    json_str = config.to_json()
    restored = NetworkConfig.from_json(json_str)
    assert config.host == restored.host
    assert config.port == restored.port

def test_copy():
    config = NetworkConfig(host="127.0.0.1", port=5000)
    copy = config.copy(port=6000)
    assert copy.port == 6000
    assert config.port == 5000  # Original unchanged
```

### Acceptance Criteria

- [ ] NetworkConfig dataclass created with all fields
- [ ] `__post_init__` validates all parameters
- [ ] `get_display_string()` returns formatted string
- [ ] `to_dict()` / `from_dict()` work correctly
- [ ] `to_json()` / `from_json()` work correctly
- [ ] Invalid protocol raises ValueError
- [ ] Invalid port raises ValueError
- [ ] Invalid mode raises ValueError
- [ ] Unit tests pass

---

## Task C.2: Create ConnectionManager

### Objective

Create a manager for storing favorites and connection history.

### File

`core/connection_manager.py`

### Code

```python
#!/usr/bin/env python3
"""
Connection Manager - Favorites and History

Manages persistent storage of connection configurations.
"""

from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime
import json

from PyQt6.QtCore import QSettings

from .network_config import NetworkConfig


@dataclass
class ConnectionHistory:
    """
    Record of a past connection.

    Attributes:
        config: The connection configuration
        last_connected: When last used
        connect_count: Number of times connected
    """
    config: NetworkConfig
    last_connected: datetime
    connect_count: int = 1

    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            'config': self.config.to_dict(),
            'last_connected': self.last_connected.isoformat(),
            'connect_count': self.connect_count
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'ConnectionHistory':
        """Create from dictionary."""
        return cls(
            config=NetworkConfig.from_dict(data['config']),
            last_connected=datetime.fromisoformat(data['last_connected']),
            connect_count=data.get('connect_count', 1)
        )


class ConnectionManager:
    """
    Manages connection favorites and history.

    Uses QSettings for persistent storage across sessions.

    Usage:
        manager = ConnectionManager()

        # Add favorite
        config = NetworkConfig(host="192.168.1.100", port=5000)
        manager.add_favorite(config)

        # Record connection
        manager.record_connection(config)

        # Get favorites/history
        favorites = manager.get_favorites()
        history = manager.get_history()
    """

    MAX_HISTORY = 20

    def __init__(self):
        """Initialize the connection manager."""
        self.settings = QSettings("NetworkTerminal", "ConnectionManager")
        self._favorites: List[NetworkConfig] = []
        self._history: List[ConnectionHistory] = []
        self._load()

    def _load(self):
        """Load favorites and history from persistent storage."""
        # Load favorites
        favorites_json = self.settings.value("favorites", "[]")
        try:
            if isinstance(favorites_json, str):
                favorites_data = json.loads(favorites_json)
                self._favorites = [
                    NetworkConfig.from_dict(f) for f in favorites_data
                ]
        except (json.JSONDecodeError, TypeError, ValueError) as e:
            print(f"Warning: Failed to load favorites: {e}")
            self._favorites = []

        # Load history
        history_json = self.settings.value("history", "[]")
        try:
            if isinstance(history_json, str):
                history_data = json.loads(history_json)
                self._history = [
                    ConnectionHistory.from_dict(h) for h in history_data
                ]
        except (json.JSONDecodeError, TypeError, ValueError) as e:
            print(f"Warning: Failed to load history: {e}")
            self._history = []

    def _save(self):
        """Save favorites and history to persistent storage."""
        # Save favorites
        favorites_json = json.dumps([f.to_dict() for f in self._favorites])
        self.settings.setValue("favorites", favorites_json)

        # Save history
        history_json = json.dumps([h.to_dict() for h in self._history])
        self.settings.setValue("history", history_json)

        self.settings.sync()

    # === Favorites Management ===

    def get_favorites(self) -> List[NetworkConfig]:
        """
        Get all favorite connections.

        Returns:
            List of NetworkConfig objects
        """
        return self._favorites.copy()

    def add_favorite(self, config: NetworkConfig) -> bool:
        """
        Add a connection to favorites.

        Args:
            config: Connection configuration to save

        Returns:
            True if added, False if already exists
        """
        # Check for duplicates
        conn_id = config.get_connection_id()
        for fav in self._favorites:
            if fav.get_connection_id() == conn_id:
                return False  # Already exists

        self._favorites.append(config)
        self._save()
        return True

    def remove_favorite(self, config: NetworkConfig) -> bool:
        """
        Remove a connection from favorites.

        Args:
            config: Connection to remove

        Returns:
            True if removed, False if not found
        """
        conn_id = config.get_connection_id()
        original_len = len(self._favorites)

        self._favorites = [
            f for f in self._favorites
            if f.get_connection_id() != conn_id
        ]

        if len(self._favorites) < original_len:
            self._save()
            return True
        return False

    def is_favorite(self, config: NetworkConfig) -> bool:
        """
        Check if a connection is in favorites.

        Args:
            config: Connection to check

        Returns:
            True if in favorites
        """
        conn_id = config.get_connection_id()
        return any(
            f.get_connection_id() == conn_id
            for f in self._favorites
        )

    def update_favorite(self, old_config: NetworkConfig,
                       new_config: NetworkConfig) -> bool:
        """
        Update an existing favorite.

        Args:
            old_config: Original configuration
            new_config: New configuration

        Returns:
            True if updated, False if not found
        """
        old_id = old_config.get_connection_id()
        for i, fav in enumerate(self._favorites):
            if fav.get_connection_id() == old_id:
                self._favorites[i] = new_config
                self._save()
                return True
        return False

    # === History Management ===

    def get_history(self, limit: Optional[int] = None) -> List[ConnectionHistory]:
        """
        Get connection history, most recent first.

        Args:
            limit: Maximum number of entries (default: all)

        Returns:
            List of ConnectionHistory objects
        """
        sorted_history = sorted(
            self._history,
            key=lambda h: h.last_connected,
            reverse=True
        )

        if limit:
            return sorted_history[:limit]
        return sorted_history

    def record_connection(self, config: NetworkConfig):
        """
        Record a successful connection.

        Updates existing entry or creates new one.

        Args:
            config: Connection that was made
        """
        conn_id = config.get_connection_id()
        now = datetime.now()

        # Update existing entry
        for hist in self._history:
            if hist.config.get_connection_id() == conn_id:
                hist.last_connected = now
                hist.connect_count += 1
                self._save()
                return

        # Create new entry
        self._history.append(ConnectionHistory(
            config=config,
            last_connected=now,
            connect_count=1
        ))

        # Trim to max size
        if len(self._history) > self.MAX_HISTORY:
            # Keep most recent
            self._history = sorted(
                self._history,
                key=lambda h: h.last_connected,
                reverse=True
            )[:self.MAX_HISTORY]

        self._save()

    def clear_history(self):
        """Clear all connection history."""
        self._history = []
        self._save()

    def remove_from_history(self, config: NetworkConfig) -> bool:
        """
        Remove a specific entry from history.

        Args:
            config: Connection to remove

        Returns:
            True if removed
        """
        conn_id = config.get_connection_id()
        original_len = len(self._history)

        self._history = [
            h for h in self._history
            if h.config.get_connection_id() != conn_id
        ]

        if len(self._history) < original_len:
            self._save()
            return True
        return False

    # === Utility Methods ===

    def get_recent_configs(self, limit: int = 5) -> List[NetworkConfig]:
        """
        Get most recently used configurations.

        Args:
            limit: Maximum number to return

        Returns:
            List of NetworkConfig objects
        """
        history = self.get_history(limit)
        return [h.config for h in history]

    def search(self, query: str) -> List[NetworkConfig]:
        """
        Search favorites and history.

        Args:
            query: Search string (matches host, port, or name)

        Returns:
            Matching configurations
        """
        query_lower = query.lower()
        results = []
        seen_ids = set()

        # Search favorites first
        for config in self._favorites:
            if self._matches_query(config, query_lower):
                results.append(config)
                seen_ids.add(config.get_connection_id())

        # Then history
        for hist in self._history:
            config = hist.config
            if config.get_connection_id() not in seen_ids:
                if self._matches_query(config, query_lower):
                    results.append(config)
                    seen_ids.add(config.get_connection_id())

        return results

    def _matches_query(self, config: NetworkConfig, query: str) -> bool:
        """Check if config matches search query."""
        searchable = f"{config.host} {config.port} {config.name or ''}"
        return query in searchable.lower()


# Export public API
__all__ = ['ConnectionManager', 'ConnectionHistory']
```

### Acceptance Criteria

- [ ] ConnectionManager persists data via QSettings
- [ ] Favorites can be added/removed/checked
- [ ] History is recorded and limited to MAX_HISTORY
- [ ] History sorted by most recent first
- [ ] Search works across favorites and history
- [ ] Data survives application restart

---

## Task C.3: Extract Utility Classes

### Objective

Extract reusable utility classes from Serial Terminal's core.py.

### File

`core/core.py`

### Classes to Extract

1. **SettingsManager** - Application settings
2. **ResponsiveWindowManager** - Window sizing calculations

### Code

```python
#!/usr/bin/env python3
"""
Core Utilities for Network Terminal

Extracted from Serial Terminal with modifications for network use.
"""

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication
from dataclasses import dataclass
import json


@dataclass
class WindowConfig:
    """Configuration for window sizing and layout."""
    width: int
    height: int
    x: int
    y: int
    is_small_screen: bool
    min_width: int = 800
    min_height: int = 600


class SettingsManager:
    """
    Manages application settings using QSettings.

    Provides persistent storage for user preferences.
    """

    def __init__(self):
        """Initialize settings manager."""
        # Changed from SerialSplit to NetworkTerminal
        self.settings = QSettings("NetworkTerminal", "Settings")

    def get_show_launch_dialog(self) -> bool:
        """Get whether to show launch dialog on startup."""
        return self.settings.value("ui/show_launch_dialog", True, type=bool)

    def set_show_launch_dialog(self, show_dialog: bool):
        """Set whether to show launch dialog on startup."""
        self.settings.setValue("ui/show_launch_dialog", show_dialog)
        self.settings.sync()

    def get_last_config(self) -> dict:
        """Get last used connection configuration."""
        config_json = self.settings.value("connection/last_config", "{}")
        try:
            if isinstance(config_json, str):
                return json.loads(config_json)
        except json.JSONDecodeError:
            pass
        return {}

    def set_last_config(self, config: dict):
        """Save last used connection configuration."""
        self.settings.setValue("connection/last_config", json.dumps(config))
        self.settings.sync()

    def get_window_geometry(self) -> dict:
        """Get saved window geometry."""
        geometry_json = self.settings.value("ui/window_geometry", "{}")
        try:
            if isinstance(geometry_json, str):
                return json.loads(geometry_json)
        except json.JSONDecodeError:
            pass
        return {}

    def set_window_geometry(self, geometry: dict):
        """Save window geometry."""
        self.settings.setValue("ui/window_geometry", json.dumps(geometry))
        self.settings.sync()

    def get_value(self, key: str, default=None):
        """Get arbitrary setting value."""
        return self.settings.value(key, default)

    def set_value(self, key: str, value):
        """Set arbitrary setting value."""
        self.settings.setValue(key, value)
        self.settings.sync()


class ResponsiveWindowManager:
    """
    Manages responsive window sizing and layout decisions.

    Copied unchanged from Serial Terminal - completely reusable.
    """

    SMALL_SCREEN_WIDTH_THRESHOLD = 1024
    SMALL_SCREEN_HEIGHT_THRESHOLD = 768
    SMALL_SCREEN_WIDTH_RATIO = 0.95
    SMALL_SCREEN_HEIGHT_RATIO = 0.90
    LARGE_SCREEN_DEFAULT_WIDTH = 1200
    LARGE_SCREEN_DEFAULT_HEIGHT = 900
    ABSOLUTE_MIN_WIDTH = 960
    ABSOLUTE_MIN_HEIGHT = 600

    @classmethod
    def get_screen_info(cls):
        """Get primary screen geometry information."""
        screen = QApplication.primaryScreen()
        if not screen:
            return 1024, 768, 0, 0

        screen_geometry = screen.availableGeometry()
        return (
            screen_geometry.width(),
            screen_geometry.height(),
            screen_geometry.x(),
            screen_geometry.y()
        )

    @classmethod
    def is_small_screen(cls, screen_width: int, screen_height: int) -> bool:
        """Determine if screen should be considered small."""
        return (screen_width < cls.SMALL_SCREEN_WIDTH_THRESHOLD or
                screen_height < cls.SMALL_SCREEN_HEIGHT_THRESHOLD)

    @classmethod
    def calculate_main_window_config(cls) -> WindowConfig:
        """Calculate optimal window configuration for main window."""
        screen_width, screen_height, screen_x, screen_y = cls.get_screen_info()
        is_small = cls.is_small_screen(screen_width, screen_height)

        if is_small:
            window_width = min(
                max(screen_width * cls.SMALL_SCREEN_WIDTH_RATIO,
                    cls.ABSOLUTE_MIN_WIDTH),
                screen_width
            )
            window_height = min(
                max(screen_height * cls.SMALL_SCREEN_HEIGHT_RATIO,
                    cls.ABSOLUTE_MIN_HEIGHT),
                screen_height
            )
            x = screen_x + (screen_width - window_width) // 2
            y = screen_y + (screen_height - window_height) // 2
        else:
            window_width = cls.LARGE_SCREEN_DEFAULT_WIDTH
            window_height = cls.LARGE_SCREEN_DEFAULT_HEIGHT
            x = screen_x + 100
            y = screen_y + 100

        return WindowConfig(
            width=int(window_width),
            height=int(window_height),
            x=int(x),
            y=int(y),
            is_small_screen=is_small,
            min_width=cls.ABSOLUTE_MIN_WIDTH,
            min_height=cls.ABSOLUTE_MIN_HEIGHT
        )

    @classmethod
    def calculate_dialog_config(cls, preferred_width: int = 800,
                               preferred_height: int = 500) -> WindowConfig:
        """Calculate optimal window configuration for dialogs."""
        screen_width, screen_height, screen_x, screen_y = cls.get_screen_info()
        is_small = cls.is_small_screen(screen_width, screen_height)

        if is_small:
            window_width = min(screen_width * 0.9, preferred_width)
            window_height = min(screen_height * 0.8, preferred_height)
            min_width = 600
            min_height = 400
        else:
            window_width = preferred_width
            window_height = preferred_height
            min_width = preferred_width // 2
            min_height = preferred_height // 2

        x = screen_x + (screen_width - window_width) // 2
        y = screen_y + (screen_height - window_height) // 2

        return WindowConfig(
            width=int(window_width),
            height=int(window_height),
            x=int(x),
            y=int(y),
            is_small_screen=is_small,
            min_width=min_width,
            min_height=min_height
        )

    @classmethod
    def get_adaptive_font_size(cls, base_size: int, is_small_screen: bool) -> int:
        """Get adaptive font size based on screen size."""
        if is_small_screen:
            return max(base_size - 2, 10)
        return base_size


# Export public API
__all__ = [
    'WindowConfig',
    'SettingsManager',
    'ResponsiveWindowManager'
]
```

### Acceptance Criteria

- [ ] SettingsManager works with NetworkTerminal namespace
- [ ] ResponsiveWindowManager copied without changes
- [ ] WindowConfig dataclass available
- [ ] All methods work correctly

---

## Stream Deliverables

After completing Stream C, the following should be ready:

| Item | File | Status |
|------|------|--------|
| NetworkConfig | `core/network_config.py` | |
| ConnectionManager | `core/connection_manager.py` | |
| ConnectionHistory | `core/connection_manager.py` | |
| SettingsManager | `core/core.py` | |
| ResponsiveWindowManager | `core/core.py` | |
| WindowConfig | `core/core.py` | |
| Unit tests | `tests/test_*.py` | |

---

## Related Documents

- [Master PRD](../PRD_NetworkTerminal.md)
- [Interface Definitions](../reference/interface-definitions.md) - NetworkConfig interface
- [Stream A: Core Network](./stream-a-core-network.md) - Uses NetworkConfig
- [Stream B: UI Layer](./stream-b-ui-layer.md) - Uses ConnectionManager

---

[← Back to Master PRD](../PRD_NetworkTerminal.md) | [← Stream B](./stream-b-ui-layer.md) | [Phase 4 →](./phase-4-integration.md)
