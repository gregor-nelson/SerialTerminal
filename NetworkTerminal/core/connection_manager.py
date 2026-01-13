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
