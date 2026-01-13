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
