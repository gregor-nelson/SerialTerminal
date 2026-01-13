# NetworkTerminal Core Package
"""
Core module for NetworkTerminal.

This module provides the fundamental classes for network configuration,
connection management, and application settings.
"""

from .network_config import NetworkConfig, get_default_configs
from .connection_manager import ConnectionManager, ConnectionHistory
from .core import WindowConfig, SettingsManager, ResponsiveWindowManager

__all__ = [
    # Network Configuration
    'NetworkConfig',
    'get_default_configs',

    # Connection Management
    'ConnectionManager',
    'ConnectionHistory',

    # Utility Classes
    'WindowConfig',
    'SettingsManager',
    'ResponsiveWindowManager',
]
