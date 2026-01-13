"""
Core module for Serial Terminal application.
Provides serial port scanning, monitoring, and virtual port management.
"""

from .core import (
    PortStatus,
    WindowConfig,
    SerialPortInfo,
    Com0comPortPair,
    SerialPacketInfo,
    AdvancedStatistics,
    PortCapabilityAnalyzer,
    SettingsManager,
    ResponsiveWindowManager,
    PortScanner,
    PortConfig,
    Hub4comProcess,
    SerialPortMonitor,
    SerialPortTester,
)
from .com0com import DefaultConfig, Com0comProcess

__all__ = [
    # Enums
    "PortStatus",
    # Data classes
    "WindowConfig",
    "SerialPortInfo",
    "Com0comPortPair",
    "SerialPacketInfo",
    "AdvancedStatistics",
    "PortConfig",
    # Managers
    "PortCapabilityAnalyzer",
    "SettingsManager",
    "ResponsiveWindowManager",
    # Threads/Workers
    "PortScanner",
    "Hub4comProcess",
    "SerialPortMonitor",
    "SerialPortTester",
    # com0com
    "DefaultConfig",
    "Com0comProcess",
]
