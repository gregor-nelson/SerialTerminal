#!/usr/bin/env python3
"""
Serial port configuration data model.
"""

from dataclasses import dataclass


@dataclass
class SerialConfig:
    """Serial port configuration"""
    port: str
    baudrate: int = 115200
    databits: int = 8
    parity: str = 'N'
    stopbits: float = 1.0

    def get_display_string(self) -> str:
        """Get display string for status bar"""
        return f"{self.baudrate} {self.databits}{self.parity}{self.stopbits}"
