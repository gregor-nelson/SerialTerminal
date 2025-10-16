#!/usr/bin/env python3
"""
Terminal Application Constants
Following VirtualPortManager's minimal approach - trust Fusion theme for UI colors
"""

class TerminalColors:
    """Terminal text formatting colors only - all UI colors come from Fusion palette"""

    # Terminal text colors (for formatted output in terminal display)
    INCOMING = "#66bb6a"      # Green for incoming data
    OUTGOING = "#4dc3ff"      # Blue for outgoing data
    TIMESTAMP = "#999999"     # Gray for timestamps
    ERROR = "#ff6b6b"         # Red for errors
    STATUS = "#ffffff"        # White for status
    WARNING = "#cccccc"       # Light gray for warnings
    DEFAULT = "#e0e0e0"       # Default terminal text


class AppInfo:
    """Application metadata"""
    NAME = "Serial Terminal"
    VERSION = "1.0"
    ORG_NAME = "Serial Terminal"
    ORG_DOMAIN = "serialterminal.local"
