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
