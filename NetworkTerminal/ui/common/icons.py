#!/usr/bin/env python3
"""
SVG icon definitions for the application.
Circular colored design matching application theme.
"""

from PyQt6.QtGui import QPalette, QIcon, QPainter, QPixmap
from PyQt6.QtCore import Qt
from PyQt6.QtSvg import QSvgRenderer


class Icons:
    """Icon set with circular colored SVG designs"""

    # Color constants
    BLUE_PRIMARY = "#0078D4"
    BLUE_STROKE = "#106EBE"
    GREEN_PRIMARY = "#28A745"
    GREEN_STROKE = "#1E7E34"
    PURPLE_PRIMARY = "#6F42C1"
    PURPLE_STROKE = "#5A2D91"
    RED_PRIMARY = "#DC3545"
    RED_STROKE = "#B02A37"
    GRAY_PRIMARY = "#6C757D"
    GRAY_STROKE = "#5A6268"

    @staticmethod
    def play(palette):
        """Play/Connect icon - blue circle with play symbol"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#0078D4" stroke="#106EBE" stroke-width="1"/>
            <path d="M12 10 L22 16 L12 22 Z" fill="#FFFFFF"/>
        </svg>"""

    @staticmethod
    def create(palette=None):
        """New/Add icon - blue circle with plus"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#0078D4" stroke="#106EBE" stroke-width="1"/>
            <path d="M16 8 L16 24 M8 16 L24 16" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def quick_setup():
        """Quick setup icon - green circle with lightning bolt"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#28A745" stroke="#1E7E34" stroke-width="1"/>
            <path d="M18 8 L10 16 L14 16 L14 24 L22 16 L18 16 Z" fill="#FFFFFF"/>
        </svg>"""

    @staticmethod
    def refresh(palette=None):
        """Refresh icon - green circle with refresh arrows"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#28A745" stroke="#1E7E34" stroke-width="1"/>
            <path d="M16 9 A7 7 0 1 1 9 16 A7 7 0 0 1 12.8 11.2" stroke="#FFFFFF" stroke-width="2.5" fill="none" stroke-linecap="round"/>
            <path d="M11 9 L15 9 L15 13" stroke="#FFFFFF" stroke-width="2.5" fill="none" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def remove():
        """Remove icon - red circle with X"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#DC3545" stroke="#B02A37" stroke-width="1"/>
            <path d="M10 10 L22 22 M22 10 L10 22" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def close():
        """Close icon - gray circle with X"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#6C757D" stroke="#5A6268" stroke-width="1"/>
            <path d="M10 10 L22 22 M22 10 L10 22" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def settings(palette):
        """Settings icon - purple circle with gear"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#6F42C1" stroke="#5A2D91" stroke-width="1"/>
            <g fill="#FFFFFF">
                <circle cx="16" cy="16" r="8"/>
                <rect x="15" y="6" width="2" height="3"/>
                <rect x="23" y="15" width="3" height="2"/>
                <rect x="15" y="23" width="2" height="3"/>
                <rect x="6" y="15" width="3" height="2"/>
                <rect x="21.5" y="8.5" width="2.1" height="2.1" transform="rotate(45 22.5 9.5)"/>
                <rect x="21.5" y="21.4" width="2.1" height="2.1" transform="rotate(45 22.5 22.5)"/>
                <rect x="8.4" y="21.4" width="2.1" height="2.1" transform="rotate(45 9.5 22.5)"/>
                <rect x="8.4" y="8.5" width="2.1" height="2.1" transform="rotate(45 9.5 9.5)"/>
                <circle cx="16" cy="16" r="3" fill="#6F42C1"/>
            </g>
        </svg>"""

    @staticmethod
    def terminal_settings(palette):
        """Terminal window icon for app icon"""
        color = palette.color(QPalette.ColorRole.ButtonText).name()
        return f"""<svg width="24" height="24" viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="18" rx="2" fill="none" stroke="{color}" stroke-width="2"/><path fill="{color}" d="M6 8l3 3-3 3m5 0h6"/></svg>"""

    @staticmethod
    def svg_to_icon(svg_str: str) -> QIcon:
        """Convert SVG string to QIcon"""
        # Create pixmap
        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.GlobalColor.transparent)

        # Render SVG
        renderer = QSvgRenderer(svg_str.encode('utf-8'))
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()

        return QIcon(pixmap)
