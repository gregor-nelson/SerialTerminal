#!/usr/bin/env python3
"""
Terminal Stream Formatter Module

This module provides formatting capabilities for terminal stream data,
applying consistent styling and color-coding for serial port communication.
Follows the same professional styling as CommandFormatter and OutputLogFormatter.
"""

import threading
from PyQt6.QtGui import QTextCharFormat, QColor, QFont, QTextCursor
from PyQt6.QtWidgets import QTextEdit
from datetime import datetime
import re
from constants import TerminalColors
from ui.resources import resource_manager


class TerminalStreamFormatter:
    """
    Formats terminal stream data with color-coded data flow and consistent styling.
    
    Follows the same professional color scheme as CommandFormatter for UI consistency.
    """
    
    def __init__(self):
        """Initialize the formatter with color definitions for different data types."""
        # Auto-scroll preference
        self.auto_scroll_enabled = True
        self._scroll_lock = threading.Lock()
        
        # Terminal color scheme for data flow
        self.colors = {
            'incoming': TerminalColors.INCOMING,    # Green for incoming data
            'outgoing': TerminalColors.OUTGOING,    # Blue for outgoing data
            'timestamp': TerminalColors.TIMESTAMP,  # Gray for timestamps
            'separator': TerminalColors.TIMESTAMP,  # Gray for separators
            'error': TerminalColors.ERROR,          # Red for errors
            'status': TerminalColors.STATUS,        # White for status
            'default': TerminalColors.DEFAULT,      # Default terminal text
            'warning': TerminalColors.WARNING,      # Light gray for warnings
            'help': TerminalColors.STATUS,          # White for help messages
        }
        
        # Windows 10 Dark Mode NMEA Message Colors - Muted Terminal Palette with Subtle Variations
        self.nmea_colors = {
            # Navigation/Position messages - Muted Blue Family (subtle variations)
            'GGA': '#8bb5d9',     # Global Positioning System Fix Data - Base soft blue
            'GLL': '#7ba8cc',     # Geographic Position - Slightly darker blue
            'RMC': '#9bc2e6',     # Recommended Minimum Navigation Information - Slightly lighter blue
            'ZDA': '#85b0d6',     # Date & Time - Slightly grayer blue
            
            # Depth/Sonar messages - Muted Green Family (subtle variations)
            'DBS': '#90c695',     # Depth Below Surface - Base soft green
            'DBT': '#84b989',     # Depth Below Transducer - Slightly darker green
            'DPT': '#9cd3a1',     # Depth - Slightly lighter green
            'DEP': '#8ac093',     # Depth (alternate format) - Slightly grayer green
            'SONDEP': '#96cc99',  # Sonar Depth - Slightly more saturated green
            
            # Heading/Attitude messages - Muted Purple Family (subtle variations)
            'HDT': '#b4a7d6',     # Heading True - Base soft purple
            'HPR': '#a89ac9',     # Heading, Pitch, Roll - Slightly darker purple
            'PASHR': '#c0b4e3',   # Proprietary Attitude Sensor - Slightly lighter purple
            'THS': '#b1a3d3',     # True Heading and Status - Slightly grayer purple
            'HEV': '#aea0d0',     # Heave - Slightly more muted purple
            
            # Velocity/Motion messages - Muted Teal Family (subtle variations)
            'VBW': '#7fb8c4',     # Dual Ground/Water Speed - Base soft teal
            'VDR': '#73abb7',     # Set and Drift - Slightly darker teal
            'VHW': '#8bc5d1',     # Water Speed and Heading - Slightly lighter teal
            'VTG': '#79b5c1',     # Track Made Good and Ground Speed - Slightly grayer teal
            
            # Weather/Environmental messages - Muted Orange Family (subtle variations)
            'WIMDA': '#d4a574',   # Meteorological Composite - Base soft orange
            'WIMWD': '#c89866',   # Wind Direction and Speed - Slightly darker orange
            'WIMWV': '#e0b282',   # Wind Speed and Angle - Slightly lighter orange
            'MDA': '#d1a271',     # Meteorological Composite (short form) - Slightly grayer orange
            'MWD': '#cb9b69',     # Wind Direction and Speed (short form) - Slightly more muted orange
            'MWV': '#ddaf7f',     # Wind Speed and Angle (short form) - Slightly warmer orange
            
            # Satellite/GPS messages - Muted Yellow Family (subtle variations)
            'GSA': '#d4c875',     # GNSS DOP and Active Satellites - Base soft yellow
            'GST': '#c8bb69',     # GNSS Pseudorange Error Statistics - Slightly darker yellow
            'GSV': '#e0d581',     # GNSS Satellites in View - Slightly lighter yellow
            'GRS': '#d1c572',     # GNSS Range Residuals - Slightly grayer yellow
            
            # Proprietary messages - Muted Pink/Rose Family (subtle variations)
            'PSAT': '#c99bb3',    # Proprietary Satellite - Base soft rose
            'PSONNAV': '#bd8fa7', # Proprietary Navigation - Slightly darker rose
            'PSXN': '#d5a7bf',    # Proprietary System - Slightly lighter rose
            'PTNL': '#c698b0',    # Proprietary Trimble - Slightly grayer rose
            'PDWA': '#c295ad',    # Proprietary Dynamic Wayfinding - Slightly more muted rose
            
            # AIS messages - Muted Red Family
            'AIVDM': '#cc8888',   # AIS VDM message - Soft red
            
            # Other/miscellaneous messages - Neutral Gray Family (subtle variations)
            'DRU': '#a0a0a0',     # Dual Rudder - Light gray
            'ROV': '#959595',     # Remotely Operated Vehicle - Slightly darker gray
            'NMEA_UNKNOWN': '#888888',  # Unknown NMEA message - Medium gray
        }
        
        # Create all text formats upfront
        self.formats = {}
        self._create_formats()
        
        # Simple NMEA detection pattern
        self.nmea_pattern = re.compile(r'^\$([A-Z]{2})([A-Z]{3,})')
        self.proprietary_pattern = re.compile(r'^\$(P[A-Z]+)')
        self.ais_pattern = re.compile(r'^!AIVDM')
        
    def _create_formats(self):
        """Create QTextCharFormat objects for styling."""
        # Get monospace font from resource manager for consistency
        # The QFont already has fallback chain configured
        base_mono_font = resource_manager.get_monospace_font()

        # Default format
        self.formats['default'] = QTextCharFormat()
        self.formats['default'].setForeground(QColor(self.colors['default']))
        self.formats['default'].setFont(base_mono_font)

        # Create formats for each color type
        for color_type, color_value in self.colors.items():
            fmt = QTextCharFormat()
            fmt.setForeground(QColor(color_value))
            fmt.setFont(base_mono_font)
            self.formats[color_type] = fmt

        # Create formats for each NMEA type
        for nmea_type, color in self.nmea_colors.items():
            fmt = QTextCharFormat()
            fmt.setForeground(QColor(color))
            fmt.setFont(base_mono_font)
            self.formats[f'nmea_{nmea_type}'] = fmt
    
    def _get_format(self, format_name: str, bold: bool = False) -> QTextCharFormat:
        """Get a format by name, optionally with bold."""
        fmt = QTextCharFormat(self.formats.get(format_name, self.formats['default']))
        if bold:
            fmt.setFontWeight(QFont.Weight.Bold)
        return fmt
    
    def _detect_nmea_message_type(self, data: str) -> str:
        """
        Simple NMEA message type detection.
        
        Args:
            data: The incoming data string to analyze
            
        Returns:
            The NMEA message type or None if not detected
        """
        data = data.strip()
        
        # Check for AIS messages
        if self.ais_pattern.match(data):
            return 'AIVDM'
        
        # Check for proprietary messages
        if data.startswith('$'):
            # Special proprietary messages with known patterns
            if data.startswith('$PSAT') and ',HPR' in data:
                return 'PSAT'
            elif data.startswith('$PASHR'):
                return 'PASHR'
            elif data.startswith('$PSONNAV'):
                return 'PSONNAV'
            elif data.startswith('$PSXN'):
                return 'PSXN'
            elif data.startswith('$PTNL') and ',AVR' in data:
                return 'PTNL'
            elif data.startswith('$PDWA'):
                return 'PDWA'
            
            # Standard NMEA messages
            match = self.nmea_pattern.match(data)
            if match:
                message_type = match.group(2)
                
                # Handle weather messages that might have longer prefixes
                if message_type.endswith(('MWV', 'MWD', 'MDA')):
                    message_type = message_type[-3:]
                
                # Return the type if we have a color for it
                if message_type in self.nmea_colors:
                    return message_type
                
                # Special cases
                if message_type == 'DEP':
                    return 'DEP'
                
        return None
    
    def append_data(self, text_edit: QTextEdit, data: str, data_type: str = "incoming", 
                   show_timestamp: bool = True, detected_type: str = None):
        """
        Format and append serial data to the text edit widget.
        
        Args:
            text_edit: The QTextEdit widget to append to
            data: The data to format and display
            data_type: The data type (incoming, outgoing, status, error)
            show_timestamp: Whether to show timestamp prefix
            detected_type: Pre-detected NMEA type (optional)
        """
        if not text_edit or not data:
            return
        
        cursor = text_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        
        # Only update the cursor position if auto-scroll is enabled
        if self.is_auto_scroll_enabled():
            text_edit.setTextCursor(cursor)
        
        # Add newline if needed
        if text_edit.toPlainText() and not text_edit.toPlainText().endswith('\n'):
            cursor.insertText('\n')
        
        # Add timestamp if requested
        if show_timestamp:
            timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
            cursor.insertText(f"[{timestamp}] ", self.formats['timestamp'])
        
        # Add data type prefix for non-incoming data
        if data_type != "incoming":
            prefix_map = {
                'outgoing': 'Send',
                'status': 'Info',
                'error': 'Error',
                'warning': 'Warning',
                'help': 'Help'
            }
            prefix = prefix_map.get(data_type, data_type.upper())
            cursor.insertText(f"[{prefix}] ", self._get_format(data_type, bold=True))
        
        # Detect NMEA message type if not provided
        if detected_type is None and data_type == "incoming":
            detected_type = self._detect_nmea_message_type(data)
        
        # Choose format based on NMEA type or default data type
        if detected_type and detected_type in self.nmea_colors:
            data_format = self.formats.get(f'nmea_{detected_type}', self.formats['default'])
        else:
            data_format = self.formats.get(data_type, self.formats['default'])
        
        # Append the data
        cursor.insertText(data, data_format)
        cursor.insertText('\n')
        
        # Auto-scroll if enabled
        self._auto_scroll_if_enabled(text_edit)
    
    def append_separator(self, text_edit: QTextEdit, label: str = ""):
        """Add a visual separator line to the terminal stream."""
        if not text_edit:
            return
        
        cursor = text_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        text_edit.setTextCursor(cursor)
        
        # Add spacing
        if text_edit.toPlainText() and not text_edit.toPlainText().endswith('\n'):
            cursor.insertText('\n')
        cursor.insertText('\n')
        
        # Add separator line
        separator = "-" * 60
        cursor.insertText(separator + "\n", self.formats['separator'])
        
        # Add label if provided
        if label:
            cursor.insertText(f" {label} \n", self._get_format('status', bold=True))
            cursor.insertText(separator + "\n", self.formats['separator'])
        
        cursor.insertText('\n')
        
        # Auto-scroll if enabled
        self._auto_scroll_if_enabled(text_edit)
    
    def append_status(self, text_edit: QTextEdit, message: str, status_type: str = "status"):
        """Add a status message to the terminal stream."""
        self.append_data(text_edit, message, status_type, show_timestamp=True)
    
    def clear(self, text_edit: QTextEdit):
        """Clear all content from the text edit."""
        if text_edit:
            text_edit.clear()
    
    def format_connection_start(self, text_edit: QTextEdit, port_name: str, baud_rate: int):
        """Format the connection start message."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_separator(text_edit, f"Connection established - {timestamp}")
        self.append_status(text_edit, f"Serial port {port_name} ready ({baud_rate} bps)", "status")
        self.append_separator(text_edit)
    
    def format_connection_end(self, text_edit: QTextEdit, port_name: str):
        """Format the connection end message."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_separator(text_edit, f"Connection closed - {timestamp}")
        self.append_status(text_edit, f"Serial port {port_name} disconnected", "status")
        self.append_separator(text_edit)
    
    def _auto_scroll_if_enabled(self, text_edit: QTextEdit):
        """Auto-scroll to bottom if auto-scroll is enabled."""
        if not text_edit:
            return
        
        try:
            with self._scroll_lock:
                if not self.auto_scroll_enabled:
                    return
                
                scrollbar = text_edit.verticalScrollBar()
                if scrollbar:
                    # Only auto-scroll if we're near the bottom
                    max_value = scrollbar.maximum()
                    current_value = scrollbar.value()
                    
                    if max_value - current_value <= 10:
                        scrollbar.setValue(max_value)
        except:
            pass  # Silently ignore scroll errors
    
    def set_auto_scroll_enabled(self, enabled: bool):
        """Set auto-scroll enabled state."""
        with self._scroll_lock:
            self.auto_scroll_enabled = enabled
    
    def is_auto_scroll_enabled(self) -> bool:
        """Check if auto-scroll is enabled."""
        with self._scroll_lock:
            return self.auto_scroll_enabled
    
    def force_scroll_to_bottom(self, text_edit: QTextEdit):
        """Force scroll to bottom regardless of auto-scroll setting."""
        if text_edit:
            scrollbar = text_edit.verticalScrollBar()
            if scrollbar:
                scrollbar.setValue(scrollbar.maximum())