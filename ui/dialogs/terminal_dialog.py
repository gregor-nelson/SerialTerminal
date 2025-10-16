#!/usr/bin/env python3
"""
Windows 10 Style Serial Monitor Dialog
A complete terminal implementation with split pane support and terminal formatting.
Refactored to match main GUI menu implementation.
"""

import sys
import serial
import serial.tools.list_ports
from typing import Optional, Dict, List
from dataclasses import dataclass
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtSvg import QSvgRenderer
from datetime import datetime
import queue
import atexit


from ui.windows.terminal_formatter import TerminalStreamFormatter
from ui.resources import resource_manager
# No color constants needed - trust Fusion theme
from core.core import SerialPortInfo, PortScanner
from ui.components import RibbonToolbar


# Minimal icon SVG definitions
class Icons:
    """Icon set matching VirtualPortManager's circular colored design"""

    @staticmethod
    def play(palette):
        """Play/Connect icon - blue circle with play symbol"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#0078D4" stroke="#106EBE" stroke-width="1"/>
            <path d="M12 10 L22 16 L12 22 Z" fill="#FFFFFF"/>
        </svg>"""

    @staticmethod
    def create(palette):
        """New/Add icon - blue circle with plus (matching VirtualPortManager 'new')"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#0078D4" stroke="#106EBE" stroke-width="1"/>
            <path d="M16 8 L16 24 M8 16 L24 16" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def refresh(palette):
        """Refresh icon - green circle with refresh arrows (exact VirtualPortManager design)"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#28A745" stroke="#1E7E34" stroke-width="1"/>
            <path d="M16 9 A7 7 0 1 1 9 16 A7 7 0 0 1 12.8 11.2" stroke="#FFFFFF" stroke-width="2.5" fill="none" stroke-linecap="round"/>
            <path d="M11 9 L15 9 L15 13" stroke="#FFFFFF" stroke-width="2.5" fill="none" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def settings(palette):
        """Settings icon - purple circle with gear (exact VirtualPortManager configure design)"""
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


class SerialPortRegistry:
    _ports = set()
    
    @classmethod
    def register(cls, worker):
        cls._ports.add(worker)
    
    @classmethod
    def unregister(cls, worker):
        cls._ports.discard(worker)
    
    @classmethod
    def cleanup_all(cls):
        for worker in list(cls._ports):
            try:
                if hasattr(worker, 'serial_port') and worker.serial_port:
                    worker.serial_port.close()
            except:
                pass

# Register cleanup on program exit
atexit.register(SerialPortRegistry.cleanup_all)


# ===== DATA CLASSES =====
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

# ===== SERIAL WORKER =====
class SerialWorker(QThread):
    """Background thread for serial communication"""
    
    dataReceived = pyqtSignal(bytes)
    errorOccurred = pyqtSignal(str)
    connectionStateChanged = pyqtSignal(bool)  # True = connected, False = disconnected
    
    def __init__(self, config: SerialConfig):
        super().__init__()
        self.config = config
        self.serial_port: Optional[serial.Serial] = None
        self.running = False
        self.write_queue = queue.Queue()
        
    def run(self):
        """Main thread loop"""
        try:
            # Open serial port
            self.serial_port = serial.Serial(
                port=self.config.port,
                baudrate=self.config.baudrate,
                bytesize=self.config.databits,
                parity=self.config.parity,
                stopbits=self.config.stopbits,
                timeout=0.1
            )
            
            # ADD THIS LINE - Register with cleanup registry
            SerialPortRegistry.register(self)
            
            self.running = True
            self.connectionStateChanged.emit(True)
            
            while self.running:
                # Read data if available
                if self.serial_port.in_waiting > 0:
                    data = self.serial_port.read(self.serial_port.in_waiting)
                    self.dataReceived.emit(data)
                
                # Write data from queue
                try:
                    while not self.write_queue.empty():
                        data = self.write_queue.get_nowait()
                        self.serial_port.write(data)
                except queue.Empty:
                    pass
                    
        except serial.SerialException as e:
            self.errorOccurred.emit(str(e))
        finally:
            # ADD THIS LINE - Unregister from cleanup registry
            SerialPortRegistry.unregister(self)
            
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            self.connectionStateChanged.emit(False)  
                 
    def stop(self):
        """Stop the worker thread safely with graceful shutdown - blocks until complete"""
        self.running = False

        # Close the serial port to interrupt any blocking reads
        if self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.close()
            except Exception as e:
                print(f"Error closing port during stop: {e}")

        # Wait for thread to finish - blocking call ensures port is fully released
        if self.isRunning():
            if not self.wait(3000):  # 3 seconds timeout
                print(f"Warning: Serial thread did not stop cleanly for {self.config.port}")
                # Force termination if thread is stuck
                self.terminate()
                self.wait(1000)  # Wait for termination to complete   
    def write(self, data: bytes):
        """Queue data to be written"""
        if self.running:
            self.write_queue.put(data)

# ===== TERMINAL PANE =====
class TerminalPane(QWidget):
    """Individual terminal display with formatter integration"""
    
    focusChanged = pyqtSignal(bool)
    splitRequested = pyqtSignal(object, str)  # (source_pane, direction)
    closeRequested = pyqtSignal(object)  # source_pane
    
    def __init__(self, config: SerialConfig, parent=None, main_window=None):
        super().__init__(parent)
        self.config = config
        self.main_window = main_window
        self.formatter = TerminalStreamFormatter()
        self.serial_worker: Optional[SerialWorker] = None
        self.is_connected = False
        self.rx_bytes = 0
        self.tx_bytes = 0
        
        # Line buffering for proper data handling
        self.line_buffer = ""
        self.buffer_timer = QTimer()
        self.buffer_timer.setSingleShot(True)
        self.buffer_timer.timeout.connect(self._flush_buffer)
        
        # Display settings
        self.encoding = 'utf-8'
        self.hex_display_mode = False
        self.local_echo_enabled = True  # Default to enabled
        
        # Baud rate detection and error handling
        self.encoding_error_count = 0
        self.encoding_error_window = 50  # Track errors over last 50 packets
        self.encoding_error_threshold = 0.3  # 30% error rate threshold
        self.last_encoding_warning = 0
        self.encoding_warning_interval = 5.0  # Minimum 5 seconds between warnings
        self.suggested_baud_rates = [9600, 19200, 38400, 57600, 115200, 230400, 460800, 921600]
        self.data_packet_count = 0
        self.baud_rate_suggestion_shown = False
        self.consecutive_errors = 0  # Track consecutive errors
        self.max_consecutive_errors = 5  # Stop processing after this many consecutive errors
        
        # Help display management
        self.help_displayed = False
        self.auto_scroll_state_before_help = True
        
        self._setup_ui()
        self._setup_context_menu()
        
    def _setup_ui(self):
        """Setup the UI components"""
        # Main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Terminal display
        self.terminal = QTextEdit()
        self.terminal.setReadOnly(True)
        self.terminal.setFont(resource_manager.get_monospace_font(size=10))
        self.terminal.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        
        layout.addWidget(self.terminal)
        
        # Focus handling
        self.terminal.installEventFilter(self)
        
        # Install event filter for ESC key handling
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.installEventFilter(self)
        
    def _setup_context_menu(self):
        """Setup right-click context menu"""
        self.terminal.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.terminal.customContextMenuRequested.connect(self._show_context_menu)
        
    def _show_context_menu(self, position):
        """Show context menu at position"""
        menu = self._create_terminal_menu()
        menu.exec(self.terminal.mapToGlobal(position))
    
    def _create_terminal_menu(self) -> QMenu:
        """Create terminal menu matching main GUI style"""
        menu = QMenu(self)
        # Remove explicit styling to match primary GUI approach
        # Header
        # menu.addAction("Terminal Settings").setEnabled(False)
        menu.addSeparator()

        # Connection section
        # menu.addAction("Connection").setEnabled(False)
        
        if self.is_connected:
            disconnect = menu.addAction("Disconnect")
            disconnect.triggered.connect(self.disconnect)
        else:
            connect = menu.addAction("Connect")
            connect.triggered.connect(self.connect)

        # Display Settings section
        # menu.addAction("Display Settings").setEnabled(False)
        menu.addSeparator()
        auto_scroll = menu.addAction(
            self.checkbox_icon(self.formatter.is_auto_scroll_enabled()), 
            "Auto-scroll"
        )
        auto_scroll.triggered.connect(lambda: self._toggle_auto_scroll(not self.formatter.is_auto_scroll_enabled()))
        
        hex_mode = menu.addAction(
            self.checkbox_icon(self.hex_display_mode), 
            "Hex Display Mode"
        )
        hex_mode.triggered.connect(lambda: self._toggle_hex_mode(not self.hex_display_mode))
        
        local_echo = menu.addAction(
            self.checkbox_icon(self.local_echo_enabled), 
            "Local Echo"
        )
        local_echo.triggered.connect(lambda: self._toggle_local_echo(not self.local_echo_enabled))
        
        menu.addSeparator()
        
        # Terminal Options section
        # menu.addAction("Terminal Options").setEnabled(False)
        
        # Font size submenu
        font_menu = menu.addMenu("Font Size")
        self._create_font_size_menu(font_menu)
        
        # Baud rate submenu
        baud_menu = menu.addMenu("Baud Rate")
        self._create_baud_rate_menu(baud_menu)
        
        # COM port submenu
        com_menu = menu.addMenu("Switch COM Port")
        self._create_com_port_menu(com_menu)
        
        clear = menu.addAction("Clear Terminal")
        clear.triggered.connect(self._clear_terminal)
        
        reset_detection = menu.addAction("Reset Baud Rate Detection")
        reset_detection.triggered.connect(self.reset_baud_rate_detection)
        
        menu.addSeparator()
        
        # Pane Management section
        # menu.addAction("Pane Management").setEnabled(False)
        
        split_v = menu.addAction("Split Pane Vertically")
        split_v.setShortcut("Alt+Shift+-")
        split_v.triggered.connect(lambda: self.splitRequested.emit(self, 'vertical'))
        
        split_h = menu.addAction("Split Pane Horizontally")
        split_h.setShortcut("Alt+Shift++")
        split_h.triggered.connect(lambda: self.splitRequested.emit(self, 'horizontal'))
        
        close = menu.addAction("Close Pane")
        close.setShortcut("Ctrl+Shift+W")
        close.triggered.connect(lambda: self.closeRequested.emit(self))
        
        menu.addSeparator()
        
        # Edit Actions section
        # menu.addAction("Edit Actions").setEnabled(False)
        
        copy = menu.addAction("Copy")
        copy.setShortcut("Ctrl+C")
        copy.triggered.connect(self.terminal.copy)
        copy.setEnabled(self.terminal.textCursor().hasSelection())
        
        select_all = menu.addAction("Select All")
        select_all.setShortcut("Ctrl+A")
        select_all.triggered.connect(self.terminal.selectAll)
        
        scroll_bottom = menu.addAction("Scroll to Bottom")
        scroll_bottom.triggered.connect(self._scroll_to_bottom)

        menu.addSeparator()

        # Virtual Port Manager section
        virtual_port_mgr = menu.addAction("Virtual Port Manager")
        virtual_port_mgr.triggered.connect(self._show_virtual_port_manager)

        menu.addSeparator()

        # Help section
        help_action = menu.addAction("Help")
        help_action.triggered.connect(self._show_help)
        
        return menu
    
    def _show_help(self):
        """Show help inline in the terminal window"""
        # Prevent help spam
        if self.help_displayed:
            self.formatter.append_status(
                self.terminal,
                "Help already displayed. Press ESC to return to auto-scroll mode.",
                "warning"
            )
            return
            
        # Store current auto-scroll state
        self.auto_scroll_state_before_help = self.formatter.is_auto_scroll_enabled()
        
        # Disable auto-scroll to keep help visible
        self.formatter.set_auto_scroll_enabled(False)
        
        # Mark help as displayed
        self.help_displayed = True
        
        # Display help with separators
        self.formatter.append_separator(self.terminal, "TERMINAL HELP")
        
        # Add initial [HELP] marker
        self.formatter.append_status(
            self.terminal,
            "Start Of Help Content",
            "help"
        )
        
        help_sections = [
            ("CONNECTION", [
                "- Auto-detects baud rate",
                "- Supports standard serial protocols",
                "- Real-time data display"
            ]),
            ("DISPLAY OPTIONS", [
                "- Auto-scroll: Enabled by default",
                "- Hex Display Mode: Show data as hex",
                "- Local Echo: Echo typed characters"
            ]),
            ("DEFAULTS", [
                "- Baud Rate: 115200",
                "- Data Bits: 8",
                "- Parity: None",
                "- Stop Bits: 1",
                "- Font Size: 10pt"
            ]),
            ("KEYBOARD SHORTCUTS", [
                "Navigation:",
                "- Alt+Arrow Keys: Navigate between panes",
                "- Ctrl+Tab: Next tab",
                "- Ctrl+Shift+Tab: Previous tab",
                "",
                "Pane Management:",
                "- Alt+Shift+-: Split pane vertically",
                "- Alt+Shift++: Split pane horizontally",
                "- Ctrl+Shift+W: Close current pane",
                "",
                "Terminal Actions:",
                "- Ctrl+C: Copy selected text",
                "- Ctrl+A: Select all text",
                "- Ctrl++: Increase font size",
                "- Ctrl+-: Decrease font size",
                "",
                "Window Management:",
                "- Ctrl+N: New connection",
                "- Ctrl+W: Close current tab",
                "- F1: Show this help"
            ])
        ]
        
        cursor = self.terminal.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        
        for section_title, section_items in help_sections:
            # Add section header without [HELP] prefix
            cursor.insertText(f"\n{section_title}\n", self.formatter._get_format('help', bold=True))
            
            # Add section items without [HELP] prefix
            for item in section_items:
                if item:  # Skip empty lines
                    cursor.insertText(f"{item}\n", self.formatter._get_format('help'))
                else:
                    # Add empty line for spacing
                    cursor.insertText("\n")
            
            # Add spacing between sections
            cursor.insertText("\n")
        
        # Add final [HELP] marker
        self.formatter.append_status(
            self.terminal,
            "End Of Help Content - Press ESC to return to auto-scroll mode",
            "help"
        )
        
        # End help display
        self.formatter.append_separator(self.terminal)
    
    def _dismiss_help(self):
        """Dismiss help display and restore auto-scroll state"""
        if not self.help_displayed:
            return
            
        # Mark help as dismissed
        self.help_displayed = False
        
        # Restore auto-scroll state
        self.formatter.set_auto_scroll_enabled(self.auto_scroll_state_before_help)
        
        # Show confirmation message
        self.formatter.append_status(
            self.terminal,
            "Help dismissed - Auto-scroll restored",
            "status"
        )
        
        # Auto-scroll to bottom if enabled
        if self.auto_scroll_state_before_help:
            self.formatter.force_scroll_to_bottom(self.terminal)

    def _show_virtual_port_manager(self):
        """Open the Virtual Port Manager dialog"""
        from ui.dialogs.virtual_port_dialog import VirtualPortDialog

        dialog = VirtualPortDialog(self)
        dialog.exec()  # Modal dialog

    def _create_font_size_menu(self, menu: QMenu):
        """Create font size submenu matching main GUI pattern"""
        # Common sizes
        common_sizes = [8, 10, 12, 14, 16]
        current_size = self.terminal.font().pointSize()
        
        for size in common_sizes:
            action = menu.addAction(f"{size}pt")
            action.triggered.connect(lambda checked, s=size: self._set_font_size(s))
            if size == current_size:
                action.setIcon(self.checkbox_icon(True))
        
        menu.addSeparator()
        
        # All sizes submenu
        all_sizes_menu = menu.addMenu("All Sizes")
        for size in range(6, 25):
            if size not in common_sizes:
                action = all_sizes_menu.addAction(f"{size}pt")
                action.triggered.connect(lambda checked, s=size: self._set_font_size(s))
                if size == current_size:
                    action.setIcon(self.checkbox_icon(True))
        
        menu.addSeparator()
        
        # Quick actions
        increase_font = menu.addAction("Increase")
        increase_font.setShortcut("Ctrl++")
        increase_font.triggered.connect(self._increase_font_size)
        
        decrease_font = menu.addAction("Decrease")
        decrease_font.setShortcut("Ctrl+-")
        decrease_font.triggered.connect(self._decrease_font_size)
        
        menu.addSeparator()
        
        reset_font = menu.addAction("Reset to Default")
        reset_font.triggered.connect(self._reset_font_size)
    
    def _create_baud_rate_menu(self, menu: QMenu):
        """Create baud rate submenu matching font size pattern"""
        # Standard baud rates
        standard_rates = [9600, 19200, 38400, 57600, 115200, 230400, 460800, 921600]
        current_baud = self.config.baudrate
        
        for rate in standard_rates:
            action = menu.addAction(f"{rate}")
            action.triggered.connect(lambda checked, r=rate: self._set_baud_rate(r))
            if rate == current_baud:
                action.setIcon(self.checkbox_icon(True))
    
    def _create_com_port_menu(self, menu: QMenu):
        """Create COM port submenu with available ports and current selection indicator"""
        try:
            # Get available ports using PortScanner
            scanner = PortScanner()
            available_ports = scanner.scan_registry_ports()
            
            # Fall back to serial.tools.list_ports if registry scan fails
            if not available_ports:
                try:
                    import serial.tools.list_ports
                    serial_ports = serial.tools.list_ports.comports()
                    # Convert to our expected format
                    available_ports = []
                    for port in serial_ports:
                        # Create a simple object with the required attributes
                        class SimplePort:
                            def __init__(self, device, description):
                                self.port_name = device
                                self.device_name = description or "Unknown"
                                self.is_moxa = False
                                self.port_type = "Hardware"
                        
                        available_ports.append(SimplePort(port.device, port.description))
                except Exception:
                    available_ports = []
            
            current_port = self.config.port
            
            # Get connected ports from main window if available
            connected_ports = set()
            if self.main_window:
                try:
                    connected_ports = self.main_window.get_connected_ports()
                except Exception:
                    pass  # Fallback if main window method not available
            
            if not available_ports:
                no_ports = menu.addAction("No ports available")
                no_ports.setEnabled(False)
                return
            
            # Add ports to menu with enhanced display format
            for port in available_ports:
                display_name = self._create_com_port_display_text(port)
                
                # Check if port is in use by another pane
                is_in_use_by_other = (port.port_name in connected_ports and 
                                     port.port_name != current_port)
                
                if is_in_use_by_other:
                    display_name += " (In Use)"
                
                action = menu.addAction(display_name)
                action.triggered.connect(lambda checked, port_name=port.port_name: self._set_com_port(port_name))
                
                # Show checkbox for current port
                if port.port_name == current_port:
                    action.setIcon(self.checkbox_icon(True))
                
                # Disable if port is in use by another pane
                if is_in_use_by_other:
                    action.setEnabled(False)
                
        except Exception as e:
            error_action = menu.addAction(f"Error scanning ports: {str(e)}")
            error_action.setEnabled(False)
    
    def _create_com_port_display_text(self, port):
        """Create display text for COM port with type indicators"""
        display_text = port.port_name
        
        if hasattr(port, 'is_moxa') and port.is_moxa:
            display_text += "  •  Moxa"
            if hasattr(port, 'device_name') and port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        elif hasattr(port, 'port_type') and port.port_type.startswith("Virtual"):
            virtual_type = port.port_type.split(' ')[1] if ' ' in port.port_type else "Virtual"
            display_text += f"  •  {virtual_type} Port"
        else:
            display_text += "  •  Hardware Port"
            if hasattr(port, 'device_name') and port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        
        return display_text
    
    def checkbox_icon(self, checked: bool) -> QIcon:
        """Generate checkbox icon using palette colors like VirtualPortManager"""
        palette = self.palette()
        border_color = palette.color(QPalette.ColorRole.Mid).name()
        bg_color = palette.color(QPalette.ColorRole.Base).name()
        check_color = palette.color(QPalette.ColorRole.Highlight).name()

        if checked:
            svg = f'''<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
                <rect x="0.5" y="0.5" width="15" height="15" fill="{border_color}" stroke="{border_color}" stroke-width="1"/>
                <rect x="2" y="2" width="12" height="12" fill="{bg_color}"/>
                <path d="M4 8l2 2 6-6" stroke="{check_color}" stroke-width="1.5" fill="none" stroke-linecap="round"/>
            </svg>'''
        else:
            svg = f'''<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
                <rect x="0.5" y="0.5" width="15" height="15" fill="{border_color}" stroke="{border_color}" stroke-width="1"/>
            </svg>'''
        
        renderer = QSvgRenderer(QByteArray(svg.encode()))
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return QIcon(pixmap)
    
    
    def eventFilter(self, obj, event):
        """Handle focus events and keyboard input for local echo"""
        if obj == self.terminal:
            if event.type() == QEvent.Type.FocusIn:
                self.focusChanged.emit(True)
            elif event.type() == QEvent.Type.FocusOut:
                self.focusChanged.emit(False)
            elif event.type() == QEvent.Type.KeyPress and self.local_echo_enabled:
                return self._handle_key_press(event)
        elif obj == self and event.type() == QEvent.Type.KeyPress:
            # Handle ESC key for help dismissal
            if event.key() == Qt.Key.Key_Escape and self.help_displayed:
                self._dismiss_help()
                return True
        return super().eventFilter(obj, event)
    
    def _handle_key_press(self, event):
        """Handle key press events for local echo"""
        if not self.is_connected:
            return False
        
        key = event.key()
        text = event.text()
        
        # Handle special keys
        if key == Qt.Key.Key_Return or key == Qt.Key.Key_Enter:
            # Send CRLF and create new line in display
            data_to_send = "\r\n"
            self._send_raw_data(data_to_send)
            self._echo_local_data(data_to_send)
            return True
        elif key == Qt.Key.Key_Backspace:
            # For now, just ignore backspace in local echo mode
            return True
        elif key == Qt.Key.Key_Tab:
            # Send tab character and echo locally
            data_to_send = "\t"
            self._send_raw_data(data_to_send)
            self._echo_local_data(data_to_send)
            return True
        elif text and text.isprintable():
            # Send printable characters and echo locally
            self._send_raw_data(text)
            self._echo_local_data(text)
            return True
        
        # Let other keys pass through normally
        return False
    
    def _send_raw_data(self, data: str):
        """Send raw data to serial port without local echo formatting"""
        if self.serial_worker and self.is_connected:
            try:
                bytes_data = data.encode(self.encoding)
                self.serial_worker.write(bytes_data)
                self.tx_bytes += len(bytes_data)
            except UnicodeEncodeError as e:
                self.formatter.append_status(
                    self.terminal,
                    f"Send encoding error: {str(e)}",
                    "error"
                )
    
    def _echo_local_data(self, data: str):
        """Echo data locally without timestamps or formatting"""
        cursor = self.terminal.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        
        # Set color for local echo (different from received data)
        format = QTextCharFormat()
        format.setForeground(QColor("#90EE90"))  # Light green for local echo
        cursor.setCharFormat(format)
        
        # Insert the text
        cursor.insertText(data)
        
        # Auto-scroll if enabled
        if self.formatter.is_auto_scroll_enabled():
            self.formatter.force_scroll_to_bottom(self.terminal)
    
    def connect(self):
        """Connect to serial port"""
        if not self.serial_worker:
            self.serial_worker = SerialWorker(self.config)
            # Use QueuedConnection to ensure thread-safe UI updates
            self.serial_worker.dataReceived.connect(
                self._on_data_received, Qt.ConnectionType.QueuedConnection
            )
            self.serial_worker.errorOccurred.connect(
                self._on_error, Qt.ConnectionType.QueuedConnection
            )
            self.serial_worker.connectionStateChanged.connect(
                self._on_connection_state_changed, Qt.ConnectionType.QueuedConnection
            )
            self.serial_worker.start()
            
    def disconnect(self):
        """Disconnect from serial port - ensures complete cleanup"""
        if self.serial_worker:
            self.serial_worker.stop()  # Now blocks until thread fully terminates
            self.serial_worker = None

        # Clear buffer and stop timer
        self.line_buffer = ""
        self.buffer_timer.stop()
    
    def cleanup(self):
        """Single point of cleanup for terminal pane"""
        if self.serial_worker:
            # Disconnect signals first to prevent crashes
            try:
                self.serial_worker.dataReceived.disconnect()
                self.serial_worker.errorOccurred.disconnect()
                self.serial_worker.connectionStateChanged.disconnect()
            except (TypeError, RuntimeError):
                pass  # Signals already disconnected
            
            # Force stop worker and ensure thread termination
            self.serial_worker.stop()
            
            # Force thread termination if it doesn't stop gracefully
            if not self.serial_worker.wait(2000):  # Wait 2 seconds
                self.serial_worker.terminate()
                self.serial_worker.wait(1000)  # Wait another second for termination
            
            self.serial_worker = None
        
        # Clear buffer and stop timer
        self.line_buffer = ""
        self.buffer_timer.stop()
        
        # Mark as disconnected to prevent further operations
        self.is_connected = False
            
    def _on_data_received(self, data: bytes):
        """Handle received data with proper line buffering"""
        # This method is called from worker thread via Qt signal/slot
        # Qt automatically handles thread safety for signal/slot connections
        try:
            self.rx_bytes += len(data)
            
            # Handle hex display mode
            if self.hex_display_mode:
                hex_data = ' '.join(f'{b:02X}' for b in data)
                ascii_data = ''.join(chr(b) if 32 <= b < 127 else '.' for b in data)
                formatted_data = f"HEX: {hex_data} | ASCII: {ascii_data}"
                
                self.formatter.append_data(
                    self.terminal,
                    formatted_data,
                    "incoming",
                    show_timestamp=True
                )
                return
            
            # Decode bytes with proper error handling
            try:
                text_data = data.decode(self.encoding)
                # Reset consecutive error count on successful decode
                self.consecutive_errors = 0
                # Gradually reduce error count on successful decodes
                if self.encoding_error_count > 0:
                    self.encoding_error_count = max(0, self.encoding_error_count - 1)
            except UnicodeDecodeError:
                self.consecutive_errors += 1
                
                # If too many consecutive errors, temporarily pause processing
                if self.consecutive_errors >= self.max_consecutive_errors:
                    self._handle_excessive_errors()
                    return
                
                # Try with replacement characters
                text_data = data.decode(self.encoding, errors='replace')
                self._handle_encoding_error()
                
                # Skip processing obviously garbled data
                if self._is_data_garbled(text_data):
                    return
            
            # Normalize line endings
            text_data = text_data.replace('\r\n', '\n').replace('\r', '\n')
            
            # Add to line buffer
            self.line_buffer += text_data
            
            # Process complete lines
            lines = self.line_buffer.split('\n')
            
            # Display all complete lines
            for line in lines[:-1]:
                if line:  # Only display non-empty lines
                    self.formatter.append_data(
                        self.terminal,
                        line,
                        "incoming",
                        show_timestamp=True
                    )
            
            # Keep the last (potentially incomplete) line
            self.line_buffer = lines[-1]
            
            # Start timer for incomplete lines
            if self.line_buffer:
                self.buffer_timer.stop()
                self.buffer_timer.start(1000)  # 1 second timeout
                
        except (UnicodeDecodeError, UnicodeError) as e:
            self.formatter.append_status(
                self.terminal,
                f"Text encoding error: {str(e)}",
                "error"
            )
        except (ValueError, TypeError) as e:
            self.formatter.append_status(
                self.terminal,
                f"Data format error: {str(e)}",
                "error"
            )
        except Exception as e:
            # Log unexpected errors but don't crash
            print(f"Unexpected data processing error: {e}")
            self.formatter.append_status(
                self.terminal,
                "Unexpected data processing error occurred",
                "error"
            )
            
    def _on_error(self, error_msg: str):
        """Handle serial errors"""
        self.formatter.append_status(self.terminal, error_msg, "error")
        
    def _on_connection_state_changed(self, connected: bool):
        """Handle connection state changes"""
        self.is_connected = connected
        # Notify main window if this pane is active
        if self.main_window and hasattr(self.main_window, '_update_ribbon_connection_state'):
            self.main_window._update_ribbon_connection_state()
        if connected:
            # Reset baud rate detection on new connection
            self.reset_baud_rate_detection()
            self.formatter.format_connection_start(
                self.terminal,
                self.config.port,
                self.config.baudrate
            )
        else:
            self.formatter.format_connection_end(self.terminal, self.config.port)
            
    def send_data(self, data: str):
        """Send data to serial port"""
        if self.serial_worker and self.is_connected:
            try:
                bytes_data = data.encode(self.encoding)
                self.serial_worker.write(bytes_data)
                self.tx_bytes += len(bytes_data)
                self.formatter.append_data(
                    self.terminal,
                    data.strip(),
                    "outgoing",
                    show_timestamp=True
                )
            except UnicodeEncodeError as e:
                self.formatter.append_status(
                    self.terminal,
                    f"Send encoding error: {str(e)}",
                    "error"
                )
            
    def get_status_info(self) -> str:
        """Get status information for status bar"""
        status = "Disconnected"
        if self.is_connected:
            status = "Connected"
        
        rx_str = self._format_bytes(self.rx_bytes)
        tx_str = self._format_bytes(self.tx_bytes)
        
        echo_indicator = " | Local Echo: ON" if self.local_echo_enabled else " | Local Echo: OFF"
        
        return f"{self.config.port}: {status} | {self.config.get_display_string()} | RX: {rx_str} | TX: {tx_str}{echo_indicator}"
    
    def _format_bytes(self, bytes_count: int) -> str:
        """Format byte count for display"""
        if bytes_count < 1024:
            return f"{bytes_count}B"
        elif bytes_count < 1024 * 1024:
            return f"{bytes_count / 1024:.1f}KB"
        else:
            return f"{bytes_count / (1024 * 1024):.1f}MB"
    
    def _flush_buffer(self):
        """Flush remaining data in buffer"""
        if self.line_buffer:
            self.formatter.append_data(
                self.terminal,
                self.line_buffer,
                "incoming",
                show_timestamp=True
            )
            self.line_buffer = ""
    
    def _toggle_auto_scroll(self, enabled: bool):
        """Toggle auto-scroll with formatter integration"""
        self.formatter.set_auto_scroll_enabled(enabled)
        if enabled:
            self.formatter.force_scroll_to_bottom(self.terminal)
    
    def _toggle_hex_mode(self, enabled: bool):
        """Toggle hex display mode"""
        self.hex_display_mode = enabled
        if enabled:
            self.formatter.append_status(
                self.terminal,
                "Hex display mode enabled",
                "status"
            )
        else:
            self.formatter.append_status(
                self.terminal,
                "Hex display mode disabled",
                "status"
            )
    
    def _toggle_local_echo(self, enabled: bool):
        """Toggle local echo mode"""
        self.local_echo_enabled = enabled
        if enabled:
            self.formatter.append_status(
                self.terminal,
                "Local echo enabled - Start typing to send data",
                "status"
            )
        else:
            self.formatter.append_status(
                self.terminal,
                "Local echo disabled - Terminal is read-only",
                "status"
            )
        
        # Trigger status bar update
        self.focusChanged.emit(True)
    
    def _scroll_to_bottom(self):
        """Manually scroll to bottom"""
        self.formatter.force_scroll_to_bottom(self.terminal)
    
    def _clear_terminal(self):
        """Clear the terminal display with proper formatter integration"""
        self.formatter.clear(self.terminal)
        
        # Clear the data buffer as well
        self.line_buffer = ""
        self.buffer_timer.stop()
        
        # Reset help display state
        self.help_displayed = False
        if not self.auto_scroll_state_before_help:
            self.formatter.set_auto_scroll_enabled(self.auto_scroll_state_before_help)
        
        # Reset error detection but keep connection state
        if self.is_connected:
            self.reset_baud_rate_detection()
            self.formatter.append_separator(self.terminal, "Terminal cleared")
    
    def _set_font_size(self, size: int):
        """Set terminal font size"""
        font = self.terminal.font()
        font.setPointSize(size)
        self.terminal.setFont(font)
    
    def _increase_font_size(self):
        """Increase terminal font size"""
        font = self.terminal.font()
        if font.pointSize() < 24:
            font.setPointSize(font.pointSize() + 1)
            self.terminal.setFont(font)
    
    def _decrease_font_size(self):
        """Decrease terminal font size"""
        font = self.terminal.font()
        if font.pointSize() > 8:
            font.setPointSize(font.pointSize() - 1)
            self.terminal.setFont(font)
    
    def _reset_font_size(self):
        """Reset font size to default"""
        self.terminal.setFont(resource_manager.get_monospace_font(size=10))
    
    def _set_baud_rate(self, baud_rate: int):
        """Set terminal baud rate with graceful automatic reconnection"""
        if self.config.baudrate == baud_rate:
            return  # Already at this baud rate
        
        old_baud = self.config.baudrate
        self.config.baudrate = baud_rate
        
        # Reset baud rate detection when rate changes
        self.reset_baud_rate_detection()
        
        if self.is_connected and self.serial_worker:
            # For active connections, always do graceful reconnect for reliability
            self.formatter.append_status(
                self.terminal,
                f"Switching baud rate from {old_baud} to {baud_rate}...",
                "info"
            )
            
            # Store connection state for seamless reconnection
            port_name = self.config.port
            was_monitoring = getattr(self, 'monitoring_active', False)
            
            try:
                # Graceful disconnect
                if was_monitoring:
                    try:
                        self._stop_monitoring()
                    except AttributeError:
                        pass  # Monitoring not implemented yet
                self.disconnect()
                
                # Brief pause to ensure port is released (1.5s for Windows COM port cleanup)
                from PyQt6.QtCore import QTimer
                QTimer.singleShot(1500, lambda: self._complete_baud_rate_change(
                    baud_rate, old_baud, port_name, was_monitoring
                ))
                
            except Exception as e:
                self.formatter.append_status(
                    self.terminal,
                    f"Error during disconnect: {str(e)}",
                    "error"
                )
                # Revert baud rate on failure
                self.config.baudrate = old_baud
        else:
            # Not connected, just update the setting
            self.formatter.append_status(
                self.terminal,
                f"Baud rate set to {baud_rate} (will apply on next connection)",
                "info"
            )
    
    def _complete_baud_rate_change(self, baud_rate: int, old_baud: int, port_name: str, was_monitoring: bool):
        """Complete the baud rate change after disconnect delay"""
        try:
            # Reconnect with new baud rate
            self.connect()

            # Restore monitoring if it was active
            if was_monitoring:
                try:
                    self._start_monitoring()
                except AttributeError:
                    pass  # Monitoring not implemented yet

            self.formatter.append_status(
                self.terminal,
                f"Successfully switched to {baud_rate} baud",
                "success"
            )

        except Exception as reconnect_error:
            # Provide detailed error diagnostics
            error_msg = str(reconnect_error).lower()
            if "access is denied" in error_msg or "permission denied" in error_msg:
                diagnostic = f"Port {port_name} is still in use or locked by another application"
            elif "cannot find" in error_msg or "does not exist" in error_msg:
                diagnostic = f"Port {port_name} is no longer available (device disconnected?)"
            elif "resource busy" in error_msg:
                diagnostic = f"Port {port_name} was not fully released - try again in a moment"
            else:
                diagnostic = f"Port {port_name} error: {str(reconnect_error)}"

            self.formatter.append_status(
                self.terminal,
                f"Failed to reconnect at {baud_rate} baud - {diagnostic}",
                "error"
            )

            # Attempt to restore original connection
            self.config.baudrate = old_baud
            try:
                self.connect()
                if was_monitoring:
                    try:
                        self._start_monitoring()
                    except AttributeError:
                        pass  # Monitoring not implemented yet
                self.formatter.append_status(
                    self.terminal,
                    f"Restored connection at original baud rate {old_baud}",
                    "warning"
                )
            except Exception as restore_error:
                self.formatter.append_status(
                    self.terminal,
                    f"Failed to restore original connection: {str(restore_error)}",
                    "error"
                )
    
    def _set_com_port(self, new_port: str):
        """Set terminal COM port with graceful automatic reconnection"""
        if self.config.port == new_port:
            return  # Already connected to this port
        
        # Check if new port is already in use by another pane
        if self.main_window:
            try:
                connected_ports = self.main_window.get_connected_ports()
                if new_port in connected_ports:
                    self.formatter.append_status(
                        self.terminal,
                        f"Port {new_port} is already in use by another pane",
                        "error"
                    )
                    return
            except Exception:
                pass  # Continue if port check fails
        
        old_port = self.config.port
        self.config.port = new_port
        
        # Reset baud rate detection when port changes
        self.reset_baud_rate_detection()
        
        # Always cleanup existing worker before attempting new connection
        if self.serial_worker:
            self.disconnect()
        
        # Show connecting message
        self.formatter.append_status(
            self.terminal,
            f"Connecting to {new_port}...",
            "info"
        )
        
        try:
            # Brief pause to ensure port is released, then connect (1.5s for Windows COM port cleanup)
            from PyQt6.QtCore import QTimer
            QTimer.singleShot(1500, self.connect)
            
        except Exception as connect_error:
            self.formatter.append_status(
                self.terminal,
                f"Failed to start connection to {new_port}: {str(connect_error)}",
                "error"
            )
            
            # Revert to original port on failure
            self.config.port = old_port
    
    def _complete_com_port_change(self, new_port: str, old_port: str, baud_rate: int, was_monitoring: bool):
        """Complete the COM port change after disconnect delay"""
        try:
            # Reconnect with new port
            self.connect()
            
            # Restore monitoring if it was active
            if was_monitoring:
                try:
                    self._start_monitoring()
                except AttributeError:
                    pass  # Monitoring not implemented yet
            
            # Note: Success/failure will be reported by _on_connection_state_changed
            
        except Exception as reconnect_error:
            self.formatter.append_status(
                self.terminal,
                f"Failed to connect to {new_port}: {str(reconnect_error)}",
                "error"
            )
            
            # Attempt to restore original connection
            self.config.port = old_port
            try:
                self.connect()
                if was_monitoring:
                    try:
                        self._start_monitoring()
                    except AttributeError:
                        pass  # Monitoring not implemented yet
                # Note: Connection status will be reported by _on_connection_state_changed
            except Exception as restore_error:
                self.formatter.append_status(
                    self.terminal,
                    f"Failed to restore original connection: {str(restore_error)}",
                    "error"
                )
    
    def _handle_encoding_error(self):
        """Handle encoding errors with intelligent baud rate detection"""
        import time
        
        self.encoding_error_count += 1
        self.data_packet_count += 1
        
        # Calculate error rate
        if self.data_packet_count >= self.encoding_error_window:
            error_rate = self.encoding_error_count / self.data_packet_count
            current_time = time.time()
            
            # Check if we should show a warning
            if (error_rate >= self.encoding_error_threshold and 
                current_time - self.last_encoding_warning >= self.encoding_warning_interval):
                
                self.last_encoding_warning = current_time
                
                # Show baud rate suggestion if not already shown
                if not self.baud_rate_suggestion_shown:
                    self._show_baud_rate_suggestion(error_rate)
                else:
                    # Just show a brief warning
                    self.formatter.append_status(
                        self.terminal,
                        f"High encoding error rate: {error_rate:.1%} - Check baud rate setting",
                        "warning"
                    )
            
            # Reset counters after window
            if self.data_packet_count >= self.encoding_error_window * 2:
                self.encoding_error_count = max(0, self.encoding_error_count // 2)
                self.data_packet_count = self.encoding_error_window
    
    def _show_baud_rate_suggestion(self, error_rate: float):
        """Show baud rate suggestion based on error patterns"""
        self.baud_rate_suggestion_shown = True
        
        current_baud = self.config.baudrate
        
        # Find current baud rate index
        current_index = -1
        for i, rate in enumerate(self.suggested_baud_rates):
            if rate == current_baud:
                current_index = i
                break
        
        # Suggest common alternatives based on typical usage
        suggestions = []
        
        # Add adjacent rates
        if current_index > 0:
            suggestions.append(self.suggested_baud_rates[current_index - 1])
        if current_index < len(self.suggested_baud_rates) - 1:
            suggestions.append(self.suggested_baud_rates[current_index + 1])
        
        # Add most common fallbacks
        common_rates = [9600, 115200, 38400]  # Most common rates
        for rate in common_rates:
            if rate != current_baud and rate not in suggestions:
                suggestions.append(rate)
        
        suggestions_str = ", ".join(map(str, suggestions[:4]))
        
        self.formatter.append_separator(self.terminal, "Baud Rate Issue Detected")
        self.formatter.append_status(
            self.terminal,
            f"High encoding error rate: {error_rate:.1%}",
            "warning"
        )
        self.formatter.append_status(
            self.terminal,
            f"Current baud rate: {current_baud} baud",
            "status"
        )
        self.formatter.append_status(
            self.terminal,
            f"Try these alternatives: {suggestions_str}",
            "status"
        )
        self.formatter.append_status(
            self.terminal,
            "Right-click → 'Reset Baud Rate Detection' to clear this warning",
            "status"
        )
        self.formatter.append_separator(self.terminal)
    
    def reset_baud_rate_detection(self):
        """Reset baud rate detection counters (call when baud rate changes)"""
        self.encoding_error_count = 0
        self.data_packet_count = 0
        self.baud_rate_suggestion_shown = False
        self.last_encoding_warning = 0
        self.consecutive_errors = 0
        
        self.formatter.append_status(
            self.terminal,
            "Baud rate detection reset",
            "status"
        )
    
    def _handle_excessive_errors(self):
        """Handle excessive consecutive encoding errors"""
        if not self.baud_rate_suggestion_shown:
            self.formatter.append_separator(self.terminal, "Connection Issue Detected")
            self.formatter.append_status(
                self.terminal,
                f"Too many consecutive encoding errors ({self.consecutive_errors})",
                "error"
            )
            self.formatter.append_status(
                self.terminal,
                f"Current baud rate: {self.config.baudrate}",
                "status"
            )
            self.formatter.append_status(
                self.terminal,
                "This usually indicates an incorrect baud rate setting",
                "warning"
            )
            self.formatter.append_status(
                self.terminal,
                "Try disconnecting and reconnecting with a different baud rate",
                "status"
            )
            self.formatter.append_separator(self.terminal)
            self.baud_rate_suggestion_shown = True
        
        # Reset consecutive errors to allow some data through
        self.consecutive_errors = 0
    
    def _is_data_garbled(self, text_data: str) -> bool:
        """Check if decoded data appears to be garbled"""
        if not text_data:
            return True
        
        # Check for high ratio of replacement characters
        replacement_ratio = text_data.count('\ufffd') / len(text_data)
        if replacement_ratio > 0.5:  # More than 50% replacement characters
            return True
        
        # Check for excessive non-printable characters
        printable_count = sum(1 for c in text_data if c.isprintable() or c in '\r\n\t')
        if len(text_data) > 0 and printable_count / len(text_data) < 0.3:  # Less than 30% printable
            return True
        
        return False

# ===== WELCOME CONFIG WIDGET =====
class WelcomeConfigWidget(QWidget):
    """Responsive welcome screen with embedded port configuration"""
    
    connectionRequested = pyqtSignal(object)  # SerialConfig
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.advanced_visible = False
  # Store enhanced port information
        self._setup_ui()
        self._populate_ports()
        
    def _setup_ui(self):
        """Setup minimal UI components"""
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setSpacing(0)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        
        # Add vertical stretch to center content
        self.main_layout.addStretch()
        
        # Create centered container for port configuration
        self.center_container = QWidget()
        self.center_container.setMaximumWidth(300)  # Limit width for better appearance
        self.center_container.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        
        # Port configuration section
        self._create_port_config_section()
        
        # Center the container horizontally
        self.h_layout = QHBoxLayout()
        self.h_layout.addStretch()
        self.h_layout.addWidget(self.center_container)
        self.h_layout.addStretch()
        
        self.main_layout.addLayout(self.h_layout)
        
        # Add vertical stretch to center content
        self.main_layout.addStretch()
        
    def _create_port_config_section(self):
        """Create the port configuration section"""
        # Simple form layout without borders
        self.form_layout = QFormLayout(self.center_container)
        self.form_layout.setSpacing(8)
        self.form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.form_layout.setContentsMargins(0, 0, 0, 0)
        self.form_layout.setHorizontalSpacing(12)
        
        # Port selection
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(180)
        self.form_layout.addRow("Port:", self.port_combo)
        
        # Baud rate
        self.baud_combo = QComboBox()
        self.baud_combo.setEditable(True)
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"])
        self.baud_combo.setCurrentText("115200")
        self.form_layout.addRow("Baud Rate:", self.baud_combo)
        
        # Create a horizontal layout for the connect button positioning
        connect_layout = QHBoxLayout()
        connect_layout.addStretch()  # Push button to the right

        # Connect button using RibbonButton class for consistency
        from ui.components import RibbonButton
        self.connect_btn = RibbonButton("Start", "enable")
        self.connect_btn.setToolTip("Start new serial terminal session")
        self.connect_btn.clicked.connect(self._handle_connect)
        
        connect_layout.addWidget(self.connect_btn)
        
        # Create a widget to hold the layout
        connect_widget = QWidget()
        connect_widget.setLayout(connect_layout)
        
        self.form_layout.addRow(connect_widget)
        
        
    def _populate_ports(self):
        """Populate available serial ports using enhanced registry scanning"""
        # Clear current ports
        self.port_combo.clear()
        self.port_combo.addItem("Scanning ports...")
        self.connect_btn.setEnabled(False)
        
        # Use direct port scanning for fast terminal loading
        self._scan_ports_direct()
    
    def _scan_ports_direct(self):
        """Direct port scanning for fast terminal loading"""
        self._fallback_to_basic_scan()
    
    def _fallback_to_basic_scan(self):
        """Direct port detection using existing PortScanner registry logic"""
        self.port_combo.clear()
        
        try:
            # Use existing PortScanner class for synchronous registry scanning
            scanner = PortScanner()
            ports = scanner.scan_registry_ports()
            
            if ports:
                for port in ports:
                    display_name = self._create_port_display_text(port)
                    self.port_combo.addItem(display_name, port.port_name)
                self.connect_btn.setEnabled(True)
            else:
                # Fallback to basic serial.tools if registry scan fails
                self._basic_serial_scan()
        except Exception as e:
            print(f"Error scanning registry ports: {e}")
            # Fallback to basic serial.tools scanning
            self._basic_serial_scan()
    
    def _basic_serial_scan(self):
        """Fallback to basic serial.tools.list_ports when registry scanning fails"""
        try:
            ports = serial.tools.list_ports.comports()
            
            if ports:
                for port in ports:
                    display_name = f"{port.device}"
                    if port.description and port.description != "n/a":
                        display_name += f" - {port.description}"
                    self.port_combo.addItem(display_name, port.device)
                self.connect_btn.setEnabled(True)
            else:
                self.port_combo.addItem("No ports available")
                self.connect_btn.setEnabled(False)
        except Exception as e:
            print(f"Error with basic port scan: {e}")
            self.port_combo.addItem("Port scanning failed")
            self.connect_btn.setEnabled(False)
    
    def _create_port_display_text(self, port):
        """Create display text for SerialPortInfo objects with type indicators"""
        display_text = port.port_name
        
        if port.is_moxa:
            display_text += "  •  Moxa"
            if port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        elif port.port_type.startswith("Virtual"):
            virtual_type = port.port_type.split(' ')[1] if ' ' in port.port_type else "Virtual"
            display_text += f"  •  {virtual_type} Port"
        else:
            display_text += "  •  Hardware Port"
            if port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        
        return display_text
    
            
    def _handle_connect(self):
        """Handle connect button click"""
        config = SerialConfig(
            port=self.port_combo.currentData() or self.port_combo.currentText(),
            baudrate=int(self.baud_combo.currentText()),
            databits=8,  # Default value
            parity="N",  # Default value
            stopbits=1.0  # Default value
        )
        
        self.connectionRequested.emit(config)
    

# ===== SPLIT CONTAINER =====
class SplitContainer(QWidget):
    """Manages split pane layout with recursive splitting"""
    
    activePaneChanged = pyqtSignal(TerminalPane)
    
    def __init__(self, initial_config: SerialConfig, parent=None, main_window=None):
        super().__init__(parent)
        self.panes: List[TerminalPane] = []
        self.active_pane: Optional[TerminalPane] = None
        self.main_window = main_window
        self._setup_ui(initial_config)
        
    def _setup_ui(self, config: SerialConfig):
        """Setup initial UI"""
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create initial pane
        initial_pane = self._create_pane(config)
        self.main_layout.addWidget(initial_pane)
        self._set_active_pane(initial_pane)
        
    def _create_pane(self, config: SerialConfig) -> TerminalPane:
        """Create a new terminal pane"""
        pane = TerminalPane(config, main_window=self.main_window)
        pane.splitRequested.connect(self._split_pane)
        pane.closeRequested.connect(self._close_pane)
        pane.focusChanged.connect(lambda focused: self._on_pane_focus(pane, focused))
        
        self.panes.append(pane)
        return pane
        
            
    def _create_welcome_pane(self):
        """Create a pane with welcome configuration widget"""
        # Create a custom pane that contains a welcome widget
        welcome_pane = QWidget()
        layout = QVBoxLayout(welcome_pane)
        layout.setContentsMargins(0, 0, 0, 0)
        
        welcome_widget = WelcomeConfigWidget()
        welcome_widget.connectionRequested.connect(
            lambda config: self._replace_welcome_with_terminal(welcome_pane, config)
        )
        
        layout.addWidget(welcome_widget)
        
        # Add to panes list (though it's not a TerminalPane)
        # We'll handle this specially in the container
        return welcome_pane
        
    def _replace_welcome_with_terminal(self, welcome_pane, config: SerialConfig):
        """Replace welcome pane with actual terminal pane"""
        # Find the parent splitter
        parent = welcome_pane.parent()
        if isinstance(parent, QSplitter):
            index = parent.indexOf(welcome_pane)
            
            # Create new terminal pane
            terminal_pane = self._create_pane(config)
            
            # Replace welcome pane with terminal pane
            parent.replaceWidget(index, terminal_pane)
            welcome_pane.deleteLater()
            
            # Connect and focus new pane
            terminal_pane.connect()
            terminal_pane.terminal.setFocus()
            self._set_active_pane(terminal_pane)
        
    def _split_pane(self, source_pane: TerminalPane, direction: str):
        """Split a pane horizontally or vertically"""
        # Always create new pane with welcome widget for new connection
        new_pane = self._create_welcome_pane()
        
        # Find the parent widget of the source pane
        parent_widget = source_pane.parent()
        
        # If parent is a splitter, we need to handle it differently
        if isinstance(parent_widget, QSplitter):
            # Get index of source pane in splitter
            index = parent_widget.indexOf(source_pane)
            
            # Create new splitter with opposite orientation
            new_orientation = Qt.Orientation.Horizontal if direction == 'vertical' else Qt.Orientation.Vertical
            
            # If splitter already has the same orientation, just add the new pane
            if parent_widget.orientation() == new_orientation:
                parent_widget.insertWidget(index + 1, new_pane)
            else:
                # Create nested splitter
                nested_splitter = self._create_splitter(new_orientation)
                parent_widget.replaceWidget(index, nested_splitter)
                nested_splitter.addWidget(source_pane)
                nested_splitter.addWidget(new_pane)
                
                # Set equal sizes
                nested_splitter.setSizes([500, 500])
        else:
            # Parent is the main layout, create new splitter
            orientation = Qt.Orientation.Horizontal if direction == 'vertical' else Qt.Orientation.Vertical
            splitter = self._create_splitter(orientation)
            
            # Replace source pane with splitter in layout
            self.main_layout.replaceWidget(source_pane, splitter)
            
            # Add panes to splitter
            splitter.addWidget(source_pane)
            splitter.addWidget(new_pane)
            
            # Set equal sizes
            splitter.setSizes([500, 500])
        
        # Focus the new welcome pane
        new_pane.setFocus()
    
    def _create_splitter(self, orientation: Qt.Orientation) -> QSplitter:
        """Create a styled splitter"""
        splitter = QSplitter(orientation)
        splitter.setHandleWidth(4)
        return splitter
    
    def cleanup(self):
        """Single point of cleanup for split container"""
        # Clean up all panes in the container
        for pane in self.panes:
            try:
                pane.cleanup()
            except Exception as e:
                print(f"Error cleaning up pane: {e}")
        
        # Clear the panes list
        self.panes.clear()
        self.active_pane = None
        
    def _close_pane(self, pane: TerminalPane):
        """Close a pane and reorganize layout"""
        if len(self.panes) == 1:
            # Can't close last pane
            return
            
        # Disconnect serial
        pane.disconnect()
        
        # Remove from list
        self.panes.remove(pane)
        
        # Find parent and remove
        parent = pane.parent()
        
        if isinstance(parent, QSplitter):
            # Remove from splitter
            pane.setParent(None)
            pane.deleteLater()
            
            # If splitter only has one widget left, replace splitter with that widget
            if parent.count() == 1:
                remaining_widget = parent.widget(0)
                grandparent = parent.parent()
                
                if isinstance(grandparent, QSplitter):
                    index = grandparent.indexOf(parent)
                    grandparent.replaceWidget(index, remaining_widget)
                else:
                    # It's in the main layout
                    self.main_layout.replaceWidget(parent, remaining_widget)
                
                parent.deleteLater()
        else:
            # Direct child of main layout
            pane.setParent(None)
            pane.deleteLater()
        
        # Set new active pane
        if self.panes:
            self._set_active_pane(self.panes[0])
            self.panes[0].terminal.setFocus()
            
    def _on_pane_focus(self, pane: TerminalPane, focused: bool):
        """Handle pane focus changes"""
        if focused and pane in self.panes:
            self._set_active_pane(pane)
            
    def _set_active_pane(self, pane: TerminalPane):
        """Set the active pane"""
        self.active_pane = pane
        self.activePaneChanged.emit(pane)
        
    def navigate_panes(self, direction: str):
        """Navigate between panes using keyboard"""
        if not self.active_pane or len(self.panes) < 2:
            return
            
        current_index = self.panes.index(self.active_pane)
        
        if direction in ['left', 'up']:
            new_index = (current_index - 1) % len(self.panes)
        else:  # right, down
            new_index = (current_index + 1) % len(self.panes)
            
        self.panes[new_index].terminal.setFocus()

# ===== CONNECTION DIALOG =====
class QuickConnectDialog(QDialog):
    """Minimal connection dialog matching Windows 10 style"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("New Connection")
        self.setModal(True)
        self.setFixedSize(400, 300)
  # Store enhanced port information
        self._setup_ui()
        self._populate_ports()
        
    def _setup_ui(self):
        """Setup dialog UI - trust Fusion theme"""
        # No custom styling needed
        
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)
        
        # Form layout
        form = QFormLayout()
        form.setSpacing(12)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        # Port selection
        self.port_combo = QComboBox()
        form.addRow("Port:", self.port_combo)
        
        # Baud rate
        self.baud_combo = QComboBox()
        self.baud_combo.setEditable(True)
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"])
        self.baud_combo.setCurrentText("115200")
        form.addRow("Baud Rate:", self.baud_combo)
        
        # Data bits
        self.databits_combo = QComboBox()
        self.databits_combo.addItems(["5", "6", "7", "8"])
        self.databits_combo.setCurrentText("8")
        form.addRow("Data Bits:", self.databits_combo)
        
        # Parity
        self.parity_combo = QComboBox()
        self.parity_combo.addItems(["None", "Even", "Odd", "Mark", "Space"])
        self.parity_combo.setCurrentText("None")
        form.addRow("Parity:", self.parity_combo)
        
        # Stop bits
        self.stopbits_combo = QComboBox()
        self.stopbits_combo.addItems(["1", "1.5", "2"])
        self.stopbits_combo.setCurrentText("1")
        form.addRow("Stop Bits:", self.stopbits_combo)
        
        layout.addLayout(form)
        layout.addStretch()
        
        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.refresh_btn = QPushButton("Refresh Ports")
        self.refresh_btn.clicked.connect(self._refresh_ports)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        self.connect_btn = QPushButton("Connect")
        self.connect_btn.clicked.connect(self.accept)
        
        button_layout.addWidget(self.refresh_btn)
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(self.connect_btn)
        layout.addLayout(button_layout)
        
    def _populate_ports(self):
        """Populate available serial ports using enhanced registry scanning"""
        # Clear current ports
        self.port_combo.clear()
        self.port_combo.addItem("Scanning ports...")
        self.connect_btn.setEnabled(False)
        
        # Use direct port scanning for fast terminal loading
        self._scan_ports_direct()
    
    def _scan_ports_direct(self):
        """Direct port scanning for fast terminal loading"""
        self._fallback_to_basic_scan()
    
    def _fallback_to_basic_scan(self):
        """Direct port detection using existing PortScanner registry logic"""
        self.port_combo.clear()
        
        try:
            # Use existing PortScanner class for synchronous registry scanning
            scanner = PortScanner()
            ports = scanner.scan_registry_ports()
            
            if ports:
                for port in ports:
                    display_name = self._create_port_display_text(port)
                    self.port_combo.addItem(display_name, port.port_name)
                self.connect_btn.setEnabled(True)
            else:
                # Fallback to basic serial.tools if registry scan fails
                self._basic_serial_scan()
        except Exception as e:
            print(f"Error scanning registry ports: {e}")
            # Fallback to basic serial.tools scanning
            self._basic_serial_scan()
    
    def _basic_serial_scan(self):
        """Fallback to basic serial.tools.list_ports when registry scanning fails"""
        try:
            ports = serial.tools.list_ports.comports()
            
            if ports:
                for port in ports:
                    display_name = f"{port.device}"
                    if port.description and port.description != "n/a":
                        display_name += f" - {port.description}"
                    self.port_combo.addItem(display_name, port.device)
                self.connect_btn.setEnabled(True)
            else:
                self.port_combo.addItem("No ports available")
                self.connect_btn.setEnabled(False)
        except Exception as e:
            print(f"Error with basic port scan: {e}")
            self.port_combo.addItem("Port scanning failed")
            self.connect_btn.setEnabled(False)
    
    def _create_port_display_text(self, port):
        """Create display text for SerialPortInfo objects with type indicators"""
        display_text = port.port_name
        
        if port.is_moxa:
            display_text += "  •  Moxa"
            if port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        elif port.port_type.startswith("Virtual"):
            virtual_type = port.port_type.split(' ')[1] if ' ' in port.port_type else "Virtual"
            display_text += f"  •  {virtual_type} Port"
        else:
            display_text += "  •  Hardware Port"
            if port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        
        return display_text
    
    def _on_scan_progress(self, message):
        """Handle scan progress updates"""
        # Update the first item to show progress
        if self.port_combo.count() > 0:
            self.port_combo.setItemText(0, message)
    
    def _on_enhanced_ports_scanned(self, ports):
        """Handle enhanced port scan results"""
        self.scanned_ports = ports
        self.port_combo.clear()
        
        if not ports:
            self.port_combo.addItem("No ports available")
            self.connect_btn.setEnabled(False)
            return
        
        # Populate with enhanced port information
        for port in ports:
            display_text = self._create_enhanced_port_display_text(port)
            self.port_combo.addItem(display_text, port.port_name)
        
        self.connect_btn.setEnabled(True)
    
    def _on_scan_finished(self):
        """Handle scan completion"""
        self.port_scanner = None
    
    def _create_enhanced_port_display_text(self, port):
        """Create enhanced display text for port with type indicators"""
        display_text = port.port_name
        
        if port.is_moxa:
            display_text += "  •  Moxa Device"
            if port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        elif port.port_type.startswith("Virtual"):
            virtual_type = port.port_type.split(' ')[1] if ' ' in port.port_type else "Virtual"
            display_text += f"  •  {virtual_type} Port"
        else:
            display_text += "  •  Hardware Port"
            if port.device_name and port.device_name != "Unknown":
                display_text += f"  •  {port.device_name}"
        
        return display_text
            
    def get_config(self) -> SerialConfig:
        """Get the serial configuration from dialog"""
        parity_map = {"None": "N", "Even": "E", "Odd": "O", "Mark": "M", "Space": "S"}
        
        return SerialConfig(
            port=self.port_combo.currentData() or self.port_combo.currentText(),
            baudrate=int(self.baud_combo.currentText()),
            databits=int(self.databits_combo.currentText()),
            parity=parity_map[self.parity_combo.currentText()],
            stopbits=float(self.stopbits_combo.currentText())
        )
    
    def _check_for_moxa_port(self):
        """Check if selected port is a MOXA device and provide recommendations"""
        current_port_name = self.port_combo.currentData() or self.port_combo.currentText()
        
        # Find the port info from scanned ports
        selected_port = None
        for port in self.scanned_ports:
            if port.port_name == current_port_name:
                selected_port = port
                break
        
        if selected_port and selected_port.is_moxa:
            # Show MOXA recommendations dialog
            self._show_moxa_recommendations(selected_port)
    
    def _show_moxa_recommendations(self, port_info):
        """Show MOXA port recommendations dialog"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("MOXA Device Detected")
        msg_box.setIcon(QMessageBox.Icon.Information)
        
        # Build recommendation text
        recommendations_text = f"MOXA virtual port detected: {port_info.port_name}\n\n"
        recommendations_text += "Recommendations for MOXA devices:\n"
        
        if hasattr(port_info, 'moxa_details') and port_info.moxa_details:
            for recommendation in port_info.moxa_details.get('recommendations', []):
                recommendations_text += f"• {recommendation}\n"
        else:
            recommendations_text += "• Disable CTS handshaking for network serial servers\n"
            recommendations_text += "• Check network connectivity to MOXA device\n"
            recommendations_text += "• Verify MOXA driver configuration\n"
            recommendations_text += "• Consider matching baud rate to source device\n"
        
        recommendations_text += "\nThese settings may help improve connection reliability."
        
        msg_box.setText(recommendations_text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()
    
    def _refresh_ports(self):
        """Refresh the ports list manually"""
        self._populate_ports()
    
    def cleanup(self):
        """Clean up any running port scanner threads"""
        if self.port_scanner and self.port_scanner.isRunning():
            self.port_scanner.wait(1000)  # Wait up to 1 second for completion


# ===== MAIN WINDOW =====
class SerialMonitorWindow(QMainWindow):
    """Main window with tab management"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serial Terminal")
        self.setMinimumSize(800, 600)
        
        # Set custom window icon using terminal settings icon
        icon_pixmap = QPixmap(64, 64)
        icon_pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(icon_pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        svg_renderer = QSvgRenderer()
        svg_renderer.load(Icons.terminal_settings(self.palette()).encode())
        svg_renderer.render(painter)
        painter.end()
        self.setWindowIcon(QIcon(icon_pixmap))
        
        self.tabs: Dict[QWidget, SplitContainer] = {}
        self.close_button_icon = None  # Store close button icon
        self._setup_ui()
        self._setup_shortcuts()
        self._apply_window_style()
        
        # Show connection dialog after window is shown
        QTimer.singleShot(200, self._show_initial_connection_dialog)
        
    def _setup_ui(self):
        """Setup main window UI"""
        # Create ribbon toolbar
        self.ribbon = RibbonToolbar()
        self.addToolBar(self.ribbon)

        # Connect ribbon signals
        self._connect_ribbon_signals()

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self._close_tab)
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

        # Set up custom close button icon - simple approach
        self._setup_close_button_icon()

        main_layout.addWidget(self.tab_widget)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        # Status update timer
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self._update_status_bar)
        self.status_timer.start(1000)  # Update every second

    def _connect_ribbon_signals(self):
        """Connect ribbon toolbar signals to existing methods."""
        self.ribbon.new_connection.connect(self._new_connection)
        self.ribbon.refresh_ports.connect(self._refresh_ports)
        self.ribbon.toggle_connection.connect(self._toggle_connection)
        self.ribbon.clear_terminal.connect(self._clear_current_terminal)
        self.ribbon.show_settings.connect(self._show_settings_menu)

    def checkbox_icon(self, checked: bool) -> QIcon:
        """Generate checkbox icon using palette colors like VirtualPortManager"""
        palette = self.palette()
        border_color = palette.color(QPalette.ColorRole.Mid).name()
        bg_color = palette.color(QPalette.ColorRole.Base).name()
        check_color = palette.color(QPalette.ColorRole.Highlight).name()

        if checked:
            svg = f'''<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
                <rect x="0.5" y="0.5" width="15" height="15" fill="{border_color}" stroke="{border_color}" stroke-width="1"/>
                <rect x="2" y="2" width="12" height="12" fill="{bg_color}"/>
                <path d="M4 8l2 2 6-6" stroke="{check_color}" stroke-width="1.5" fill="none" stroke-linecap="round"/>
            </svg>'''
        else:
            svg = f'''<svg width="16" height="16" xmlns="http://www.w3.org/2000/svg">
                <rect x="0.5" y="0.5" width="15" height="15" fill="{border_color}" stroke="{border_color}" stroke-width="1"/>
            </svg>'''
        
        renderer = QSvgRenderer(QByteArray(svg.encode()))
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return QIcon(pixmap)
    
    def _setup_close_button_icon(self):
        """Set up custom close button - done in _apply_window_style"""
        pass
        
    def _apply_window_style(self):
        """Apply minimal window style - match background colors for clean look"""
        # Get the darker window background color from palette
        palette = self.palette()
        window_bg = palette.color(QPalette.ColorRole.Window)
        highlight_color = palette.color(QPalette.ColorRole.Highlight)

        # Apply window background to tab widget for seamless look
        tab_palette = self.tab_widget.palette()
        tab_palette.setColor(QPalette.ColorRole.Base, window_bg)
        self.tab_widget.setPalette(tab_palette)

        # Create custom white X close button icon
        close_icon_svg = """<svg width="16" height="16" viewBox="0 0 16 16" xmlns="http://www.w3.org/2000/svg">
            <line x1="4" y1="4" x2="12" y2="12" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/>
            <line x1="12" y1="4" x2="4" y2="12" stroke="#ffffff" stroke-width="2" stroke-linecap="round"/>
        </svg>"""

        renderer = QSvgRenderer(QByteArray(close_icon_svg.encode()))
        pixmap = QPixmap(16, 16)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()

        # Store the close button icon for reuse
        self.close_button_icon = QIcon(pixmap)

        # Apply to existing tabs
        self._apply_close_icon_to_tabs()

        # Style the tab widget and tab bar for ultra-minimal clean look with white text
        self.tab_widget.setStyleSheet(f"""
            QTabWidget::pane {{
                background-color: {window_bg.name()};
                border: none;
            }}
            QTabBar::tab {{
                background-color: {window_bg.name()};
                color: #ffffff;
                border: none;
                padding: 8px 16px;
                min-width: 100px;
            }}
            QTabBar::tab:selected {{
                background-color: {window_bg.name()};
                color: #ffffff;
                border-bottom: 2px solid {highlight_color.name()};
            }}
            QTabBar::tab:hover {{
                background-color: {window_bg.lighter(110).name()};
            }}
            QTabBar::close-button {{
                background-color: transparent;
                border: none;
                border-radius: 2px;
                margin: 2px;
                padding: 2px;
            }}
            QTabBar::close-button:hover {{
                background-color: rgba(255, 255, 255, 0.1);
            }}
        """)
        
    def _apply_close_icon_to_tabs(self):
        """Apply custom close button icon to all tabs"""
        if not self.close_button_icon:
            return

        for i in range(self.tab_widget.count()):
            button = self.tab_widget.tabBar().tabButton(i, QTabBar.ButtonPosition.RightSide)
            if button:
                button.setIcon(self.close_button_icon)
                button.setIconSize(QSize(12, 12))
                button.setStyleSheet("border: none; background: transparent;")

    def _setup_shortcuts(self):
        """Setup global keyboard shortcuts"""
        # Tab management
        QShortcut(QKeySequence("Ctrl+N"), self, self._new_connection)
        QShortcut(QKeySequence("Ctrl+W"), self, lambda: self._close_tab(self.tab_widget.currentIndex()))
        QShortcut(QKeySequence("Ctrl+Tab"), self, self._next_tab)
        QShortcut(QKeySequence("Ctrl+Shift+Tab"), self, self._prev_tab)
        
        # Pane navigation
        QShortcut(QKeySequence("Alt+Left"), self, lambda: self._navigate_panes("left"))
        QShortcut(QKeySequence("Alt+Right"), self, lambda: self._navigate_panes("right"))
        QShortcut(QKeySequence("Alt+Up"), self, lambda: self._navigate_panes("up"))
        QShortcut(QKeySequence("Alt+Down"), self, lambda: self._navigate_panes("down"))
        
        # Split shortcuts
        QShortcut(QKeySequence("Alt+Shift+-"), self, lambda: self._split_current_pane("vertical"))
        QShortcut(QKeySequence("Alt+Shift++"), self, lambda: self._split_current_pane("horizontal"))
        
        # Close pane
        QShortcut(QKeySequence("Ctrl+Shift+W"), self, self._close_current_pane)
        
    def _show_initial_connection_dialog(self):
        """Show welcome tab when window opens with no tabs"""
        if self.tab_widget.count() == 0:
            self._show_welcome_tab()
    
    def _refresh_ports(self):
        """Refresh ports in all welcome widgets"""
        for i in range(self.tab_widget.count()):
            widget = self.tab_widget.widget(i)
            if isinstance(widget, WelcomeConfigWidget):
                widget._populate_ports()
    
    def _show_welcome_tab(self):
        """Show welcome tab with responsive embedded port configuration"""
        # Check if welcome tab already exists
        if self._has_welcome_tab():
            return
            
        try:
            welcome_widget = WelcomeConfigWidget()
            welcome_widget.connectionRequested.connect(self._handle_welcome_connection)
            
            index = self.tab_widget.addTab(welcome_widget, "New tab")
            self.tab_widget.setCurrentIndex(index)
            self._apply_close_icon_to_tabs()  # Apply custom close icon
        except Exception as e:
            print(f"Error creating welcome tab: {e}")
    
    def _has_welcome_tab(self) -> bool:
        """Check if a welcome tab already exists"""
        for i in range(self.tab_widget.count()):
            if self.tab_widget.tabText(i) == "New tab":
                return True
        return False
    
    def _handle_welcome_connection(self, config: SerialConfig):
        """Handle connection request from welcome widget"""
        try:
            self._create_tab(config)
            
            # Remove welcome tab after successful connection
            self._remove_welcome_tab()
        except Exception as e:
            print(f"Error handling welcome connection: {e}")
    
    def _remove_welcome_tab(self):
        """Safely remove welcome tab"""
        try:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == "New tab":
                    widget = self.tab_widget.widget(i)
                    self.tab_widget.removeTab(i)
                    if widget:
                        widget.deleteLater()
                    break
        except Exception as e:
            print(f"Error removing welcome tab: {e}")
    
    def _new_connection(self):
        """Create new tab with welcome screen"""
        self._show_welcome_tab()
            
    def _create_tab(self, config: SerialConfig):
        """Create a new tab with split container"""
        # Create container
        container = SplitContainer(config, main_window=self)
        container.activePaneChanged.connect(self._on_active_pane_changed)
        
        # Store references
        self.tabs[container] = container
        
        # Add tab directly with the container
        index = self.tab_widget.addTab(container, config.port)
        self.tab_widget.setCurrentIndex(index)
        self._apply_close_icon_to_tabs()  # Apply custom close icon
        
        # Auto-connect the first pane
        if container.active_pane:
            container.active_pane.connect()
            
    def cleanup(self):
        """Cleanup all tabs"""
        for container in self.tabs.values():
            container.cleanup()
    
    def get_all_active_panes(self):
        """Get all terminal panes across all containers"""
        all_panes = []
        for container in self.tabs.values():
            all_panes.extend(container.panes)
        return all_panes
    
    def get_connected_ports(self):
        """Get set of all currently connected ports"""
        connected_ports = set()
        for pane in self.get_all_active_panes():
            if pane.is_connected:
                connected_ports.add(pane.config.port)
        return connected_ports
    
    def _close_tab(self, index: int):
        """Close a tab and cleanup"""
        if index < 0 or index >= self.tab_widget.count():
            return
            
        widget = self.tab_widget.widget(index)
        if not widget:
            return
            
        # Don't allow closing the last welcome tab if it's the only one
        if (self.tab_widget.count() == 1 and 
            self.tab_widget.tabText(index) == "New tab"):
            return
            
        # Check for active connections and cleanup
        if widget in self.tabs:
            container = self.tabs[widget]
            
            # Cleanup all panes
            try:
                container.cleanup()
                del self.tabs[widget]
            except Exception as e:
                print(f"Error cleaning up container: {e}")
                
        # Remove tab and schedule widget deletion
        self.tab_widget.removeTab(index)
        
        # Schedule widget deletion and check for empty tabs
        if widget:
            widget.deleteLater()
            
        # Use QTimer to ensure the count is updated after widget deletion
        QTimer.singleShot(0, self._check_empty_tabs)
    
    def _check_empty_tabs(self):
        """Check if tabs are empty and show welcome tab if needed"""
        try:
            if self.tab_widget.count() == 0:
                self._show_welcome_tab()
        except Exception as e:
            print(f"Error checking empty tabs: {e}")
            
    def _next_tab(self):
        """Switch to next tab"""
        current = self.tab_widget.currentIndex()
        count = self.tab_widget.count()
        if count > 0:
            next_index = (current + 1) % count
            self.tab_widget.setCurrentIndex(next_index)
            
    def _prev_tab(self):
        """Switch to previous tab"""
        current = self.tab_widget.currentIndex()
        count = self.tab_widget.count()
        if count > 0:
            prev_index = (current - 1) % count
            self.tab_widget.setCurrentIndex(prev_index)
            
    def _on_tab_changed(self, index: int):
        """Handle tab change"""
        self._update_status_bar()
        self._update_ribbon_connection_state()
        
    def _get_current_container(self) -> Optional[SplitContainer]:
        """Get current tab's split container"""
        widget = self.tab_widget.currentWidget()
        return self.tabs.get(widget)
        
    def _navigate_panes(self, direction: str):
        """Navigate panes in current tab"""
        container = self._get_current_container()
        if container:
            container.navigate_panes(direction)
            
    def _split_current_pane(self, direction: str):
        """Split the current active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container._split_pane(container.active_pane, direction)
            
    def _close_current_pane(self):
        """Close the current active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container._close_pane(container.active_pane)

    def _clear_current_terminal(self):
        """Clear the current active terminal pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            container.active_pane._clear_terminal()

    def _toggle_connection(self):
        """Toggle connection state of active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            pane = container.active_pane
            if pane.is_connected:
                pane.disconnect()
            else:
                pane.connect()
            # Update ribbon button state
            self._update_ribbon_connection_state()

    def _update_ribbon_connection_state(self):
        """Update ribbon toolbar connection button based on active pane state"""
        container = self._get_current_container()
        if container and container.active_pane:
            is_connected = container.active_pane.is_connected
            self.ribbon.set_connection_state(is_connected)
        else:
            # No active pane, default to disconnected state
            self.ribbon.set_connection_state(False)

    def _show_help(self):
        """Show help dialog for Serial Terminal"""
        from PyQt6.QtWidgets import QMessageBox
        help_text = """<h2>Serial Terminal Help</h2>
        <p><b>Connection:</b></p>
        <ul>
        <li>Use 'New' button to create a new connection tab</li>
        <li>Select port and baud rate, then click play button to connect</li>
        <li>Use 'Refresh' to update available ports</li>
        </ul>
        <p><b>View Options:</b></p>
        <ul>
        <li><b>Split Vertical (⬌):</b> Split current pane side-by-side</li>
        <li><b>Split Horizontal (⬍):</b> Split current pane top-to-bottom</li>
        <li><b>Close:</b> Close the active pane</li>
        <li><b>Clear:</b> Clear terminal output</li>
        </ul>
        <p><b>Keyboard Shortcuts:</b></p>
        <ul>
        <li>Ctrl+N - New connection</li>
        <li>Ctrl+W - Close tab</li>
        <li>Alt+Shift+- - Split vertical</li>
        <li>Alt+Shift++ - Split horizontal</li>
        </ul>"""

        QMessageBox.information(self, "Serial Terminal Help", help_text)

    def _on_active_pane_changed(self, pane: TerminalPane):
        """Handle active pane change"""
        self._update_status_bar()
        self._update_ribbon_connection_state()
        
    def _update_status_bar(self):
        """Update status bar with active pane info"""
        container = self._get_current_container()
        if container and container.active_pane:
            status = container.active_pane.get_status_info()
            self.status_bar.showMessage(status)
        else:
            self.status_bar.showMessage("No active connection")
    
    def closeEvent(self, event):
        """Handle window close with simple cleanup"""
        try:
            # Stop the status timer
            self.status_timer.stop()
            
            # Signal all workers to stop gracefully
            for container in self.tabs.values():
                container.cleanup()
            
            # Accept the event and let atexit handle final cleanup
            event.accept()
        except Exception as e:
            print(f"Error during close: {e}")
            event.accept()

    def _show_settings_menu(self):
        """Show settings menu for the active pane"""
        container = self._get_current_container()
        if container and container.active_pane:
            # Get the active pane
            active_pane = container.active_pane

            # Calculate menu position (show below the settings button in ribbon)
            button_global_pos = self.ribbon.settings_button.mapToGlobal(self.ribbon.settings_button.rect().bottomLeft())

            # Create menu using the pane's method
            menu = active_pane._create_terminal_menu()

            # Show menu at calculated position
            menu.exec(button_global_pos)
        else:
            # Show a simple message if no active pane
            menu = QMenu(self)
            # Remove explicit styling to match primary GUI approach

            no_connection = menu.addAction("No active connection")
            no_connection.setEnabled(False)

            button_global_pos = self.ribbon.settings_button.mapToGlobal(self.ribbon.settings_button.rect().bottomLeft())
            menu.exec(button_global_pos)

# ===== MAIN ENTRY POINT =====
def main():
    """Main application entry point for testing"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Serial Monitor")
    app.setOrganizationName("SerialMonitor")
    
    # Apply Windows 10 style
    app.setStyle("Fusion")
    
    # Trust Fusion theme - no custom palette needed
    
    # Create and show main window
    window = SerialMonitorWindow()
    window.resize(1200, 800)
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()