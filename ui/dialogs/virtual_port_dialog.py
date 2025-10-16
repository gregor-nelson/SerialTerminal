#!/usr/bin/env python3
"""
Virtual Port Manager Dialog
Manages com0com virtual serial port pairs through a minimal, clean UI.
"""

import os
import re
import subprocess
import time
import shlex
import logging
import json
import sys
import tempfile
import ctypes
from ctypes import wintypes
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List
from enum import Enum

# Configure module logger
logger = logging.getLogger(__name__)

# ============================================================================
# Windows API Definitions for UAC Elevation
# ============================================================================

# Windows constants
SEE_MASK_NOCLOSEPROCESS = 0x00000040
SEE_MASK_NO_CONSOLE = 0x00008000
SW_HIDE = 0

# Define SHELLEXECUTEINFO structure
class SHELLEXECUTEINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('fMask', ctypes.c_ulong),
        ('hwnd', wintypes.HANDLE),
        ('lpVerb', wintypes.LPCWSTR),
        ('lpFile', wintypes.LPCWSTR),
        ('lpParameters', wintypes.LPCWSTR),
        ('lpDirectory', wintypes.LPCWSTR),
        ('nShow', ctypes.c_int),
        ('hInstApp', wintypes.HINSTANCE),
        ('lpIDList', ctypes.c_void_p),
        ('lpClass', wintypes.LPCWSTR),
        ('hkeyClass', wintypes.HKEY),
        ('dwHotKey', wintypes.DWORD),
        ('hIconOrMonitor', wintypes.HANDLE),
        ('hProcess', wintypes.HANDLE),
    ]

# Declare ShellExecuteEx function
ShellExecuteEx = ctypes.windll.shell32.ShellExecuteExW
ShellExecuteEx.argtypes = [ctypes.POINTER(SHELLEXECUTEINFO)]
ShellExecuteEx.restype = wintypes.BOOL

# Declare process wait functions
WaitForSingleObject = ctypes.windll.kernel32.WaitForSingleObject
WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
WaitForSingleObject.restype = wintypes.DWORD

GetExitCodeProcess = ctypes.windll.kernel32.GetExitCodeProcess
GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
GetExitCodeProcess.restype = wintypes.BOOL

CloseHandle = ctypes.windll.kernel32.CloseHandle
CloseHandle.argtypes = [wintypes.HANDLE]
CloseHandle.restype = wintypes.BOOL

# Constants for WaitForSingleObject
WAIT_OBJECT_0 = 0x00000000
WAIT_TIMEOUT = 0x00000102
INFINITE = 0xFFFFFFFF

# Explicit PyQt6 imports for better maintainability
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget,
    QPushButton, QLabel, QWidget, QTableWidgetItem,
    QHeaderView, QApplication
)
from PyQt6.QtCore import (
    QThread, pyqtSignal, QTimer, Qt, QSize
)
from PyQt6.QtGui import (
    QIcon, QPainter, QPixmap, QFont, QScreen
)
from PyQt6.QtSvg import QSvgRenderer


# ============================================================================
# Data Models (from VirtualPortManager)
# ============================================================================

class PortStatus(Enum):
    """Port pair status enumeration."""
    ACTIVE = "Active"
    DISABLED = "Disabled"
    ERROR = "Error"
    UNKNOWN = "Unknown"


@dataclass
class Port:
    """Individual virtual port model."""
    identifier: str  # e.g., "CNCA0", "CNCB0"
    port_name: str = ""  # e.g., "COM8", "COM9"
    parameters: Dict[str, Any] = field(default_factory=dict)


@dataclass
class PortPair:
    """Virtual port pair model."""
    number: int
    port_a: Port
    port_b: Port
    status: PortStatus = PortStatus.UNKNOWN

    def __post_init__(self):
        """Initialize port identifiers if not set."""
        if not self.port_a.identifier:
            self.port_a.identifier = f"CNCA{self.number}"
        if not self.port_b.identifier:
            self.port_b.identifier = f"CNCB{self.number}"


@dataclass
class CommandResult:
    """Result of a setupc.exe command execution."""
    success: bool
    output: str = ""
    error: str = ""
    return_code: int = 0
    execution_time: float = 0.0
    command: str = ""

    def get_error_message(self) -> str:
        """Get user-friendly error message."""
        if self.success:
            return ""
        if self.error:
            return self.error
        elif self.return_code != 0:
            return f"Command failed with exit code {self.return_code}"
        else:
            return "Unknown error occurred"


class PortListParser:
    """Parser for setupc.exe list command output."""

    @staticmethod
    def parse_port_list(output: str) -> List[PortPair]:
        """Parse setupc.exe list output into PortPair objects."""
        port_pairs = []

        for line in output.strip().split('\n'):
            line = line.strip()
            if not line or line.startswith('command>'):
                continue

            # Look for port identifiers (CNCA0, CNCB0, etc.)
            if line.startswith('CNC'):
                parts = line.split()
                if len(parts) >= 1:
                    port_id = parts[0]

                    # Validate port ID format before parsing
                    if len(port_id) < 5 or port_id[3] not in ('A', 'B'):
                        logger.warning(f"Skipping malformed port identifier: {port_id}")
                        continue  # Skip malformed port identifier

                    try:
                        # Extract pair number
                        pair_num = int(port_id[4:])  # Skip "CNCA" or "CNCB"
                        port_type = port_id[3]  # "A" or "B"
                    except (ValueError, IndexError) as e:
                        logger.warning(f"Failed to parse port identifier {port_id}: {e}")
                        continue  # Skip if pair number is invalid

                    # Find or create port pair
                    pair = next((p for p in port_pairs if p.number == pair_num), None)
                    if not pair:
                        pair = PortPair(
                            number=pair_num,
                            port_a=Port(identifier=f"CNCA{pair_num}"),
                            port_b=Port(identifier=f"CNCB{pair_num}"),
                            status=PortStatus.ACTIVE
                        )
                        port_pairs.append(pair)

                    # Set port data
                    port = pair.port_a if port_type == 'A' else pair.port_b
                    port.identifier = port_id

                    # Parse parameters from the rest of the line
                    if len(parts) > 1:
                        params_str = ' '.join(parts[1:])
                        port.parameters = PortListParser._parse_parameters(params_str)

                        # Extract port name if present
                        if 'PortName' in port.parameters:
                            port.port_name = port.parameters['PortName']

        return sorted(port_pairs, key=lambda p: p.number)

    @staticmethod
    def _parse_parameters(params_str: str) -> Dict[str, str]:
        """Parse parameter string into dictionary."""
        parameters = {}

        # Simple parameter parsing
        for param in params_str.split(','):
            param = param.strip()
            if '=' in param:
                key, value = param.split('=', 1)
                parameters[key.strip()] = value.strip()

        return parameters


# ============================================================================
# Worker Thread (from VirtualPortManager)
# ============================================================================

class ElevatedHelperWorker(QThread):
    """Worker thread for executing commands via elevated helper process."""

    command_finished = pyqtSignal(CommandResult)

    def __init__(self, setupc_path: str, command: str, timeout: int = 30, dev_mode_fallback: bool = True):
        super().__init__()
        self.setupc_path = setupc_path
        self.command = command
        self.timeout = timeout
        self.dev_mode_fallback = dev_mode_fallback
        self.result = None
        self._process = None

    def _get_helper_path(self) -> Optional[str]:
        """Get path to the elevated helper executable."""
        # Check if running as bundled app (PyInstaller)
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            app_dir = os.path.dirname(sys.executable)
            helper_path = os.path.join(app_dir, 'SerialPortManager.exe')
            return helper_path if os.path.exists(helper_path) else None
        else:
            # Running as script - for development
            # ONLY look for compiled .exe in dev mode
            # Python scripts cannot trigger UAC elevation
            script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            helper_path = os.path.join(script_dir, 'dist', 'SerialPortManager.exe')

            if os.path.exists(helper_path):
                return helper_path
            else:
                # No compiled helper - return None to trigger dev mode direct execution
                logger.debug("DEV MODE: Compiled helper not found, will use direct execution")
                return None

    def _is_dev_mode(self) -> bool:
        """Check if running in development mode (not compiled)."""
        return not getattr(sys, 'frozen', False)

    def _execute_direct_command(self) -> CommandResult:
        """
        Execute setupc command directly (dev mode fallback).
        This will likely fail with permission errors but allows testing the flow.
        """
        start_time = time.time()
        logger.warning("DEV MODE: Executing command directly without elevation")

        try:
            full_command = f'"{self.setupc_path}" {self.command}'
            command_args = shlex.split(full_command)

            result = subprocess.run(
                command_args,
                capture_output=True,
                text=True,
                timeout=self.timeout,
                shell=False,
                cwd=os.path.dirname(self.setupc_path)
            )

            execution_time = time.time() - start_time

            command_result = CommandResult(
                success=result.returncode == 0,
                output=result.stdout,
                error=result.stderr if result.stderr else (
                    "Command completed (dev mode - no elevation)" if result.returncode == 0
                    else "Command failed - may require elevation"
                ),
                return_code=result.returncode,
                execution_time=execution_time,
                command=self.command
            )

            if result.returncode == 0:
                logger.info(f"DEV MODE: Command succeeded in {execution_time:.2f}s")
            else:
                logger.error(f"DEV MODE: Command failed (code {result.returncode}) - may need elevation")

            return command_result

        except Exception as e:
            execution_time = time.time() - start_time
            logger.exception(f"DEV MODE: Command execution failed")
            return CommandResult(
                success=False,
                output="",
                error=f"Dev mode execution failed: {str(e)}",
                return_code=-3,
                execution_time=execution_time,
                command=self.command
            )

    def _execute_elevated_command(self, helper_path: str, output_file: str) -> tuple[int, float]:
        """
        Execute helper using ShellExecuteEx with UAC elevation.

        Args:
            helper_path: Path to the elevated helper executable
            output_file: Path to temporary file for JSON output

        Returns:
            Tuple of (return_code, execution_time)
        """
        start_time = time.time()

        # Build parameters string
        parameters = f'"{self.setupc_path}" {self.command} {self.timeout} --output-file "{output_file}"'

        logger.debug(f"Launching elevated helper with ShellExecuteEx")
        logger.debug(f"  File: {helper_path}")
        logger.debug(f"  Parameters: {parameters}")

        # Initialize SHELLEXECUTEINFO structure
        sei = SHELLEXECUTEINFO()
        sei.cbSize = ctypes.sizeof(sei)
        sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NO_CONSOLE
        sei.hwnd = None
        sei.lpVerb = "runas"  # Request elevation
        sei.lpFile = helper_path
        sei.lpParameters = parameters
        sei.lpDirectory = os.path.dirname(helper_path)
        sei.nShow = SW_HIDE
        sei.hProcess = None

        # Execute with elevation
        if not ShellExecuteEx(ctypes.byref(sei)):
            error_code = ctypes.get_last_error()
            if error_code == 1223:  # ERROR_CANCELLED - User denied UAC
                logger.warning("UAC prompt was denied by user")
                return 1223, time.time() - start_time
            else:
                logger.error(f"ShellExecuteEx failed with error code: {error_code}")
                return -1, time.time() - start_time

        # Store process handle for cleanup
        self._process = sei.hProcess

        if not sei.hProcess:
            logger.error("No process handle returned from ShellExecuteEx")
            return -1, time.time() - start_time

        # Wait for process completion with timeout (add extra buffer for UAC prompt)
        uac_timeout = self.timeout + 30  # Extra time for user to respond to UAC
        timeout_ms = uac_timeout * 1000

        logger.debug(f"Waiting for elevated process (timeout: {uac_timeout}s)")
        wait_result = WaitForSingleObject(sei.hProcess, timeout_ms)

        if wait_result == WAIT_TIMEOUT:
            logger.error(f"Helper process timed out after {uac_timeout}s")
            CloseHandle(sei.hProcess)
            return -1, time.time() - start_time

        # Get exit code
        exit_code = wintypes.DWORD()
        if not GetExitCodeProcess(sei.hProcess, ctypes.byref(exit_code)):
            logger.error("Failed to get process exit code")
            CloseHandle(sei.hProcess)
            return -1, time.time() - start_time

        CloseHandle(sei.hProcess)
        execution_time = time.time() - start_time

        logger.debug(f"Elevated process completed with exit code {exit_code.value} in {execution_time:.2f}s")
        return exit_code.value, execution_time

    def run(self):
        """Execute the command via elevated helper process."""
        start_time = time.time()
        logger.info(f"Executing elevated command: {self.command}")

        try:
            # Get helper path
            helper_path = self._get_helper_path()
            if not helper_path:
                # Dev mode fallback - execute directly without elevation
                if self._is_dev_mode() and self.dev_mode_fallback:
                    logger.warning("DEV MODE: Helper not found, falling back to direct execution")
                    self.result = self._execute_direct_command()
                    self.command_finished.emit(self.result)
                    return
                else:
                    raise FileNotFoundError("Elevated helper executable not found")

            # Create temporary file for output
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as temp_file:
                output_file = temp_file.name

            try:
                # Execute elevated command using ShellExecuteEx
                return_code, execution_time = self._execute_elevated_command(helper_path, output_file)

                # Read output from temp file
                stdout = ""
                if os.path.exists(output_file):
                    try:
                        with open(output_file, 'r', encoding='utf-8') as f:
                            stdout = f.read()
                    except Exception as e:
                        logger.error(f"Failed to read output file: {e}")
                else:
                    logger.warning("Output file was not created by helper")

                # Parse JSON response from helper
                try:
                    if stdout.strip():
                        result_data = json.loads(stdout.strip())
                        command_result = CommandResult(
                            success=result_data.get('success', False),
                            output=result_data.get('output', ''),
                            error=result_data.get('error', ''),
                            return_code=result_data.get('return_code', return_code),
                            execution_time=result_data.get('execution_time', execution_time),
                            command=result_data.get('command', self.command)
                        )

                        if command_result.success:
                            logger.info(f"Elevated command succeeded in {execution_time:.2f}s")
                        else:
                            logger.error(f"Elevated command failed: {command_result.error}")
                    else:
                        # No output - likely UAC was denied or helper crashed
                        error_msg = "No response from helper"
                        if return_code == 1223:  # ERROR_CANCELLED - UAC denied
                            error_msg = "Administrator privileges required (UAC prompt was denied)"
                        else:
                            error_msg = f"Helper process failed with exit code {return_code}"

                        command_result = CommandResult(
                            success=False,
                            output="",
                            error=error_msg,
                            return_code=return_code,
                            execution_time=execution_time,
                            command=self.command
                        )
                        logger.error(f"Helper process failed: {error_msg}")

                except json.JSONDecodeError as e:
                    logger.error(f"Failed to parse helper output: {stdout}")
                    command_result = CommandResult(
                        success=False,
                        output=stdout,
                        error=f"Invalid response from helper: {str(e)}",
                        return_code=return_code,
                        execution_time=execution_time,
                        command=self.command
                    )

            finally:
                # Clean up temp file
                try:
                    if os.path.exists(output_file):
                        os.unlink(output_file)
                except Exception as e:
                    logger.warning(f"Failed to delete temp file: {e}")

        except FileNotFoundError as e:
            execution_time = time.time() - start_time
            logger.error(f"Helper executable not found: {e}")
            command_result = CommandResult(
                success=False,
                output="",
                error="Port Manager helper not found. Please reinstall the application.",
                return_code=-2,
                execution_time=execution_time,
                command=self.command
            )

        except Exception as e:
            execution_time = time.time() - start_time
            logger.exception(f"Unexpected error launching helper: {self.command}")
            command_result = CommandResult(
                success=False,
                output="",
                error=f"Failed to launch elevated helper: {str(e)}",
                return_code=-3,
                execution_time=execution_time,
                command=self.command
            )

        self.result = command_result
        self.command_finished.emit(command_result)

    def terminate(self):
        """Terminate the helper process if running."""
        if self._process:
            # Check if process is still running
            exit_code = wintypes.DWORD()
            if GetExitCodeProcess(self._process, ctypes.byref(exit_code)):
                STILL_ACTIVE = 259  # Windows constant for still-running process
                if exit_code.value == STILL_ACTIVE:
                    logger.warning("Terminating elevated helper process")
                    # Terminate the process using Windows API
                    TerminateProcess = ctypes.windll.kernel32.TerminateProcess
                    TerminateProcess.argtypes = [wintypes.HANDLE, ctypes.c_uint]
                    TerminateProcess.restype = wintypes.BOOL
                    TerminateProcess(self._process, 1)
                # Close handle
                CloseHandle(self._process)
                self._process = None
        super().terminate()


# ============================================================================
# UI Icons
# ============================================================================

class Icons:
    """Circular SVG icons matching application design"""

    # Color constants from terminal_dialog.py
    BLUE_PRIMARY = "#0078D4"
    BLUE_STROKE = "#106EBE"
    GREEN_PRIMARY = "#28A745"
    GREEN_STROKE = "#1E7E34"
    PURPLE_PRIMARY = "#6F42C1"
    PURPLE_STROKE = "#5A2D91"
    RED_PRIMARY = "#DC3545"
    RED_STROKE = "#B02A37"
    YELLOW_WARNING = "#FFC107"

    @staticmethod
    def create():
        """Blue circular create icon"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#0078D4" stroke="#106EBE" stroke-width="1"/>
            <path d="M16 8 L16 24 M8 16 L24 16" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def quick_setup():
        """Green circular quick setup icon (lightning bolt)"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#28A745" stroke="#1E7E34" stroke-width="1"/>
            <path d="M18 8 L10 16 L14 16 L14 24 L22 16 L18 16 Z" fill="#FFFFFF"/>
        </svg>"""

    @staticmethod
    def remove():
        """Red circular remove icon"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#DC3545" stroke="#B02A37" stroke-width="1"/>
            <path d="M10 10 L22 22 M22 10 L10 22" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def refresh():
        """Green circular refresh icon"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#28A745" stroke="#1E7E34" stroke-width="1"/>
            <path d="M16 9 A7 7 0 1 1 9 16 A7 7 0 0 1 12.8 11.2" stroke="#FFFFFF" stroke-width="2.5" fill="none" stroke-linecap="round"/>
            <path d="M11 9 L15 9 L15 13" stroke="#FFFFFF" stroke-width="2.5" fill="none" stroke-linecap="round"/>
        </svg>"""

    @staticmethod
    def close():
        """Gray circular close icon"""
        return """<svg width="32" height="32" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg">
            <circle cx="16" cy="16" r="14" fill="#6C757D" stroke="#5A6268" stroke-width="1"/>
            <path d="M10 10 L22 22 M22 10 L10 22" stroke="#FFFFFF" stroke-width="3" stroke-linecap="round"/>
        </svg>"""

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


class VirtualPortToolbar(QWidget):
    """Ribbon-style toolbar for Virtual Port Manager dialog"""

    # Signals
    refresh_clicked = pyqtSignal()
    create_clicked = pyqtSignal()
    close_clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._setup_ui()

    def _create_button(self, text: str, icon_svg: str, tooltip: str, callback) -> QPushButton:
        """Create a toolbar button with icon and tooltip"""
        button = QPushButton(text)
        button.setIcon(Icons.svg_to_icon(icon_svg))
        button.setIconSize(QSize(16, 16))
        button.setToolTip(tooltip)
        button.clicked.connect(callback)
        return button

    def _setup_ui(self):
        """Setup toolbar UI matching RibbonToolbar style"""
        self.setFixedHeight(48)

        # Main layout
        layout = QHBoxLayout(self)
        layout.setSpacing(4)
        layout.setContentsMargins(5, 5, 5, 5)

        # Create buttons
        self.refresh_button = self._create_button("Refresh", Icons.refresh(), "Refresh port list", self.refresh_clicked.emit)
        self.create_button = self._create_button("Create", Icons.create(), "Create new port pair", self.create_clicked.emit)
        self.close_button = self._create_button("Close", Icons.close(), "Close Virtual Port Manager", self.close_clicked.emit)

        layout.addWidget(self.refresh_button)
        layout.addWidget(self.create_button)
        layout.addWidget(self.close_button)

        # Stretch to push status to right
        layout.addStretch()

        # Status label
        self.status_label = QLabel("Ready")
        self.status_label.setFont(QFont("Segoe UI", 9))
        layout.addWidget(self.status_label)

        # Availability counter
        self.counter_label = QLabel("")
        self.counter_label.setFont(QFont("Segoe UI", 9))
        self.counter_label.setStyleSheet("color: #6C757D;")  # Subtle gray
        layout.addWidget(self.counter_label)

    def set_status(self, message: str, available: Optional[int] = None, total: Optional[int] = None):
        """Update status text and availability counter"""
        self.status_label.setText(message)

        if available is not None and total is not None:
            self.counter_label.setText(f"({available}/{total})")
        else:
            self.counter_label.setText("")

    def set_buttons_enabled(self, enabled: bool):
        """Enable/disable all action buttons"""
        self.refresh_button.setEnabled(enabled)
        self.create_button.setEnabled(enabled)
        # Close button always stays enabled


class VirtualPortDialog(QDialog):
    """Dialog for managing virtual COM port pairs"""

    SETUPC_PATH = r"C:\Program Files (x86)\com0com\setupc.exe"

    # Virtual port auto-assignment range
    VIRTUAL_PORT_RANGE_START = 150
    VIRTUAL_PORT_RANGE_END = 199  # Max 199 so pair 199↔200 stays in range

    # Default com0com port parameters
    PORT_PARAMS = {
        'EmuBR': 'yes',
        'EmuOverrun': 'yes',
        'ExclusiveMode': 'no',
        'AllDataBits': 'yes',
        'cts': 'rrts',
        'dsr': 'rdtr',
        'dcd': 'rdtr'
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        logger.info("Initializing VirtualPortDialog")
        self.worker = None
        self.port_pairs = []  # Store PortPair objects instead of dict
        self.working_directory = os.path.dirname(self.SETUPC_PATH) if self.SETUPC_PATH else None
        self._operation_in_progress = False  # Flag to prevent concurrent operations
        self._closing = False  # Flag to prevent operations during shutdown
        self._load_timer_id = None  # Track QTimer for cleanup
        self._dev_mode = not getattr(sys, 'frozen', False)  # Check if running as script

        # Set window flags to show proper title bar with minimize/maximize/close buttons
        self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.WindowTitleHint |
                           Qt.WindowType.WindowCloseButtonHint | Qt.WindowType.WindowMinimizeButtonHint |
                           Qt.WindowType.WindowMaximizeButtonHint)

        self.setWindowTitle("Virtual Port Manager")
        self.setModal(True)
        self.resize(520, 340)

        self._setup_ui()
        self._check_com0com_installed()
        self._check_dev_mode()

        # Center on screen
        self._center_on_screen()

        # Load existing pairs after UI is setup (store timer ID for cleanup)
        self._load_timer_id = QTimer.singleShot(100, self._safe_load_existing_pairs)

    def _setup_ui(self):
        """Setup the dialog UI with ribbon-style toolbar"""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # === Toolbar ===
        self.toolbar = VirtualPortToolbar(self)
        main_layout.addWidget(self.toolbar)

        # === Port Pairs Table ===
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Port A", "Port B", "Actions"])
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(2, 100)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setFont(QFont("Consolas", 10))  # Monospace for port numbers
        self.table.setMinimumHeight(240)

        main_layout.addWidget(self.table)

        # Wire up toolbar signals
        self.toolbar.refresh_clicked.connect(self._load_existing_pairs)
        self.toolbar.create_clicked.connect(self._create_new_pair)
        self.toolbar.close_clicked.connect(self.accept)

    def _center_on_screen(self):
        """Center the dialog on screen"""
        # Get the screen geometry
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            # Calculate center position
            x = (screen_geometry.width() - self.width()) // 2
            y = (screen_geometry.height() - self.height()) // 2
            self.move(x, y)

    def _check_com0com_installed(self):
        """Check if com0com is installed and disable buttons if not"""
        if not os.path.exists(self.SETUPC_PATH):
            self.toolbar.create_button.setEnabled(False)
            self.toolbar.set_status("com0com not installed at " + self.SETUPC_PATH)

    def _check_dev_mode(self):
        """Check if running in dev mode and show warning"""
        if self._dev_mode:
            # Check if helper is available
            script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            helper_exe = os.path.join(script_dir, 'dist', 'SerialPortManager.exe')
            helper_py = os.path.join(script_dir, 'port_manager_helper.py')

            if not os.path.exists(helper_exe) and not os.path.exists(helper_py):
                logger.warning("DEV MODE: No helper found - will attempt direct execution")
                # Show dev mode indicator in status
                current_status = self.toolbar.status_label.text()
                if current_status == "Ready":
                    self.toolbar.set_status("DEV MODE: Direct execution (may require admin console)")
                self.toolbar.status_label.setStyleSheet("color: #FFC107;")  # Yellow warning

    def _cleanup_worker(self):
        """Cleanup existing worker thread safely"""
        if self.worker:
            # Disconnect all signals
            try:
                self.worker.command_finished.disconnect()
            except TypeError:
                pass  # Already disconnected or no connections

            # If still running, terminate and wait
            if self.worker.isRunning():
                logger.debug("Cleaning up running worker thread")
                self.worker.terminate()  # Terminate helper process if running
                self.worker.wait(1000)  # Wait up to 1s

    def _execute_command(self, command: str, callback, timeout: int = 30):
        """Execute setupc command via elevated helper in background worker thread"""
        if self._operation_in_progress:
            logger.warning("Command execution blocked - operation already in progress")
            return
        if self._closing:
            logger.warning("Command execution blocked - dialog is closing")
            return

        logger.debug(f"Starting elevated command execution with timeout={timeout}s")
        self._operation_in_progress = True
        self._cleanup_worker()

        # Use ElevatedHelperWorker instead of direct command execution
        self.worker = ElevatedHelperWorker(self.SETUPC_PATH, command, timeout=timeout)
        self.worker.command_finished.connect(callback)
        self.worker.start()

    def _safe_load_existing_pairs(self):
        """Safe wrapper for _load_existing_pairs that checks if dialog is closing"""
        if not self._closing:
            self._load_existing_pairs()

    def _load_existing_pairs(self):
        """Load existing virtual port pairs using elevated helper"""
        if not os.path.exists(self.SETUPC_PATH):
            return

        self._show_progress("Loading...")
        # Just pass the command argument (helper will add setupc path)
        self._execute_command("list", self._on_list_result, timeout=30)

    def _on_list_result(self, result: CommandResult):
        """Handle list command result"""
        if self._closing:
            logger.debug("Ignoring list result - dialog is closing")
            return  # Dialog is closing, ignore result

        self._operation_in_progress = False
        self._hide_progress()

        if result.success:
            try:
                # Use PortListParser to parse the output
                self.port_pairs = PortListParser.parse_port_list(result.output)
                logger.info(f"Successfully loaded {len(self.port_pairs)} port pairs")
                self._update_table()

                # Calculate availability
                _, available = self._get_port_availability()
                total = self.VIRTUAL_PORT_RANGE_END - self.VIRTUAL_PORT_RANGE_START + 1

                if len(self.port_pairs) > 0:
                    self.toolbar.set_status("Ready", available, total)
                else:
                    self.toolbar.set_status("No virtual port pairs found", available, total)
            except Exception as e:
                logger.exception(f"Failed to parse port list output")
                self.toolbar.set_status(f"Error parsing pairs - {str(e)}")
        else:
            error_msg = result.get_error_message()
            logger.error(f"List command failed: {error_msg}")
            self.toolbar.set_status(f"Error - {error_msg}")

        # Update button state with available slots
        self._update_create_button_state()

    def _create_port_cell(self, port_name: str) -> QTableWidgetItem:
        """Create a table cell item for a port name"""
        item = QTableWidgetItem(port_name)
        item.setFont(QFont("Consolas", 10))
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        return item

    def _create_action_cell(self, pair: PortPair) -> QWidget:
        """Create action cell widget with remove button"""
        port_a_name = pair.port_a.port_name or pair.port_a.identifier
        port_b_name = pair.port_b.port_name or pair.port_b.identifier

        remove_button = QPushButton()
        remove_button.setIcon(Icons.svg_to_icon(Icons.remove()))
        remove_button.setIconSize(QSize(20, 20))
        remove_button.setToolTip(f"Remove {port_a_name} ↔ {port_b_name}")
        remove_button.setMaximumWidth(100)
        remove_button.clicked.connect(lambda checked, p=pair: self._remove_pair(p))

        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(5, 2, 5, 2)
        layout.addWidget(remove_button)
        return widget

    def _update_table(self):
        """Update the table with existing pairs"""
        self.table.setRowCount(0)

        for pair in self.port_pairs:
            row = self.table.rowCount()
            self.table.insertRow(row)

            # Port A and Port B cells
            port_a_name = pair.port_a.port_name or pair.port_a.identifier
            port_b_name = pair.port_b.port_name or pair.port_b.identifier
            self.table.setItem(row, 0, self._create_port_cell(port_a_name))
            self.table.setItem(row, 1, self._create_port_cell(port_b_name))

            # Action cell with remove button
            self.table.setCellWidget(row, 2, self._create_action_cell(pair))

    def _get_port_availability(self) -> tuple[Optional[int], int]:
        """
        Get port availability info: (next_available_port, total_available_count).
        Returns (None, 0) if no ports available.
        """
        existing = self._get_existing_port_numbers()
        next_port = None
        count = 0

        for candidate in range(self.VIRTUAL_PORT_RANGE_START, self.VIRTUAL_PORT_RANGE_END + 1):
            if candidate not in existing and (candidate + 1) not in existing:
                if next_port is None:
                    next_port = candidate
                count += 1

        return next_port, count

    def _update_create_button_state(self):
        """Update create button enabled state and tooltip based on available ports."""
        next_port, available_count = self._get_port_availability()
        max_slots = self.VIRTUAL_PORT_RANGE_END - self.VIRTUAL_PORT_RANGE_START + 1

        if next_port is not None:
            if os.path.exists(self.SETUPC_PATH):
                self.toolbar.create_button.setEnabled(True)
            self.toolbar.create_button.setToolTip(
                f"Create COM{next_port}↔COM{next_port+1} ({available_count}/{max_slots} slots available)"
            )
        else:
            self.toolbar.create_button.setEnabled(False)
            self.toolbar.create_button.setToolTip("No available port slots in range 150-200")

    def _create_new_pair(self):
        """Create new port pair with auto-assigned port numbers"""
        if not os.path.exists(self.SETUPC_PATH):
            self.toolbar.set_status("com0com not installed")
            return

        # Find next available port pair
        port_a, _ = self._get_port_availability()
        if port_a is None:
            self.toolbar.set_status("No available ports in range 150-200")
            return

        # Create pair with fixed parameters
        self._show_progress(f"Creating...")
        self._create_pair_with_full_params(port_a, port_a + 1)

    def _build_port_params(self, port_name: str) -> str:
        """Build parameter string for a port (quoted for safety)"""
        # Quote port name for safety (in case of future configuration changes)
        params = [f"PortName={shlex.quote(port_name)}"]
        params.extend(f"{k}={shlex.quote(str(v))}" for k, v in self.PORT_PARAMS.items())
        return ','.join(params)

    def _create_pair_with_full_params(self, port_a: int, port_b: int):
        """Create port pair with all com0com parameters using elevated helper"""
        com_a = f"COM{port_a}"
        com_b = f"COM{port_b}"

        port_a_params = self._build_port_params(com_a)
        port_b_params = self._build_port_params(com_b)

        # Just pass command arguments (helper will add setupc path)
        command = f'install {port_a_params} {port_b_params}'
        self._execute_command(command, lambda result: self._on_pair_created(result, com_a, com_b), timeout=45)

    def _on_pair_created(self, result: CommandResult, com_a: str, com_b: str):
        """Handle result of creating a port pair"""
        if self._closing:
            logger.debug("Ignoring create result - dialog is closing")
            return  # Dialog is closing, ignore result

        self._operation_in_progress = False
        self._hide_progress()

        if result.success:
            logger.info(f"Successfully created port pair: {com_a} ↔ {com_b}")
            self.toolbar.set_status(f"Created {com_a} ↔ {com_b}")
            # Refresh the table
            self._load_existing_pairs()
        else:
            error_msg = result.get_error_message()
            logger.error(f"Failed to create port pair {com_a} ↔ {com_b}: {error_msg}")
            self.toolbar.set_status(f"Error - {error_msg}")

    def _get_existing_port_numbers(self) -> set:
        """Get set of existing COM port numbers"""
        port_numbers = set()
        for pair in self.port_pairs:
            # Extract port numbers from COM names (e.g., "COM131" -> 131)
            port_a = pair.port_a.port_name
            port_b = pair.port_b.port_name

            if port_a and port_a.startswith('COM'):
                try:
                    port_numbers.add(int(port_a[3:]))
                except ValueError:
                    pass

            if port_b and port_b.startswith('COM'):
                try:
                    port_numbers.add(int(port_b[3:]))
                except ValueError:
                    pass

        return port_numbers


    def _remove_pair(self, pair: PortPair):
        """Remove a virtual port pair"""
        port_a_name = pair.port_a.port_name or pair.port_a.identifier
        port_b_name = pair.port_b.port_name or pair.port_b.identifier

        # Directly remove without confirmation dialog
        self._show_progress("Removing...")
        self._remove_pair_threaded(pair.number)

    def _remove_pair_threaded(self, pair_number: int):
        """Remove port pair using elevated helper"""
        # Just pass command arguments (helper will add setupc path)
        command = f'remove {pair_number}'
        self._execute_command(command, self._on_remove_result, timeout=30)

    def _on_remove_result(self, result: CommandResult):
        """Handle remove command result"""
        if self._closing:
            logger.debug("Ignoring remove result - dialog is closing")
            return  # Dialog is closing, ignore result

        self._operation_in_progress = False
        self._hide_progress()

        if result.success:
            logger.info("Successfully removed port pair")
            self.toolbar.set_status("Port pair removed")
            # Refresh the table
            self._load_existing_pairs()
        else:
            error_msg = result.get_error_message()
            logger.error(f"Failed to remove port pair: {error_msg}")
            self.toolbar.set_status(f"Error - {error_msg}")

    def _show_progress(self, message: str):
        """Show progress with message and disable buttons"""
        self.toolbar.set_status(message)
        self.toolbar.set_buttons_enabled(False)

    def _hide_progress(self):
        """Hide progress and re-enable buttons"""
        self.toolbar.set_buttons_enabled(True)

    def closeEvent(self, event):
        """Handle dialog close - cleanup running workers and helper processes"""
        logger.info("VirtualPortDialog closing")
        self._closing = True

        if self.worker and self.worker.isRunning():
            logger.warning("Worker thread still running during close - waiting for termination")
            # Disconnect signals to prevent callbacks on destroyed dialog
            try:
                self.worker.command_finished.disconnect()
            except TypeError:
                pass  # Already disconnected

            # Terminate helper process and wait for worker
            self.worker.terminate()
            if not self.worker.wait(2000):  # 2 second timeout
                # Force terminate if still running
                logger.error("Worker thread did not terminate gracefully - forcing termination")
                self.worker.terminate()
                self.worker.wait(1000)  # Wait up to 1 more second
            else:
                logger.info("Worker thread terminated gracefully")

        event.accept()
