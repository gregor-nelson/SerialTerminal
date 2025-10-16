#!/usr/bin/env python3
"""
Core components for Hub4com GUI Launcher
Contains data classes, managers, and worker threads
"""

import subprocess
import time
import threading
from pathlib import Path
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from enum import Enum
from dataclasses import field

# Try to import winreg, fallback gracefully if not available
try:
    import winreg
    WINREG_AVAILABLE = True
except ImportError:
    WINREG_AVAILABLE = False
    print("Warning: winreg module not available. Port scanning will be limited.")

# Try to import serial, fallback gracefully if not available
try:
    import serial
    SERIAL_AVAILABLE = True
except ImportError:
    SERIAL_AVAILABLE = False
    print("Warning: pyserial module not available. Port monitoring will be disabled.")

# Try to import WMI for advanced hardware detection, fallback gracefully
try:
    import wmi
    WMI_AVAILABLE = True
except ImportError:
    WMI_AVAILABLE = False

from PyQt6.QtCore import QThread, pyqtSignal, QMutex, QSettings
from PyQt6.QtWidgets import QApplication


class PortStatus(Enum):
    """Enumeration for serial port status"""
    AVAILABLE = "Available"
    IN_USE = "In Use"
    BUSY = "Busy"
    RESERVED = "Reserved"
    ERROR = "Error"
    UNKNOWN = "Unknown"


@dataclass
class WindowConfig:
    """Configuration for window sizing and layout"""
    width: int
    height: int
    x: int
    y: int
    is_small_screen: bool
    min_width: int = 800
    min_height: int = 600


@dataclass
class SerialPortInfo:
    """Information about a detected serial port"""
    port_name: str
    device_name: str
    port_type: str  # 'Physical', 'Virtual (Moxa)', 'Virtual (Other)', 'Virtual (COM0COM)'
    registry_key: str
    description: str = ""
    is_moxa: bool = False
    moxa_details: Optional[Dict] = None
    
    # Enhanced fields for detailed port information
    manufacturer: str = "Unknown"
    status: PortStatus = PortStatus.UNKNOWN
    location: str = ""  # USB Hub, Internal, Network, etc.
    capabilities: List[str] = field(default_factory=list)  # Hardware Flow Control, High Speed, etc.
    last_activity: Optional[datetime] = None
    driver_version: str = ""
    hardware_id: str = ""


@dataclass
class Com0comPortPair:
    """Information about a com0com virtual port pair"""
    port_a: str  # e.g., "CNCA0"
    port_b: str  # e.g., "CNCB0"
    port_a_params: Dict[str, str]  # Parameters like PortName, EmuBR, etc.
    port_b_params: Dict[str, str]


@dataclass
class SerialPacketInfo:
    """Information about a detected serial packet/frame"""
    timestamp: float
    size: int
    direction: str  # 'RX' or 'TX'
    data: bytes
    inter_frame_gap: float = 0.0  # Time since last packet
    has_errors: bool = False
    error_type: str = ""


@dataclass
class AdvancedStatistics:
    """Advanced statistics for serial port monitoring"""
    # Peak/Average rates
    rx_peak_rate: float = 0.0
    tx_peak_rate: float = 0.0
    rx_average_rate: float = 0.0
    tx_average_rate: float = 0.0
    
    # Error tracking
    rx_errors: int = 0
    tx_errors: int = 0
    timeout_errors: int = 0
    buffer_overruns: int = 0
    framing_errors: int = 0
    parity_errors: int = 0
    
    # Packet statistics
    rx_packet_count: int = 0
    tx_packet_count: int = 0
    rx_packet_sizes: List[int] = None
    tx_packet_sizes: List[int] = None
    average_inter_frame_gap: float = 0.0
    min_inter_frame_gap: float = float('inf')
    max_inter_frame_gap: float = 0.0
    
    def __post_init__(self):
        if self.rx_packet_sizes is None:
            self.rx_packet_sizes = []
        if self.tx_packet_sizes is None:
            self.tx_packet_sizes = []


class PortCapabilityAnalyzer:
    """Analyzes serial port hardware capabilities, driver information, and system integration"""
    
    def __init__(self, fast_mode=True):
        """Initialize the port capability analyzer"""
        self.wmi_available = WMI_AVAILABLE
        self.winreg_available = WINREG_AVAILABLE
        self._wmi_connection = None
        self._manufacturer_cache = {}
        self._wmi_ports_cache = None
        self._wmi_drivers_cache = None
        self._enable_fast_mode = fast_mode
        self._cache_initialized = False
        
    def _get_wmi_connection(self):
        """Get or create WMI connection with error handling"""
        if not self.wmi_available:
            return None
            
        if self._wmi_connection is None:
            try:
                self._wmi_connection = wmi.WMI()
            except Exception:
                self.wmi_available = False
                return None
        return self._wmi_connection
    
    def _initialize_wmi_cache(self):
        """Initialize WMI cache with bulk data retrieval for performance"""
        if self._cache_initialized or not self.wmi_available:
            return
        
        try:
            wmi_conn = self._get_wmi_connection()
            if wmi_conn:
                # Cache all serial ports in one query
                self._wmi_ports_cache = {}
                for port in wmi_conn.Win32_SerialPort():
                    if hasattr(port, 'DeviceID') and port.DeviceID:
                        self._wmi_ports_cache[port.DeviceID] = port
                
                # Skip driver cache - not needed for simplified dialog
                self._wmi_drivers_cache = {}
                
            self._cache_initialized = True
        except Exception as e:
            print(f"Warning: WMI cache initialization failed: {str(e)}")
            self._cache_initialized = True  # Prevent retries
    
    def analyze_port_capabilities(self, port_info: SerialPortInfo) -> SerialPortInfo:
        """
        Analyze port capabilities and enhance SerialPortInfo with detailed information
        
        Args:
            port_info: Basic SerialPortInfo to enhance
            
        Returns:
            SerialPortInfo: Enhanced port information with capabilities
        """
        try:
            # Initialize WMI cache if needed
            if not self._cache_initialized:
                self._initialize_wmi_cache()
            
            # Always detect manufacturer (fast operation)
            port_info.manufacturer = self._detect_manufacturer(port_info)
            
            # Always analyze basic capabilities (fast operation)
            port_info.capabilities = self._analyze_hardware_capabilities(port_info)
            
            # Always get location (fast operation)
            port_info.location = self._get_connection_topology(port_info)
            
            if self._enable_fast_mode:
                # Fast mode: use heuristic status detection
                port_info.status = self._get_heuristic_status(port_info)
                # Skip expensive driver version and hardware ID lookup
                port_info.driver_version = ""
                port_info.hardware_id = ""
            else:
                # Full mode: complete analysis
                port_info.driver_version = self._get_driver_version(port_info)
                port_info.status = self._check_port_status(port_info)
                port_info.hardware_id = self._get_hardware_id(port_info)
            
        except Exception as e:
            # Production-grade error handling - don't crash on capability detection failures
            print(f"Warning: Capability analysis failed for {port_info.port_name}: {str(e)}")
            
        return port_info
    
    def _detect_manufacturer(self, port_info: SerialPortInfo) -> str:
        """Detect port manufacturer from hardware information"""
        # Check cache first
        cache_key = f"{port_info.device_name}_{port_info.port_name}"
        if cache_key in self._manufacturer_cache:
            return self._manufacturer_cache[cache_key]
        
        manufacturer = "Unknown"
        
        try:
            # Virtual port manufacturer detection
            if port_info.port_type.startswith("Virtual"):
                if "COM0COM" in port_info.port_type:
                    manufacturer = "com0com"
                elif "Moxa" in port_info.port_type:
                    manufacturer = "Moxa"
                elif "VSPD" in port_info.device_name.upper():
                    manufacturer = "Eltima"
                else:
                    manufacturer = "Virtual"
            else:
                # Physical port manufacturer detection
                manufacturer = self._detect_physical_manufacturer(port_info)
        
        except Exception:
            # Fallback to device name pattern matching
            manufacturer = self._fallback_manufacturer_detection(port_info)
        
        # Cache the result
        self._manufacturer_cache[cache_key] = manufacturer
        return manufacturer
    
    def _detect_physical_manufacturer(self, port_info: SerialPortInfo) -> str:
        """Detect manufacturer for physical ports using cached WMI and optimized registry"""
        try:
            # Try cached WMI data first
            if self._wmi_ports_cache and port_info.port_name in self._wmi_ports_cache:
                port = self._wmi_ports_cache[port_info.port_name]
                if hasattr(port, 'Manufacturer') and port.Manufacturer:
                    return port.Manufacturer
            
            # Fast mode: skip expensive registry scanning
            if self._enable_fast_mode:
                return self._fallback_manufacturer_detection(port_info)
            
            # Full mode: use optimized registry scanning
            return self._optimized_registry_manufacturer_detection(port_info)
            
        except Exception:
            return self._fallback_manufacturer_detection(port_info)
    
    def _registry_manufacturer_detection(self, port_info: SerialPortInfo) -> str:
        """Detect manufacturer from Windows registry"""
        if not self.winreg_available:
            return "Unknown"
        
        try:
            # Search in device enumeration keys
            enum_key = r"SYSTEM\CurrentControlSet\Enum"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, enum_key) as key:
                # Common device paths for serial ports
                device_paths = ["USB", "FTDIBUS", "ACPI", "PCI"]
                
                for device_path in device_paths:
                    try:
                        with winreg.OpenKey(key, device_path) as device_key:
                            # Enumerate devices under this path
                            i = 0
                            while i < 100:  # Reasonable limit
                                try:
                                    device_name = winreg.EnumKey(device_key, i)
                                    if self._device_matches_port(device_name, port_info):
                                        with winreg.OpenKey(device_key, device_name) as specific_device:
                                            # Look for manufacturer info
                                            j = 0
                                            while j < 50:
                                                try:
                                                    instance_name = winreg.EnumKey(specific_device, j)
                                                    with winreg.OpenKey(specific_device, instance_name) as instance:
                                                        try:
                                                            mfg, _ = winreg.QueryValueEx(instance, "Mfg")
                                                            if mfg and mfg != "Unknown":
                                                                return mfg
                                                        except FileNotFoundError:
                                                            pass
                                                    j += 1
                                                except OSError:
                                                    break
                                    i += 1
                                except OSError:
                                    break
                    except FileNotFoundError:
                        continue
        except Exception:
            pass
        
        return self._fallback_manufacturer_detection(port_info)
    
    def _optimized_registry_manufacturer_detection(self, port_info: SerialPortInfo) -> str:
        """Optimized registry manufacturer detection with reduced iteration"""
        if not self.winreg_available:
            return "Unknown"
        
        try:
            # Search in device enumeration keys with reduced limits
            enum_key = r"SYSTEM\CurrentControlSet\Enum"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, enum_key) as key:
                # Prioritized device paths for serial ports
                device_paths = ["USB", "FTDIBUS"]  # Most common first
                
                for device_path in device_paths:
                    try:
                        with winreg.OpenKey(key, device_path) as device_key:
                            # Reduced enumeration limits for performance
                            i = 0
                            while i < 20:  # Reduced from 100
                                try:
                                    device_name = winreg.EnumKey(device_key, i)
                                    if self._device_matches_port(device_name, port_info):
                                        with winreg.OpenKey(device_key, device_name) as specific_device:
                                            # Reduced instance enumeration
                                            j = 0
                                            while j < 5:  # Reduced from 50
                                                try:
                                                    instance_name = winreg.EnumKey(specific_device, j)
                                                    with winreg.OpenKey(specific_device, instance_name) as instance:
                                                        try:
                                                            mfg, _ = winreg.QueryValueEx(instance, "Mfg")
                                                            if mfg and mfg != "Unknown":
                                                                return mfg
                                                        except FileNotFoundError:
                                                            pass
                                                    j += 1
                                                except OSError:
                                                    break
                                    i += 1
                                except OSError:
                                    break
                    except FileNotFoundError:
                        continue
        except Exception:
            pass
        
        return self._fallback_manufacturer_detection(port_info)
    
    def _get_heuristic_status(self, port_info: SerialPortInfo) -> PortStatus:
        """Fast heuristic status detection without expensive I/O operations"""
        # Virtual ports are typically always available
        if port_info.port_type.startswith("Virtual"):
            if "COM0COM" in port_info.port_type:
                return PortStatus.AVAILABLE  # com0com ports are usually available
            elif "Moxa" in port_info.port_type:
                return PortStatus.UNKNOWN  # Network status requires checking
            else:
                return PortStatus.AVAILABLE
        
        # Physical ports - assume available unless we know otherwise
        return PortStatus.AVAILABLE
    
    def _device_matches_port(self, device_name: str, port_info: SerialPortInfo) -> bool:
        """Check if a device name might match the port"""
        device_lower = device_name.lower()
        port_device_lower = port_info.device_name.lower()
        
        # Look for common patterns
        if "serial" in device_lower or "uart" in device_lower:
            return True
        if port_info.port_name.replace("COM", "") in device_name:
            return True
        if any(part in device_lower for part in port_device_lower.split() if len(part) > 3):
            return True
        
        return False
    
    def _fallback_manufacturer_detection(self, port_info: SerialPortInfo) -> str:
        """Fallback manufacturer detection using device name patterns"""
        device_name = port_info.device_name.upper()
        
        # Common manufacturer patterns
        if "FTDI" in device_name:
            return "FTDI"
        elif "PROLIFIC" in device_name or "PL23" in device_name:
            return "Prolific"
        elif "CH34" in device_name or "CH91" in device_name:
            return "WCH"
        elif "CP21" in device_name or "SILABS" in device_name:
            return "Silicon Labs"
        elif "INTEL" in device_name:
            return "Intel"
        elif "VIA" in device_name:
            return "VIA"
        elif "MOXA" in device_name or "NPDRV" in device_name:
            return "Moxa"
        elif "16550" in device_name:
            return "Generic 16550"
        else:
            return "Unknown"
    
    def _get_driver_version(self, port_info: SerialPortInfo) -> str:
        """Get driver version information from cache"""
        if not self._wmi_drivers_cache:
            return ""
        
        try:
            # Search cached driver data
            for device_name, driver in self._wmi_drivers_cache.items():
                if port_info.port_name in device_name or \
                   any(part in device_name for part in port_info.device_name.split() if len(part) > 3):
                    if hasattr(driver, 'DriverVersion') and driver.DriverVersion:
                        return driver.DriverVersion
        except Exception:
            pass
        
        return ""
    
    def _analyze_hardware_capabilities(self, port_info: SerialPortInfo) -> List[str]:
        """Analyze hardware capabilities of the port"""
        capabilities = []
        
        try:
            # Virtual port capabilities
            if port_info.port_type.startswith("Virtual"):
                if "COM0COM" in port_info.port_type:
                    capabilities.extend(["Null Modem", "Configurable"])
                    if port_info.moxa_details:
                        capabilities.append("Network Capable")
                elif "Moxa" in port_info.port_type:
                    capabilities.extend(["Network Serial", "TCP/IP"])
                else:
                    capabilities.append("Virtual")
            else:
                # Physical port capabilities
                capabilities.extend(self._detect_physical_capabilities(port_info))
                
        except Exception:
            capabilities = ["Standard"]
        
        return capabilities if capabilities else ["Standard"]
    
    def _detect_physical_capabilities(self, port_info: SerialPortInfo) -> List[str]:
        """Detect capabilities for physical serial ports"""
        capabilities = []
        manufacturer = port_info.manufacturer
        device_name = port_info.device_name.upper()
        
        # High-speed capabilities
        if any(pattern in device_name for pattern in ["FT232", "FT4232", "CH340", "CP21"]):
            capabilities.append("High Speed")
        
        # Flow control capabilities
        if any(pattern in manufacturer.upper() for pattern in ["FTDI", "PROLIFIC", "SILICON"]):
            capabilities.append("Hardware Flow Control")
        
        # USB capabilities
        if any(pattern in device_name for pattern in ["USB", "FT", "CH34", "CP21"]):
            capabilities.append("USB")
        
        # Multi-port capabilities
        if any(pattern in device_name for pattern in ["4232", "2232", "QUAD"]):
            capabilities.append("Multi-port")
        
        return capabilities
    
    def _get_connection_topology(self, port_info: SerialPortInfo) -> str:
        """Get connection topology/location information"""
        try:
            # Virtual port locations
            if port_info.port_type.startswith("Virtual"):
                if "COM0COM" in port_info.port_type:
                    return "Virtual Pair"
                elif "Moxa" in port_info.port_type:
                    return "Network"
                else:
                    return "Virtual"
            
            # Physical port topology detection
            return self._detect_physical_location(port_info)
            
        except Exception:
            return "Unknown"
    
    def _detect_physical_location(self, port_info: SerialPortInfo) -> str:
        """Detect physical location for hardware ports using cache"""
        try:
            # Try cached WMI data first
            if self._wmi_ports_cache and port_info.port_name in self._wmi_ports_cache:
                port = self._wmi_ports_cache[port_info.port_name]
                if hasattr(port, 'PNPDeviceID') and port.PNPDeviceID:
                    device_id = port.PNPDeviceID.upper()
                    if device_id.startswith("USB"):
                        return self._parse_usb_location(device_id)
                    elif device_id.startswith("PCI"):
                        return "Internal PCI"
                    elif device_id.startswith("ACPI"):
                        return "Internal"
        except Exception:
            pass
        
        return self._fallback_location_detection(port_info)
    
    def _parse_usb_location(self, device_id: str) -> str:
        """Parse USB device location from device ID"""
        try:
            # Extract USB hub information if available
            if "VID_" in device_id and "PID_" in device_id:
                # This is a USB device
                parts = device_id.split("\\")
                if len(parts) > 2:
                    usb_info = parts[2]
                    # Try to extract hub/port information
                    if "&" in usb_info:
                        location_part = usb_info.split("&")[-1]
                        if location_part.isdigit():
                            return f"USB Port {location_part}"
                return "USB"
        except Exception:
            pass
        
        return "USB"
    
    def _fallback_location_detection(self, port_info: SerialPortInfo) -> str:
        """Fallback location detection using device name patterns"""
        device_name = port_info.device_name.upper()
        
        if any(pattern in device_name for pattern in ["USB", "FT", "CH34", "CP21"]):
            return "USB"
        elif any(pattern in device_name for pattern in ["PCI", "UART"]):
            return "Internal"
        elif "MOXA" in device_name:
            return "Network"
        else:
            return "Unknown"
    
    def _check_port_status(self, port_info: SerialPortInfo) -> PortStatus:
        """Check current port status"""
        if not SERIAL_AVAILABLE:
            return PortStatus.UNKNOWN
        
        try:
            # Try to open the port briefly to check availability
            test_port = serial.Serial()
            test_port.port = port_info.port_name
            test_port.timeout = 0.1
            
            try:
                test_port.open()
                test_port.close()
                return PortStatus.AVAILABLE
            except serial.SerialException as e:
                error_msg = str(e).lower()
                if "access is denied" in error_msg or "permission denied" in error_msg:
                    return PortStatus.IN_USE
                elif "could not open port" in error_msg:
                    if "moxa" in port_info.manufacturer.lower():
                        return PortStatus.ERROR  # Network issue
                    else:
                        return PortStatus.ERROR  # Device disconnected
                else:
                    return PortStatus.BUSY
        except Exception:
            return PortStatus.UNKNOWN
    
    def _get_hardware_id(self, port_info: SerialPortInfo) -> str:
        """Get hardware ID for the port from cache"""
        if not self._wmi_ports_cache:
            return ""
        
        try:
            if port_info.port_name in self._wmi_ports_cache:
                port = self._wmi_ports_cache[port_info.port_name]
                if hasattr(port, 'PNPDeviceID') and port.PNPDeviceID:
                    return port.PNPDeviceID
        except Exception:
            pass
        
        return ""


class SettingsManager:
    """Manages application settings using QSettings for cross-platform persistence"""
    
    def __init__(self):
        self.settings = QSettings("SerialSplit", "Hub4com")
    
    def get_show_launch_dialog(self):
        """Get whether to show launch dialog on startup (default: True)"""
        return self.settings.value("ui/show_launch_dialog", True, type=bool)
    
    def set_show_launch_dialog(self, show_dialog):
        """Set whether to show launch dialog on startup"""
        self.settings.setValue("ui/show_launch_dialog", show_dialog)
        self.settings.sync()


class DefaultConfig:
    """Default COM pairs and settings to create on application launch"""
    # Default pairs to create: CNCA31<->CNCB31 (COM131<->COM132) and CNCA41<->CNCB41 (COM141<->COM142)
    default_pairs = [
        {"port_a": "CNCA31", "port_b": "CNCB31", "com_a": "COM131", "com_b": "COM132"},
        {"port_a": "CNCA41", "port_b": "CNCB41", "com_a": "COM141", "com_b": "COM142"}
    ]
    default_baud = "115200"
    # Settings for each port in the pair
    default_settings = {
        "EmuBR": "yes",        # Baud rate timing emulation
        "EmuOverrun": "yes"    # Buffer overrun emulation
    }
    # Output port mapping for GUI pre-population
    output_mapping = [
        {"port": "COM131", "baud": "115200"},
        {"port": "COM141", "baud": "115200"}
    ]


class ResponsiveWindowManager:
    """Manages responsive window sizing and layout decisions"""
    
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
        """Get primary screen geometry information"""
        screen = QApplication.primaryScreen()
        if not screen:
            # Fallback if no screen detected
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
        """Determine if screen should be considered small"""
        return (screen_width < cls.SMALL_SCREEN_WIDTH_THRESHOLD or 
                screen_height < cls.SMALL_SCREEN_HEIGHT_THRESHOLD)
    
    @classmethod
    def calculate_main_window_config(cls) -> WindowConfig:
        """Calculate optimal window configuration for main application window"""
        screen_width, screen_height, screen_x, screen_y = cls.get_screen_info()
        is_small = cls.is_small_screen(screen_width, screen_height)
        
        if is_small:
            # Small screen: use most of available space with minimum constraints
            window_width = min(
                max(screen_width * cls.SMALL_SCREEN_WIDTH_RATIO, cls.ABSOLUTE_MIN_WIDTH),
                screen_width
            )
            window_height = min(
                max(screen_height * cls.SMALL_SCREEN_HEIGHT_RATIO, cls.ABSOLUTE_MIN_HEIGHT),
                screen_height
            )
            
            # Center on screen
            x = screen_x + (screen_width - window_width) // 2
            y = screen_y + (screen_height - window_height) // 2
        else:
            # Large screen: use comfortable default size
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
    def calculate_dialog_config(cls, preferred_width: int = 800, preferred_height: int = 500) -> WindowConfig:
        """Calculate optimal window configuration for dialog windows"""
        screen_width, screen_height, screen_x, screen_y = cls.get_screen_info()
        is_small = cls.is_small_screen(screen_width, screen_height)
        
        if is_small:
            # Small screen: use most of available space
            window_width = min(screen_width * 0.9, preferred_width)
            window_height = min(screen_height * 0.8, preferred_height)
            min_width = 600
            min_height = 400
        else:
            # Large screen: use preferred size
            window_width = preferred_width
            window_height = preferred_height
            min_width = preferred_width // 2
            min_height = preferred_height // 2
        
        # Center the dialog
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
        """Get adaptive font size based on screen size"""
        if is_small_screen:
            return max(base_size - 2, 10)  # Reduce by 2, minimum 10
        return base_size
    
    @classmethod
    def get_adaptive_button_size(cls, is_small_screen: bool) -> tuple:
        """Get adaptive button dimensions (width, height)"""
        if is_small_screen:
            return (60, 30)
        return (70, None)  # None means no height restriction
    
    @classmethod
    def get_adaptive_text_height(cls, base_height: int, is_small_screen: bool) -> dict:
        """Get adaptive text widget height configuration"""
        if is_small_screen:
            return {
                'min_height': max(base_height - 50, 100),
                'max_height': base_height
            }
        return {
            'min_height': base_height,
            'max_height': None
        }


class PortScanner(QThread):
    """Thread for scanning Windows registry for serial ports with progressive loading capability"""
    scan_completed = pyqtSignal(list)
    scan_progress = pyqtSignal(str)
    
    # Progressive loading signals
    port_basic_data = pyqtSignal(int, object)  # row, SerialPortInfo with basic data
    port_enhanced_data = pyqtSignal(int, object)  # row, SerialPortInfo with enhanced data  
    port_status_data = pyqtSignal(int, object)  # row, SerialPortInfo with status/driver data
    scan_phase_changed = pyqtSignal(str)  # Current phase description
    
    def __init__(self, complete_scan=False):
        super().__init__()
        self.complete_scan = complete_scan
        self.capability_analyzer = PortCapabilityAnalyzer(fast_mode=not complete_scan)
    
    def run(self):
        """Progressive port scanning with real-time UI updates"""
        try:
            # Phase 1: Registry scan - immediate basic data
            self.scan_phase_changed.emit("Scanning registry...")
            self.scan_progress.emit("Scanning Windows registry...")
            
            basic_ports = self.scan_registry_ports()
            if not basic_ports:
                self.scan_completed.emit([])
                return
            
            # Emit basic port data immediately (Port, Type columns)
            for row, port in enumerate(basic_ports):
                self.port_basic_data.emit(row, port)
            
            # PRODUCTION FIX: Emit basic ports immediately for main GUI
            # This allows port selection while advanced scanning continues in background
            self.scan_completed.emit(basic_ports)
            
            # Phase 2: Quick enhancement - per-port analysis
            self.scan_phase_changed.emit("Analyzing capabilities...")
            enhanced_ports = []
            
            for row, port in enumerate(basic_ports):
                try:
                    # Quick enhancement (Manufacturer, Location, Capabilities, Parameters)
                    enhanced_port = self.quick_enhance_port(port)
                    enhanced_ports.append(enhanced_port)
                    self.port_enhanced_data.emit(row, enhanced_port)
                    
                    # Update progress
                    progress = f"Analyzing {port.port_name} ({row+1}/{len(basic_ports)})"
                    self.scan_progress.emit(progress)
                    
                except Exception as e:
                    print(f"Warning: Quick enhancement failed for {port.port_name}: {str(e)}")
                    enhanced_ports.append(port)
                    self.port_enhanced_data.emit(row, port)
            
            # Phase 3: Complete analysis - status info (always performed)
            self.scan_phase_changed.emit("Checking port status...")
            final_ports = []
            
            for row, port in enumerate(enhanced_ports):
                try:
                    # Complete analysis (Status column)
                    final_port = self.complete_enhance_port(port)
                    final_ports.append(final_port)
                    self.port_status_data.emit(row, final_port)
                    
                    # Update progress
                    progress = f"Checking {port.port_name} status ({row+1}/{len(enhanced_ports)})"
                    self.scan_progress.emit(progress)
                    
                except Exception as e:
                    print(f"Warning: Status check failed for {port.port_name}: {str(e)}")
                    final_ports.append(port)
                    self.port_status_data.emit(row, port)
            
            # Final enhanced data is available, but main GUI already has basic ports
            # Advanced dialog can listen to port_status_data signals for progressive updates
                
        except Exception as e:
            self.scan_progress.emit(f"Error scanning ports: {str(e)}")
            self.scan_completed.emit([])
    
    def scan_registry_ports(self) -> List[SerialPortInfo]:
        """Scan Windows registry for all serial ports"""
        ports = []
        
        if not WINREG_AVAILABLE:
            raise Exception("Windows registry access not available")
        
        try:
            # Check if winreg is available and working
            if not hasattr(winreg, 'OpenKey'):
                raise ImportError("winreg module not properly available")
                
            # Open the SERIALCOMM registry key
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"HARDWARE\DEVICEMAP\SERIALCOMM")
            
            # Enumerate all values
            i = 0
            while i < 256:  # Reasonable limit to prevent infinite loops
                try:
                    device_name, port_name, _ = winreg.EnumValue(key, i)
                    
                    # Classify the port type
                    port_info = self.classify_port(device_name, port_name)
                    ports.append(port_info)
                    
                    i += 1
                except OSError:
                    # No more values
                    break
                except Exception as e:
                    # Skip this value and continue
                    i += 1
                    continue
            
            winreg.CloseKey(key)
            
        except FileNotFoundError:
            # Registry key doesn't exist - this is normal on some systems
            pass
        except ImportError:
            # winreg not available - fallback
            raise Exception("Windows registry access not available")
        except Exception as e:
            # Other registry errors
            raise Exception(f"Registry scan failed: {str(e)}")
        
        # Sort ports by port name
        ports.sort(key=lambda p: self.port_sort_key(p.port_name))
        return ports
    
    def classify_port(self, device_name: str, port_name: str) -> SerialPortInfo:
        """Classify a port based on its registry device name"""
        
        # Check for Moxa devices (Npdrv pattern)
        if device_name.startswith("Npdrv"):
            moxa_details = self.parse_moxa_device(device_name, port_name)
            return SerialPortInfo(
                port_name=port_name,
                device_name=device_name,
                port_type="Virtual (Moxa)",
                registry_key=device_name,
                description=f"Moxa RealCOM virtual port",
                is_moxa=True,
                moxa_details=moxa_details
            )
        
        # Check for COM0COM devices
        elif "CNCB" in device_name or "CNCA" in device_name:
            return SerialPortInfo(
                port_name=port_name,
                device_name=device_name,
                port_type="Virtual (COM0COM)",
                registry_key=device_name,
                description="COM0COM virtual null-modem pair"
            )
        
        # Check for other virtual port patterns
        elif any(pattern in device_name.lower() for pattern in ["com0com", "virtual", "vspd"]):
            return SerialPortInfo(
                port_name=port_name,
                device_name=device_name,
                port_type="Virtual (Other)",
                registry_key=device_name,
                description="Virtual serial port"
            )
        
        # Physical ports (USB, PCI, etc.)
        else:
            return SerialPortInfo(
                port_name=port_name,
                device_name=device_name,
                port_type="Physical",
                registry_key=device_name,
                description="Physical serial port"
            )
    
    def parse_moxa_device(self, device_name: str, port_name: str) -> Dict:
        """Parse Moxa-specific device information"""
        details = {
            "driver_name": device_name,
            "port_number": port_name.replace("COM", ""),
            "connection_type": "Virtual/Network",
            "recommendations": [
                "Disable CTS handshaking for network serial servers",
                "Check network connectivity to Moxa device",
                "Verify Moxa driver configuration",
                "Consider matching baud rate to source device"
            ]
        }
        return details
    
    def port_sort_key(self, port_name: str) -> tuple:
        """Generate sort key for port names"""
        # Extract number from COM port name for proper sorting
        try:
            if port_name.startswith("COM"):
                num = int(port_name[3:])
                return (0, num)  # COM ports first
            else:
                return (1, port_name)  # Other ports second
        except:
            return (2, port_name)  # Fallback
    
    def enhance_port_information(self, ports: List[SerialPortInfo]) -> List[SerialPortInfo]:
        """Enhance port information with detailed capabilities and status"""
        enhanced_ports = []
        
        for i, port in enumerate(ports):
            try:
                # Update progress for capability analysis
                progress_msg = f"Analyzing {port.port_name} ({i+1}/{len(ports)})"
                self.scan_progress.emit(progress_msg)
                
                # Use the capability analyzer to enhance port information
                enhanced_port = self.capability_analyzer.analyze_port_capabilities(port)
                enhanced_ports.append(enhanced_port)
                
            except Exception as e:
                # If capability analysis fails, add the basic port info
                print(f"Warning: Failed to enhance port {port.port_name}: {str(e)}")
                enhanced_ports.append(port)
        
        return enhanced_ports
    
    def quick_enhance_port(self, port_info: SerialPortInfo) -> SerialPortInfo:
        """Quick enhancement for immediate UI feedback (Phase 2)"""
        try:
            # Quick manufacturer detection
            port_info.manufacturer = self.capability_analyzer._detect_manufacturer(port_info)
            
            # Quick capabilities analysis  
            port_info.capabilities = self.capability_analyzer._analyze_hardware_capabilities(port_info)
            
            # Quick location detection
            port_info.location = self.capability_analyzer._get_connection_topology(port_info)
            
            # Set default status for fast mode
            port_info.status = PortStatus.UNKNOWN
            
            # Skip driver info entirely - not needed for dialog
            port_info.driver_version = ""
            port_info.hardware_id = ""
            
        except Exception as e:
            print(f"Warning: Quick enhancement failed for {port_info.port_name}: {str(e)}")
            
        return port_info
    
    def complete_enhance_port(self, port_info: SerialPortInfo) -> SerialPortInfo:
        """Complete enhancement for detailed data (Phase 3)"""
        try:
            # Initialize WMI cache if needed
            if not self.capability_analyzer._cache_initialized:
                self.capability_analyzer._initialize_wmi_cache()
            
            # Get detailed status information
            port_info.status = self.capability_analyzer._check_port_status(port_info)
            
            # Skip driver info - not needed for simplified dialog
            port_info.driver_version = ""
            port_info.hardware_id = ""
            
        except Exception as e:
            print(f"Warning: Complete enhancement failed for {port_info.port_name}: {str(e)}")
            # Provide fallback values
            port_info.status = PortStatus.UNKNOWN
            port_info.driver_version = ""
            port_info.hardware_id = ""
                
        return port_info



class PortConfig:
    """Configuration for a single port"""
    def __init__(self, port_name="", baud_rate="115200"):
        self.port_name = port_name
        self.baud_rate = baud_rate


class Hub4comProcess(QThread):
    """Thread to run hub4com process"""
    output_received = pyqtSignal(str)
    process_started = pyqtSignal()
    process_stopped = pyqtSignal()
    error_occurred = pyqtSignal(str)
    
    def __init__(self, command):
        super().__init__()
        self.command = command
        self.process = None
        self.should_stop = False
    
    def run(self):
        try:
            # Configure subprocess to hide console window
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            self.process = subprocess.Popen(
                self.command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                startupinfo=startupinfo
            )
            
            # Wait a moment to check if process starts successfully
            import time
            time.sleep(1)
            
            if self.process.poll() is not None:
                # Process terminated immediately
                output, _ = self.process.communicate()
                self.error_occurred.emit(f"hub4com exited immediately:\n{output}")
                return
            
            self.process_started.emit()
            
            # Read output line by line
            while self.process.poll() is None and not self.should_stop:
                try:
                    line = self.process.stdout.readline()
                    if line:
                        self.output_received.emit(line.strip())
                except:
                    break
            
            self.process_stopped.emit()
            
        except FileNotFoundError:
            self.error_occurred.emit(f"Could not find hub4com.exe at: {self.command[0]}")
        except Exception as e:
            self.error_occurred.emit(f"Failed to start hub4com: {str(e)}")
    
    def stop_process(self):
        self.should_stop = True
        if self.process and self.process.poll() is None:
            # Give hub4com 2 seconds to cleanup gracefully
            self.process.terminate()
            try:
                self.process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                self.process.kill()
    
    def cleanup_com_ports(self):
        """Force release all COM ports by restarting com0com service"""
        try:
            # Stop com0com service
            subprocess.run(['net', 'stop', 'com0com'], 
                         capture_output=True, check=False)
            # Start com0com service
            subprocess.run(['net', 'start', 'com0com'], 
                         capture_output=True, check=False)
        except Exception:
            # Ignore service restart errors
            pass


class Com0comProcess(QThread):
    """Thread to execute com0com setupc commands"""
    command_completed = pyqtSignal(bool, str)  # success, output
    command_output = pyqtSignal(str)
    pairs_checked = pyqtSignal(list)  # Emitted with existing pairs list
    
    def __init__(self, command_args, operation_type="command"):
        super().__init__()
        self.setupc_path = r"C:\Program Files (x86)\com0com\setupc.exe"
        self.command_args = command_args
        self.operation_type = operation_type  # "command", "list", "create_default", "check_and_create_default"
        
    def run(self):
        try:
            if self.operation_type == "create_default":
                self._create_default_pairs()
            elif self.operation_type == "check_and_create_default":
                self._check_and_create_default_pairs()
            elif self.operation_type == "list":
                self._list_existing_pairs()
            else:
                self._execute_command()
                
        except subprocess.TimeoutExpired:
            self.command_completed.emit(False, "Command timed out")
        except FileNotFoundError:
            self.command_completed.emit(False, f"setupc.exe not found at {self.setupc_path}")
        except Exception as e:
            self.command_completed.emit(False, f"Error: {str(e)}")
    
    def _execute_command(self):
        """Execute a standard setupc command"""
        cmd = [self.setupc_path] + self.command_args
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        output = result.stdout + result.stderr
        success = result.returncode == 0
        
        self.command_completed.emit(success, output)
    
    def _list_existing_pairs(self):
        """List existing COM0COM pairs"""
        cmd = [self.setupc_path, "list"]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode == 0:
            existing_pairs = self._parse_pairs_output(result.stdout)
            self.pairs_checked.emit(existing_pairs)
        else:
            self.pairs_checked.emit([])
    
    def _create_default_pairs(self):
        """Create default COM pairs if they don't exist"""
        # First, list existing pairs
        list_cmd = [self.setupc_path, "list"]
        list_result = subprocess.run(
            list_cmd,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        existing_pairs = []
        if list_result.returncode == 0:
            existing_pairs = self._parse_pairs_output(list_result.stdout)
        
        # Check which default pairs need to be created
        default_config = DefaultConfig()
        created_pairs = []
        
        for pair_config in default_config.default_pairs:
            port_a, port_b = pair_config["port_a"], pair_config["port_b"]
            com_a, com_b = pair_config["com_a"], pair_config["com_b"]
            
            # Check if this pair already exists
            pair_exists = any(
                (p.get("port_a") == port_a and p.get("port_b") == port_b) or
                (p.get("com_a") == com_a and p.get("com_b") == com_b)
                for p in existing_pairs
            )
            
            if not pair_exists:
                # Create the pair with specific settings
                create_cmd = [
                    self.setupc_path, "install",
                    f"PortName={com_a},EmuBR=yes,EmuOverrun=yes,AllDataBits=yes,AddRTTO=100,AddRITO=100",
                    f"PortName={com_b},EmuBR=yes,EmuOverrun=yes,AllDataBits=yes,AddRTTO=100,AddRITO=100"
                ]
                
                create_result = subprocess.run(
                    create_cmd,
                    capture_output=True,
                    text=True,
                    timeout=45
                )
                
                if create_result.returncode == 0:
                    created_pairs.append(f"{com_a}<->{com_b}")
        
        if created_pairs:
            success_msg = f"Successfully created virtual COM port pairs: {', '.join(created_pairs)} with baud rate timing and buffer overrun protection enabled"
            self.command_completed.emit(True, success_msg)
        else:
            self.command_completed.emit(True, "Virtual COM port pairs are already configured and ready for marine operations")
    
    def _parse_pairs_output(self, output: str) -> List[Dict]:
        """Parse setupc list output to extract existing pairs"""
        pairs = []
        lines = output.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if 'PortName=' in line:
                # Extract port information from setupc output
                # This is a simplified parser - may need refinement based on actual output format
                if 'COM' in line:
                    # Try to extract COM port number
                    import re
                    com_match = re.search(r'COM(\d+)', line)
                    if com_match:
                        com_num = com_match.group(1)
                        pairs.append({
                            "com_a": f"COM{com_num}",
                            "com_b": "",  # Would need to parse paired port
                            "port_a": "",
                            "port_b": "",
                            "raw_line": line
                        })
        
        return pairs
    
    def _parse_com0com_output(self, output: str) -> Dict:
        """Parse com0com list output to extract existing pairs"""
        pairs = {}
        lines = output.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if line and not line.startswith('command>'):
                parts = line.split(None, 1)
                if len(parts) >= 1:
                    port = parts[0]
                    params = parts[1] if len(parts) > 1 else ""
                    
                    if port.startswith('CNCA'):
                        pair_num = port[4:]  # Extract number after "CNCA"
                        if pair_num not in pairs:
                            pairs[pair_num] = {}
                        pairs[pair_num]['A'] = (port, params)
                    elif port.startswith('CNCB'):
                        pair_num = port[4:]  # Extract number after "CNCB"
                        if pair_num not in pairs:
                            pairs[pair_num] = {}
                        pairs[pair_num]['B'] = (port, params)
        
        return pairs
    
    def _extract_actual_port_name(self, virtual_name: str, params: str) -> str:
        """Extract the actual COM port name from parameters"""
        if not params:
            return virtual_name
        
        if "RealPortName=" in params:
            real_name = params.split("RealPortName=")[1].split(",")[0]
            if real_name and real_name != "-":
                return real_name
        
        if "PortName=" in params:
            port_name = params.split("PortName=")[1].split(",")[0]
            if port_name and port_name not in ["-", "COM#"]:
                return port_name
        
        return virtual_name
    
    def _check_and_create_default_pairs(self):
        """Check which default pairs exist using setupc.exe list and only create missing ones"""
        try:
            # First, get existing pairs using setupc.exe list command
            list_cmd = [self.setupc_path, "list"]
            list_result = subprocess.run(
                list_cmd,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            existing_pairs_dict = {}
            existing_com_ports = set()
            
            if list_result.returncode == 0:
                # Parse the existing pairs
                parsed_pairs = self._parse_com0com_output(list_result.stdout)
                
                # Extract COM port names from existing pairs
                for pair_num, pair_data in parsed_pairs.items():
                    if 'A' in pair_data and 'B' in pair_data:
                        port_a, params_a = pair_data['A']
                        port_b, params_b = pair_data['B']
                        
                        # Get actual COM port names
                        com_a = self._extract_actual_port_name(port_a, params_a)
                        com_b = self._extract_actual_port_name(port_b, params_b)
                        
                        existing_com_ports.add(com_a)
                        existing_com_ports.add(com_b)
                        existing_pairs_dict[f"{com_a}<->{com_b}"] = True
            
            # Check which default pairs need to be created
            default_config = DefaultConfig()
            created_pairs = []
            existing_pairs = []
            
            for pair_config in default_config.default_pairs:
                com_a, com_b = pair_config["com_a"], pair_config["com_b"]
                pair_key = f"{com_a}<->{com_b}"
                
                # Check if both COM ports exist
                pair_exists = com_a in existing_com_ports and com_b in existing_com_ports
                
                if pair_exists:
                    existing_pairs.append(pair_key)
                else:
                    # Create the missing pair
                    create_cmd = [
                        self.setupc_path, "install",
                        f"PortName={com_a},EmuBR=yes,EmuOverrun=yes,AllDataBits=yes,AddRTTO=100,AddRITO=100",
                        f"PortName={com_b},EmuBR=yes,EmuOverrun=yes,AllDataBits=yes,AddRTTO=100,AddRITO=100"
                    ]
                    
                    create_result = subprocess.run(
                        create_cmd,
                        capture_output=True,
                        text=True,
                        timeout=45
                    )
                    
                    if create_result.returncode == 0:
                        created_pairs.append(pair_key)
            
            # Build status message
            messages = []
            if existing_pairs:
                messages.append(f"Found existing virtual COM port pairs: {', '.join(existing_pairs)}")
            if created_pairs:
                messages.append(f"Successfully created new virtual COM port pairs: {', '.join(created_pairs)} with baud rate timing and buffer overrun protection enabled")
            
            if messages:
                final_message = ". ".join(messages) + ". All virtual COM port pairs are now ready for marine operations."
            else:
                final_message = "Virtual COM port configuration completed successfully."
            
            self.command_completed.emit(True, final_message)
            
        except Exception as e:
            # Fallback to original behavior if detection fails
            self._create_default_pairs()


# ============================================================================
# SERIAL PORT MONITORING
# ============================================================================

class SerialPortMonitor(QThread):
    """
    Serial port monitoring class for real-time statistics and data flow observation.
    """
    # Signals
    stats_updated = pyqtSignal(dict)  # Emits updated port statistics
    data_received = pyqtSignal(bytes)  # Emits raw data received
    error_occurred = pyqtSignal(str)  # Emits error messages
    
    def __init__(self, port_name, baudrate=9600):
        """
        Initialize the serial port monitor.
        
        Args:
            port_name: Serial port name
            baudrate: Baud rate for monitoring
        """
        super().__init__()
        
        self.port_name = port_name
        self.baudrate = baudrate
        
        # Basic statistics
        self.stats = {
            "rx_bytes": 0,
            "tx_bytes": 0,
            "rx_rate": 0.0,  # bytes per second
            "tx_rate": 0.0,  # bytes per second
            "errors": 0,
            "start_time": None,
            "running_time": 0.0,
            "is_monitoring": False
        }
        
        # Advanced statistics
        self.advanced_stats = AdvancedStatistics()
        
        # Rate calculation windows
        self.rx_window = []  # List of (timestamp, bytes) tuples
        self.tx_window = []  # List of (timestamp, bytes) tuples
        self.window_size = 2  # seconds for rate calculation
        
        # Packet detection
        self.rx_packets = []  # List of SerialPacketInfo
        self.tx_packets = []  # List of SerialPacketInfo
        self.last_rx_timestamp = 0.0
        self.last_tx_timestamp = 0.0
        self.packet_gap_threshold = 0.01  # 10ms gap indicates new packet
        self.current_rx_buffer = bytearray()
        self.current_tx_buffer = bytearray()
        
        # Operation flags
        self.monitoring = False
        self.ser = None
        
        # Thread safety for TX operations
        self.tx_mutex = QMutex()
        
    def start_monitoring(self):
        """Start monitoring the serial port."""
        if not SERIAL_AVAILABLE:
            self.error_occurred.emit("pyserial module not available for port monitoring")
            return False
            
        if self.monitoring:
            return True
            
        try:
            # Reset stats
            self.stats = {
                "rx_bytes": 0,
                "tx_bytes": 0,
                "rx_rate": 0.0,
                "tx_rate": 0.0,
                "errors": 0,
                "start_time": datetime.now(),
                "running_time": 0.0,
                "is_monitoring": True
            }
            
            self.rx_window = []
            self.tx_window = []
            
            # Start the monitor thread
            self.monitoring = True
            self.start()
            
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"Error starting monitor for {self.port_name}: {str(e)}")
            self.monitoring = False
            return False
    
    def stop_monitoring(self):
        """Stop monitoring the serial port."""
        if not self.monitoring:
            return
            
        self.monitoring = False
        self.stats["is_monitoring"] = False
        
        # Wait for thread to exit
        if self.isRunning():
            self.wait(1000)
        
        # Close the port if open
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
            except:
                pass
    
    def run(self):
        """Main monitoring loop running in the thread."""
        last_stats_update = time.time()
        
        # Try to open port for monitoring (non-blocking)
        try:
            if SERIAL_AVAILABLE:
                self.ser = serial.Serial()
                self.ser.port = self.port_name
                self.ser.baudrate = self.baudrate
                self.ser.timeout = 0.1
                
                # Try to open the port with enhanced error handling
                try:
                    self.ser.open()
                except serial.SerialException as e:
                    # Provide specific error feedback for different scenarios
                    error_msg = str(e).lower()
                    if "access is denied" in error_msg or "permission denied" in error_msg:
                        self.error_occurred.emit(f"Port {self.port_name} is busy or in use by another application")
                    elif "could not open port" in error_msg:
                        if "moxa" in self.port_name.lower() or "mxser" in error_msg:
                            self.error_occurred.emit(f"Moxa port {self.port_name} network connection unavailable")
                        else:
                            self.error_occurred.emit(f"Port {self.port_name} could not be opened - device may be disconnected")
                    else:
                        self.error_occurred.emit(f"Port {self.port_name} monitoring unavailable: {str(e)}")
                    self.ser = None
                except Exception as e:
                    self.error_occurred.emit(f"Unexpected error opening {self.port_name}: {str(e)}")
                    self.ser = None
                    
        except Exception as e:
            self.error_occurred.emit(f"Failed to initialize monitoring for {self.port_name}: {str(e)}")
            self.ser = None
        
        while self.monitoring:
            try:
                # If we have an open serial port, monitor it
                if self.ser and self.ser.is_open:
                    # Check for incoming data
                    if self.ser.in_waiting > 0:
                        data = self.ser.read(self.ser.in_waiting)
                        if data:
                            # Update statistics
                            self.stats["rx_bytes"] += len(data)
                            now = time.time()
                            self.rx_window.append((now, len(data)))
                            
                            # Process for advanced statistics
                            self._process_rx_data(data, now)
                            
                            # Emit the data
                            self.data_received.emit(data)
                
                # Update running time and rates periodically
                now = time.time()
                if now - last_stats_update >= 1.0:  # Update stats every second
                    self._update_rates(now)
                    if self.stats["start_time"]:
                        self.stats["running_time"] = (datetime.now() - self.stats["start_time"]).total_seconds()
                    
                    # Finalize any pending packets based on timeout
                    if self.current_rx_buffer and (now - self.last_rx_timestamp) > self.packet_gap_threshold:
                        self._finalize_rx_packet(self.last_rx_timestamp, now - self.last_rx_timestamp)
                    if self.current_tx_buffer and (now - self.last_tx_timestamp) > self.packet_gap_threshold:
                        self._finalize_tx_packet(self.last_tx_timestamp, now - self.last_tx_timestamp)
                    
                    # Emit updated stats
                    self.stats_updated.emit(self.stats.copy())
                    last_stats_update = now
                
                # Short sleep to prevent CPU thrashing
                time.sleep(0.1)
                
            except Exception as e:
                self.stats["errors"] += 1
                self._handle_serial_error(e)
                self.error_occurred.emit(f"Monitor error: {str(e)}")
                time.sleep(0.5)  # Wait before retrying
        
        # Ensure port is closed on exit
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
            except:
                pass
    
    def _update_rates(self, now):
        """
        Update RX and TX rates based on windowed data.
        
        Args:
            now: Current timestamp
        """
        # Remove old data points outside the window
        window_start = now - self.window_size
        self.rx_window = [(ts, sz) for ts, sz in self.rx_window if ts >= window_start]
        self.tx_window = [(ts, sz) for ts, sz in self.tx_window if ts >= window_start]
        
        # Calculate rates
        if self.rx_window:
            total_rx_bytes = sum(sz for _, sz in self.rx_window)
            oldest_ts = min(ts for ts, _ in self.rx_window)
            if now > oldest_ts:
                time_span = now - oldest_ts
                self.stats["rx_rate"] = total_rx_bytes / time_span
            else:
                self.stats["rx_rate"] = 0.0
        else:
            self.stats["rx_rate"] = 0.0
        
        if self.tx_window:
            total_tx_bytes = sum(sz for _, sz in self.tx_window)
            oldest_ts = min(ts for ts, _ in self.tx_window)
            if now > oldest_ts:
                time_span = now - oldest_ts
                self.stats["tx_rate"] = total_tx_bytes / time_span
            else:
                self.stats["tx_rate"] = 0.0
        else:
            self.stats["tx_rate"] = 0.0
    
    def get_formatted_stats(self):
        """
        Get formatted statistics as a string.
        
        Returns:
            str: Formatted statistics string
        """
        if not self.stats["start_time"] or not self.stats["is_monitoring"]:
            return "Not monitoring"
            
        # Format rates
        rx_rate = self.stats["rx_rate"]
        tx_rate = self.stats["tx_rate"]
        
        # Choose appropriate units
        if rx_rate < 1024:
            rx_rate_str = f"{rx_rate:.1f} B/s"
        else:
            rx_rate_str = f"{rx_rate/1024:.1f} KB/s"
            
        if tx_rate < 1024:
            tx_rate_str = f"{tx_rate:.1f} B/s"
        else:
            tx_rate_str = f"{tx_rate/1024:.1f} KB/s"
            
        # Format running time
        seconds = int(self.stats["running_time"])
        running_time = str(timedelta(seconds=seconds))
        
        # Format the statistics string
        stats_str = f"Monitoring: {running_time}\n"
        stats_str += f"RX: {self.stats['rx_bytes']} bytes ({rx_rate_str})\n"
        stats_str += f"TX: {self.stats['tx_bytes']} bytes ({tx_rate_str})\n"
        stats_str += f"Errors: {self.stats['errors']}"
        
        return stats_str
    
    def send_data(self, data):
        """
        Send data through the monitored serial port (thread-safe).
        
        Args:
            data: bytes or str to send
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not self.monitoring or not self.ser or not self.ser.is_open:
            return False
            
        try:
            # Convert string to bytes if needed
            if isinstance(data, str):
                data = data.encode('utf-8')
            
            # Thread-safe TX operation
            self.tx_mutex.lock()
            try:
                bytes_written = self.ser.write(data)
                self.ser.flush()  # Ensure data is sent immediately
                
                # Update TX statistics
                if bytes_written > 0:
                    self.stats["tx_bytes"] += bytes_written
                    now = time.time()
                    self.tx_window.append((now, bytes_written))
                    
                    # Process for advanced statistics
                    self._process_tx_data(data, now)
                
                return True
                
            finally:
                self.tx_mutex.unlock()
                
        except Exception as e:
            self.stats["errors"] += 1
            self.error_occurred.emit(f"TX error: {str(e)}")
            return False
    
    def _process_rx_data(self, data, timestamp):
        """Process received data for packet detection and statistics"""
        self.current_rx_buffer.extend(data)
        
        # Check for packet boundary (gap-based detection)
        if self.last_rx_timestamp > 0:
            gap = timestamp - self.last_rx_timestamp
            if gap > self.packet_gap_threshold and len(self.current_rx_buffer) > 0:
                # Previous buffer was a complete packet
                self._finalize_rx_packet(self.last_rx_timestamp, gap)
        
        self.last_rx_timestamp = timestamp
        
        # Update peak rates
        if self.stats["rx_rate"] > self.advanced_stats.rx_peak_rate:
            self.advanced_stats.rx_peak_rate = self.stats["rx_rate"]
    
    def _process_tx_data(self, data, timestamp):
        """Process transmitted data for packet detection and statistics"""
        self.current_tx_buffer.extend(data)
        
        # Check for packet boundary (gap-based detection)
        if self.last_tx_timestamp > 0:
            gap = timestamp - self.last_tx_timestamp
            if gap > self.packet_gap_threshold and len(self.current_tx_buffer) > 0:
                # Previous buffer was a complete packet
                self._finalize_tx_packet(self.last_tx_timestamp, gap)
        
        self.last_tx_timestamp = timestamp
        
        # Update peak rates
        if self.stats["tx_rate"] > self.advanced_stats.tx_peak_rate:
            self.advanced_stats.tx_peak_rate = self.stats["tx_rate"]
    
    def _finalize_rx_packet(self, timestamp, inter_frame_gap):
        """Finalize an RX packet and update statistics"""
        if len(self.current_rx_buffer) == 0:
            return
        
        packet = SerialPacketInfo(
            timestamp=timestamp,
            size=len(self.current_rx_buffer),
            direction="RX",
            data=bytes(self.current_rx_buffer),
            inter_frame_gap=inter_frame_gap
        )
        
        self.rx_packets.append(packet)
        self.advanced_stats.rx_packet_count += 1
        self.advanced_stats.rx_packet_sizes.append(packet.size)
        
        # Update inter-frame gap statistics
        if inter_frame_gap < self.advanced_stats.min_inter_frame_gap:
            self.advanced_stats.min_inter_frame_gap = inter_frame_gap
        if inter_frame_gap > self.advanced_stats.max_inter_frame_gap:
            self.advanced_stats.max_inter_frame_gap = inter_frame_gap
        
        # Update running average
        total_gaps = sum(p.inter_frame_gap for p in self.rx_packets if p.inter_frame_gap > 0)
        gap_count = len([p for p in self.rx_packets if p.inter_frame_gap > 0])
        if gap_count > 0:
            self.advanced_stats.average_inter_frame_gap = total_gaps / gap_count
        
        # Clear buffer for next packet
        self.current_rx_buffer.clear()
        
        # Keep only recent packets (last 1000)
        if len(self.rx_packets) > 1000:
            self.rx_packets.pop(0)
    
    def _finalize_tx_packet(self, timestamp, inter_frame_gap):
        """Finalize a TX packet and update statistics"""
        if len(self.current_tx_buffer) == 0:
            return
        
        packet = SerialPacketInfo(
            timestamp=timestamp,
            size=len(self.current_tx_buffer),
            direction="TX",
            data=bytes(self.current_tx_buffer),
            inter_frame_gap=inter_frame_gap
        )
        
        self.tx_packets.append(packet)
        self.advanced_stats.tx_packet_count += 1
        self.advanced_stats.tx_packet_sizes.append(packet.size)
        
        # Clear buffer for next packet
        self.current_tx_buffer.clear()
        
        # Keep only recent packets (last 1000)
        if len(self.tx_packets) > 1000:
            self.tx_packets.pop(0)
    
    def _update_advanced_stats(self):
        """Update advanced statistics calculations"""
        # Calculate average rates
        if self.rx_window:
            total_rx = sum(size for _, size in self.rx_window)
            time_span = self.rx_window[-1][0] - self.rx_window[0][0] if len(self.rx_window) > 1 else 1
            self.advanced_stats.rx_average_rate = total_rx / time_span if time_span > 0 else 0
        
        if self.tx_window:
            total_tx = sum(size for _, size in self.tx_window)
            time_span = self.tx_window[-1][0] - self.tx_window[0][0] if len(self.tx_window) > 1 else 1
            self.advanced_stats.tx_average_rate = total_tx / time_span if time_span > 0 else 0
    
    def get_advanced_stats(self):
        """Get current advanced statistics"""
        self._update_advanced_stats()
        return self.advanced_stats
    
    def _handle_serial_error(self, error):
        """Handle and classify serial errors"""
        error_str = str(error).lower()
        
        if "timeout" in error_str:
            self.advanced_stats.timeout_errors += 1
        elif "overrun" in error_str or "buffer" in error_str:
            self.advanced_stats.buffer_overruns += 1
        elif "framing" in error_str:
            self.advanced_stats.framing_errors += 1
        elif "parity" in error_str:
            self.advanced_stats.parity_errors += 1
        else:
            # Generic error
            self.advanced_stats.rx_errors += 1 if "rx" in error_str or "read" in error_str else 0
            self.advanced_stats.tx_errors += 1 if "tx" in error_str or "write" in error_str else 0


class SerialPortTester:
    """
    Serial port testing functionality for parameter detection and diagnostics.
    """
    
    def __init__(self):
        """Initialize the port tester."""
        self.available = SERIAL_AVAILABLE
    
    def test_port(self, port_name: str) -> Dict:
        """
        Test a serial port and return comprehensive information.
        
        Args:
            port_name (str): Name of the serial port to test
            
        Returns:
            Dict: Dictionary containing test results and port information
        """
        if not self.available:
            return {
                "status": "Error",
                "message": "pyserial module not available",
                "details": {}
            }
        
        try:
            # Try to open the port with minimal configuration
            ser = serial.Serial(port_name, timeout=1)
            
            # Collect basic port information
            port_info = {
                "port": port_name,
                "bytesize": ser.bytesize,
                "parity": ser.parity,
                "stopbits": ser.stopbits,
                "timeout": ser.timeout,
                "xonxoff": ser.xonxoff,
                "rtscts": ser.rtscts,
                "dsrdtr": ser.dsrdtr,
                "write_timeout": getattr(ser, 'write_timeout', 'N/A'),
                "inter_byte_timeout": getattr(ser, 'inter_byte_timeout', 'N/A')
            }
            
            # Get modem status if available
            try:
                modem_status = {
                    "CTS": ser.cts if hasattr(ser, 'cts') else 'N/A',
                    "DSR": ser.dsr if hasattr(ser, 'dsr') else 'N/A',
                    "RI": ser.ri if hasattr(ser, 'ri') else 'N/A',
                    "CD": ser.cd if hasattr(ser, 'cd') else 'N/A'
                }
                port_info["modem_status"] = modem_status
            except Exception:
                port_info["modem_status"] = "Not available"
            
            # Get buffer information if available
            try:
                port_info["in_waiting"] = ser.in_waiting
                port_info["out_waiting"] = ser.out_waiting if hasattr(ser, 'out_waiting') else 'N/A'
            except Exception:
                port_info["in_waiting"] = 'N/A'
                port_info["out_waiting"] = 'N/A'
            
            # Close the port
            ser.close()
            
            return {
                "status": "Available",
                "message": f"Port {port_name} is available and ready",
                "details": port_info
            }
            
        except serial.SerialException as e:
            error_msg = str(e).lower()
            if "access is denied" in error_msg or "permission denied" in error_msg:
                status_msg = f"Port {port_name} is in use by another application"
            elif "file not found" in error_msg or "system cannot find" in error_msg:
                status_msg = f"Port {port_name} does not exist or is not available"
            else:
                status_msg = f"Port {port_name} has an error: {str(e)}"
            
            return {
                "status": "Error",
                "message": status_msg,
                "details": {"error": str(e)}
            }
        except Exception as e:
            return {
                "status": "Error", 
                "message": f"Unexpected error testing port {port_name}",
                "details": {"error": str(e)}
            }
    
    def format_test_results(self, test_results: Dict) -> str:
        """
        Format test results into a readable string.
        
        Args:
            test_results (Dict): Results from test_port()
            
        Returns:
            str: Formatted test results
        """
        if test_results["status"] == "Error":
            return f"{test_results['message']}\n\nError Details:\n{test_results['details'].get('error', 'Unknown error')}"
        
        details = test_results["details"]
        result = f"{test_results['message']}\n\n"
        
        # Basic port configuration
        result += "Port Configuration:\n"
        result += f"Data Bits: {details.get('bytesize', 'N/A')}\n"
        result += f"Parity: {details.get('parity', 'N/A')}\n"
        result += f"Stop Bits: {details.get('stopbits', 'N/A')}\n"
        result += f"Timeout: {details.get('timeout', 'N/A')}s\n\n"
        
        # Flow control
        result += "Flow Control:\n"
        result += f"XON/XOFF: {details.get('xonxoff', 'N/A')}\n"
        result += f"RTS/CTS: {details.get('rtscts', 'N/A')}\n"
        result += f"DSR/DTR: {details.get('dsrdtr', 'N/A')}\n\n"
        
        # Modem status
        if "modem_status" in details and isinstance(details["modem_status"], dict):
            result += "Modem Status:\n"
            for signal, value in details["modem_status"].items():
                result += f"  {signal}: {value}\n"
            result += "\n"
        
        # Buffer status
        if "in_waiting" in details:
            result += "Buffer Status:\n"
            result += f"Input Buffer: {details['in_waiting']} bytes\n"
            if details.get("out_waiting") != 'N/A':
                result += f"  Output Buffer: {details['out_waiting']} bytes\n"
            result += "\n"
        
        # Additional timeouts
        if details.get("write_timeout") != 'N/A' or details.get("inter_byte_timeout") != 'N/A':
            result += "Advanced Timeouts:\n"
            if details.get("write_timeout") != 'N/A':
                result += f"  Write Timeout: {details['write_timeout']}s\n"
            if details.get("inter_byte_timeout") != 'N/A':
                result += f"  Inter-byte Timeout: {details['inter_byte_timeout']}s\n"
        
        return result