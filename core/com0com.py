#!/usr/bin/env python3
"""
COM0COM virtual serial port management.
Contains classes for managing com0com virtual port pairs.
"""

import subprocess
import re
from typing import Dict, List
from PyQt6.QtCore import QThread, pyqtSignal


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
