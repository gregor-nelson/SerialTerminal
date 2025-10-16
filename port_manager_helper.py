#!/usr/bin/env python3
"""
Elevated Helper for Virtual Port Manager
This script runs with administrator privileges to execute com0com setupc.exe commands.
It communicates via command-line arguments and outputs JSON results to stdout.

Designed for robustness:
- Comprehensive error handling
- JSON-based communication
- Timeout protection
- Detailed error messages
"""

import sys
import os
import subprocess
import json
import time
import shlex
from typing import Dict, Any


class CommandResult:
    """Result of a setupc.exe command execution."""

    def __init__(self, success: bool, output: str = "", error: str = "",
                 return_code: int = 0, execution_time: float = 0.0, command: str = ""):
        self.success = success
        self.output = output
        self.error = error
        self.return_code = return_code
        self.execution_time = execution_time
        self.command = command

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            "success": self.success,
            "output": self.output,
            "error": self.error,
            "return_code": self.return_code,
            "execution_time": self.execution_time,
            "command": self.command
        }


def execute_setupc_command(setupc_path: str, command: str, working_dir: str = None, timeout: int = 30) -> CommandResult:
    """
    Execute a setupc.exe command with elevated privileges.

    Args:
        setupc_path: Full path to setupc.exe
        command: The command arguments (e.g., "list", "install ...")
        working_dir: Working directory for command execution
        timeout: Command timeout in seconds

    Returns:
        CommandResult with execution details
    """
    start_time = time.time()
    full_command = f'"{setupc_path}" {command}'

    try:
        # Validate setupc_path exists
        if not os.path.exists(setupc_path):
            return CommandResult(
                success=False,
                output="",
                error=f"setupc.exe not found at: {setupc_path}",
                return_code=-2,
                execution_time=time.time() - start_time,
                command=full_command
            )

        # Parse the command properly
        try:
            command_args = shlex.split(full_command)
        except ValueError as e:
            return CommandResult(
                success=False,
                output="",
                error=f"Invalid command format: {str(e)}",
                return_code=-4,
                execution_time=time.time() - start_time,
                command=full_command
            )

        # Execute the command
        result = subprocess.run(
            command_args,
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=False,
            cwd=working_dir
        )

        execution_time = time.time() - start_time

        return CommandResult(
            success=result.returncode == 0,
            output=result.stdout,
            error=result.stderr,
            return_code=result.returncode,
            execution_time=execution_time,
            command=full_command
        )

    except subprocess.TimeoutExpired:
        execution_time = time.time() - start_time
        return CommandResult(
            success=False,
            output="",
            error=f"Command timed out after {timeout} seconds",
            return_code=-1,
            execution_time=execution_time,
            command=full_command
        )

    except FileNotFoundError:
        execution_time = time.time() - start_time
        return CommandResult(
            success=False,
            output="",
            error=f"setupc.exe not found at: {setupc_path}",
            return_code=-2,
            execution_time=execution_time,
            command=full_command
        )

    except Exception as e:
        execution_time = time.time() - start_time
        return CommandResult(
            success=False,
            output="",
            error=f"Unexpected error: {str(e)}",
            return_code=-3,
            execution_time=execution_time,
            command=full_command
        )


def main():
    """Main entry point for the elevated helper."""

    output_file = None

    try:
        # Expected arguments:
        # python port_manager_helper.py <setupc_path> <command> [timeout] [--output-file <path>]

        if len(sys.argv) < 3:
            result = CommandResult(
                success=False,
                error="Usage: port_manager_helper.py <setupc_path> <command> [timeout] [--output-file <path>]",
                return_code=-10
            )
            output_json = json.dumps(result.to_dict())
            print(output_json, flush=True)
            sys.exit(1)

        # Parse arguments
        setupc_path = sys.argv[1]
        command = sys.argv[2]

        # Check for optional arguments
        timeout = 30
        arg_idx = 3
        while arg_idx < len(sys.argv):
            if sys.argv[arg_idx] == '--output-file' and arg_idx + 1 < len(sys.argv):
                output_file = sys.argv[arg_idx + 1]
                arg_idx += 2
            else:
                # Assume it's the timeout
                try:
                    timeout = int(sys.argv[arg_idx])
                except ValueError:
                    pass
                arg_idx += 1

        # Validate timeout
        if timeout <= 0 or timeout > 300:
            result = CommandResult(
                success=False,
                error=f"Invalid timeout value: {timeout} (must be 1-300 seconds)",
                return_code=-11
            )
            output_json = json.dumps(result.to_dict())
            if output_file:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(output_json)
            print(output_json, flush=True)
            sys.exit(1)

        # Determine working directory (directory containing setupc.exe)
        working_dir = os.path.dirname(setupc_path) if setupc_path else None

        # Execute the command
        result = execute_setupc_command(setupc_path, command, working_dir, timeout)

        # Output result as JSON
        output_json = json.dumps(result.to_dict())

        # Write to file if specified (for UAC elevation scenarios)
        if output_file:
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(output_json)
            except Exception as e:
                # If file write fails, still output to stdout
                print(f'{{"success": false, "error": "Failed to write output file: {str(e)}", "return_code": -12}}', flush=True)
                sys.exit(1)

        # Also output to stdout (for non-elevated scenarios)
        print(output_json, flush=True)

        # Exit with appropriate code
        sys.exit(0 if result.success else 1)

    except Exception as e:
        # Catastrophic error - output basic error info
        result = CommandResult(
            success=False,
            error=f"Helper crashed: {str(e)}",
            return_code=-99
        )
        output_json = json.dumps(result.to_dict())
        try:
            if output_file:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(output_json)
            print(output_json, flush=True)
        except:
            # Even JSON serialization failed - output raw error
            error_json = f'{{"success": false, "error": "Critical failure: {str(e)}", "return_code": -99}}'
            if output_file:
                try:
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(error_json)
                except:
                    pass
            print(error_json, flush=True)
        sys.exit(99)


if __name__ == "__main__":
    main()
