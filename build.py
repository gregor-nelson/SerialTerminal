#!/usr/bin/env python3
"""
Serial Terminal Build Script
Builds both the main application and the elevated port manager helper.
"""

import subprocess
import sys
import shutil
import os
from pathlib import Path
from datetime import datetime


class BuildScript:
    """Handles the build process for Serial Terminal application."""

    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.specs_dir = self.root_dir / "specs"
        self.dist_dir = self.root_dir / "dist"
        self.build_dir = self.root_dir / "build"
        self.specs = [
            ("specs/serial_terminal.spec", "SerialTerminal.exe"),
            ("specs/port_manager_helper.spec", "SerialPortManager.exe")
        ]

    def print_header(self, message):
        """Print a formatted header message."""
        print("\n" + "=" * 70)
        print(f"  {message}")
        print("=" * 70)

    def print_step(self, step_num, total_steps, message):
        """Print a formatted step message."""
        print(f"\n[{step_num}/{total_steps}] {message}")
        print("-" * 70)

    def clean_build_artifacts(self):
        """Remove previous build artifacts."""
        self.print_step(1, 4, "Cleaning previous build artifacts")

        dirs_to_clean = [self.build_dir, self.dist_dir]

        for directory in dirs_to_clean:
            if directory.exists():
                print(f"  Removing: {directory}")
                try:
                    shutil.rmtree(directory)
                    print(f"  ✓ Removed {directory}")
                except Exception as e:
                    print(f"  ⚠ Warning: Could not remove {directory}: {e}")
            else:
                print(f"  • {directory} does not exist (skipping)")

        print("  ✓ Cleanup complete")

    def build_executable(self, spec_file, exe_name, step_num):
        """Build a single executable using PyInstaller."""
        self.print_step(step_num, 4, f"Building {exe_name}")

        spec_path = self.root_dir / spec_file

        if not spec_path.exists():
            print(f"  ✗ ERROR: Spec file not found: {spec_path}")
            return False

        print(f"  Spec file: {spec_file}")
        print(f"  Running PyInstaller...")

        try:
            # Run PyInstaller with the spec file
            result = subprocess.run(
                [sys.executable, "-m", "PyInstaller", str(spec_file)],
                cwd=self.root_dir,
                capture_output=True,
                text=True
            )

            if result.returncode != 0:
                print(f"  ✗ PyInstaller failed with return code {result.returncode}")
                print("\n  STDOUT:")
                print(result.stdout)
                print("\n  STDERR:")
                print(result.stderr)
                return False

            # Check if the executable was created
            exe_path = self.dist_dir / exe_name
            if exe_path.exists():
                file_size = exe_path.stat().st_size / (1024 * 1024)  # Convert to MB
                print(f"  ✓ Built successfully: {exe_name} ({file_size:.2f} MB)")
                return True
            else:
                print(f"  ✗ ERROR: Expected output not found: {exe_path}")
                return False

        except Exception as e:
            print(f"  ✗ ERROR: {e}")
            return False

    def verify_outputs(self):
        """Verify that all expected executables were created."""
        self.print_step(4, 4, "Verifying build outputs")

        all_present = True

        for spec_file, exe_name in self.specs:
            exe_path = self.dist_dir / exe_name
            if exe_path.exists():
                file_size = exe_path.stat().st_size / (1024 * 1024)
                print(f"  ✓ {exe_name} ({file_size:.2f} MB)")
            else:
                print(f"  ✗ Missing: {exe_name}")
                all_present = False

        return all_present

    def print_summary(self, success, start_time):
        """Print build summary."""
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()

        self.print_header("Build Summary")

        if success:
            print(f"  Status: ✓ SUCCESS")
            print(f"  Duration: {duration:.1f} seconds")
            print(f"  Output directory: {self.dist_dir}")
            print("\n  Built executables:")
            for _, exe_name in self.specs:
                print(f"    • {exe_name}")
        else:
            print(f"  Status: ✗ FAILED")
            print(f"  Duration: {duration:.1f} seconds")
            print("\n  Please check the error messages above.")

        print("=" * 70 + "\n")

    def run(self):
        """Execute the complete build process."""
        self.print_header("Serial Terminal Build Process")
        print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"  Python: {sys.version.split()[0]}")
        print(f"  Working directory: {self.root_dir}")

        start_time = datetime.now()

        # Step 1: Clean
        self.clean_build_artifacts()

        # Step 2-3: Build both executables
        build_results = []
        for idx, (spec_file, exe_name) in enumerate(self.specs, start=2):
            result = self.build_executable(spec_file, exe_name, idx)
            build_results.append(result)

        # Step 4: Verify
        verification_passed = self.verify_outputs()

        # Determine overall success
        success = all(build_results) and verification_passed

        # Print summary
        self.print_summary(success, start_time)

        return 0 if success else 1


def main():
    """Main entry point."""
    try:
        builder = BuildScript()
        exit_code = builder.run()
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n\n⚠ Build interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n✗ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
