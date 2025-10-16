@echo off
REM Build script for Serial Terminal application using PyInstaller

echo Building Serial Terminal...
echo.

REM Navigate to project root (one level up from scripts folder)
cd /d "%~dp0.."

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller not found. Installing...
    python -m pip install pyinstaller
)

echo.
echo Running PyInstaller...
echo.

REM Run PyInstaller command
REM --onefile: Create a single executable
REM --windowed: Don't show console window
REM --uac-admin: Request admin privileges (for serial port access)
REM --name: Application name
REM --icon: Application icon (optional, will use default if not present)

if exist "assets\icons\app.ico" (
    pyinstaller --onefile --windowed --uac-admin --name "Serial Terminal" --icon "assets\icons\app.ico" main.py
) else (
    echo Warning: No icon file found at assets\icons\app.ico
    echo Building without custom icon...
    pyinstaller --onefile --windowed --uac-admin --name "Serial Terminal" main.py
)

echo.
if %ERRORLEVEL% EQU 0 (
    echo ========================================
    echo Build successful!
    echo ========================================
    echo Executable created at: dist\Serial Terminal.exe
    echo.
    echo You can now run the application by executing:
    echo   dist\"Serial Terminal.exe"
    echo.
) else (
    echo ========================================
    echo Build failed with error code: %ERRORLEVEL%
    echo ========================================
    echo.
    echo Common issues:
    echo - Make sure all dependencies are installed: pip install -r requirements.txt
    echo - Check that Python is properly installed and in PATH
    echo - Ensure no antivirus is blocking PyInstaller
    echo.
)

pause
