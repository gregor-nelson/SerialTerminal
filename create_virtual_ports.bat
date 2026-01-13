@echo off
setlocal EnableDelayedExpansion
:: ============================================================================
:: Virtual COM Port Pair Creator
:: Creates a com0com virtual serial port pair with standard settings
::
:: Usage: Double-click to run (will auto-request admin privileges)
:: ============================================================================

:: Change to script directory first
cd /d "%~dp0"

title Virtual COM Port Pair Creator

echo.
echo ============================================
echo   Virtual COM Port Pair Creator
echo ============================================
echo.
echo This tool creates a virtual COM port pair
echo using com0com. Data sent to one port will
echo appear on the other port and vice versa.
echo.
echo ============================================
echo.

:: Check for admin privileges and self-elevate if needed
echo [Step 1/4] Checking administrator privileges...
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo    Not running as administrator.
    echo    Requesting elevation - please click YES on the UAC prompt...
    echo.

    :: Create a temporary VBScript to elevate with visible window
    echo Set UAC = CreateObject^("Shell.Application"^) > "%TEMP%\elevate.vbs"
    echo UAC.ShellExecute "%~f0", "", "%~dp0", "runas", 1 >> "%TEMP%\elevate.vbs"
    cscript //nologo "%TEMP%\elevate.vbs"
    del "%TEMP%\elevate.vbs"
    exit /b
)
echo    OK - Running with administrator privileges.
echo.

:: Set com0com path
set "SETUPC=C:\Program Files (x86)\com0com\setupc.exe"

:: Check if com0com is installed
echo [Step 2/4] Checking com0com installation...
if not exist "%SETUPC%" goto :com0com_not_found
echo    OK - Found com0com at: %SETUPC%
echo.
goto :com0com_found

:com0com_not_found
echo.
echo ============================================
echo   ERROR: com0com not installed
echo ============================================
echo.
echo Could not find setupc.exe at:
echo %SETUPC%
echo.
echo Please install com0com first from:
echo https://sourceforge.net/projects/com0com/
echo.
pause
exit /b 1

:com0com_found

:: Prompt for mode selection
echo [Step 3/5] Select configuration mode...
echo.
echo    [1] Serial Router defaults (COM131/132 and COM141/142)
echo    [2] Custom port numbers
echo.
set /p MODE_CHOICE="    Enter choice (1 or 2): "

if "%MODE_CHOICE%"=="1" goto :serial_router_defaults
if "%MODE_CHOICE%"=="2" goto :custom_ports
echo.
echo    Invalid choice. Using Serial Router defaults...
goto :serial_router_defaults

:serial_router_defaults
echo.
echo    Using Serial Router default port pairs:
echo      - COM131 ^<-^> COM132 (for application 1)
echo      - COM141 ^<-^> COM142 (for application 2)
echo.
set "CREATE_PAIR1=yes"
set "PAIR1_A=131"
set "PAIR1_B=132"
set "CREATE_PAIR2=yes"
set "PAIR2_A=141"
set "PAIR2_B=142"
goto :create_ports

:custom_ports
echo.
echo    Enter the COM port numbers you want to create.
echo    Recommended range: 150-199 (to avoid conflicts)
echo.
set /p PAIR1_A="    Enter first COM port number (e.g., 150): COM"
set /p PAIR1_B="    Enter second COM port number (e.g., 151): COM"

:: Validate input
if "%PAIR1_A%"=="" goto :empty_port_error
if "%PAIR1_B%"=="" goto :empty_port_error

set "CREATE_PAIR1=yes"
set "CREATE_PAIR2=no"
goto :create_ports

:empty_port_error
echo.
echo ERROR: Port numbers cannot be empty
pause
exit /b 1

:create_ports
echo.

:: Build the command parameters
set "PORT_PARAMS=EmuBR=yes,EmuOverrun=yes,PlugInMode=no,ExclusiveMode=no,HiddenMode=no,AllDataBits=yes,cts=rrts,dsr=rdtr,dcd=rdtr"

:: Change to com0com directory (setupc.exe needs com0com.inf in current directory)
pushd "C:\Program Files (x86)\com0com"

:: Clean up existing virtual port pairs first
echo [Step 4/6] Removing existing virtual port pairs...
echo.
echo    Checking for existing pairs...

:: Get list output and remove pairs (try removing pairs 0-9)
for /L %%i in (0,1,9) do (
    setupc.exe remove %%i >nul 2>&1
)
echo    Cleanup complete.
echo.

:: Execute the command(s)
echo [Step 5/6] Creating virtual port pair(s)...
echo.

set TOTAL_RESULT=0

:: Create first pair
if "%CREATE_PAIR1%"=="yes" (
    echo    Creating: COM%PAIR1_A% ^<-^> COM%PAIR1_B%
    echo    -------------------------------------------------
    setupc.exe install PortName=COM%PAIR1_A%,%PORT_PARAMS% PortName=COM%PAIR1_B%,%PORT_PARAMS%
    set RESULT1=!errorlevel!
    echo    -------------------------------------------------
    if !RESULT1! neq 0 set TOTAL_RESULT=1
    echo.
)

:: Create second pair (only for Serial Router defaults)
if "%CREATE_PAIR2%"=="yes" (
    echo    Creating: COM%PAIR2_A% ^<-^> COM%PAIR2_B%
    echo    -------------------------------------------------
    setupc.exe install PortName=COM%PAIR2_A%,%PORT_PARAMS% PortName=COM%PAIR2_B%,%PORT_PARAMS%
    set RESULT2=!errorlevel!
    echo    -------------------------------------------------
    if !RESULT2! neq 0 set TOTAL_RESULT=1
    echo.
)

popd

echo [Step 6/6] Results...
echo.

if %TOTAL_RESULT% equ 0 goto :success_message
goto :error_message

:success_message
echo ============================================
echo   SUCCESS!
echo ============================================
echo.
echo   Virtual port pair(s) created successfully:
echo.
if "%CREATE_PAIR1%"=="yes" echo      COM%PAIR1_A%  ^<----^>  COM%PAIR1_B%
if "%CREATE_PAIR2%"=="yes" echo      COM%PAIR2_A%  ^<----^>  COM%PAIR2_B%
echo.
if "%MODE_CHOICE%"=="1" (
    echo   For Serial Router:
    echo     - Connect app 1 to COM132
    echo     - Connect app 2 to COM142
    echo     - Router uses COM131 and COM141 internally
    echo.
)
echo   Data sent to one port appears on its pair.
echo.
echo ============================================
goto :end_script

:error_message
echo ============================================
echo   ERROR: Failed to create port pair(s)
echo ============================================
echo.
echo   One or more commands failed.
echo.
echo   This may happen if:
echo     - The COM port numbers are already in use
echo     - com0com drivers are not properly installed
echo     - Another application is using the ports
echo.
echo ============================================

:end_script
echo.
echo ============================================
echo Press any key to exit...
echo ============================================
pause >nul
exit /b 0
