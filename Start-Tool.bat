@echo off
REM M365 User Provisioning Tool - Easy Launcher for First Line Support
REM This batch file handles PowerShell execution for non-technical users

echo.
echo ================================================
echo M365 User Provisioning Tool - Easy Launcher
echo ================================================
echo.

REM Check if PowerShell 7 is available
pwsh -Command "exit 0" >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: PowerShell 7 is not available on this system
    echo This tool requires PowerShell 7.0 or later
    echo.
    echo Install PowerShell 7 from:
    echo https://github.com/PowerShell/PowerShell/releases
    echo.
    echo Or use: winget install Microsoft.PowerShell
    pause
    exit /b 1
)

echo Checking prerequisites...
echo.

REM Run prerequisites check first
pwsh -ExecutionPolicy Bypass -File "Install-Prerequisites.ps1"
if %errorlevel% neq 0 (
    echo.
    echo WARNING: Prerequisites check completed with warnings
    echo The tool may still work, but some features might be limited
    echo.
    set /p continue="Continue anyway? (y/n): "
    if /i not "%continue%"=="y" (
        echo Setup cancelled by user
        pause
        exit /b 1
    )
)

echo.
echo Starting M365 User Provisioning Tool...
echo.

REM Launch the main tool
pwsh -ExecutionPolicy Bypass -File "M365-UserProvisioning-Enterprise.ps1"

echo.
echo Tool has exited. Press any key to close this window.
pause >nul