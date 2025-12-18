@echo off
:: ============================================================================
:: xsukax Windows Users Manager Launcher
:: ============================================================================
:: This batch file launches the xsukax Windows Users Manager PowerShell GUI
:: with administrative privileges.
::
:: Author: xsukax
:: Website: Tech Me Away !!!
:: Version: 2.1
:: ============================================================================

title xsukax Windows Users Manager Launcher

:: Check for administrative privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :RunScript
) else (
    echo.
    echo ============================================================
    echo  Administrator Privileges Required
    echo ============================================================
    echo.
    echo This application requires administrator privileges to
    echo manage Windows user accounts.
    echo.
    echo The script will now request elevation...
    echo.
    timeout /t 3 >nul
    goto :ElevateScript
)

:ElevateScript
:: Request elevation using PowerShell
powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
exit /b

:RunScript
cls
echo.
echo ============================================================
echo  xsukax Windows Users Manager
echo  Created by: xsukax
echo ============================================================
echo.
echo Starting PowerShell GUI application...
echo.

:: Set the path to the PowerShell script
:: Assumes the .ps1 file is in the same directory as this batch file
set "SCRIPT_PATH=%~dp0WindowsUsersManager.ps1"

:: Check if the PowerShell script exists
if not exist "%SCRIPT_PATH%" (
    echo ERROR: PowerShell script not found!
    echo.
    echo Expected location: %SCRIPT_PATH%
    echo.
    echo Please ensure WindowsUsersManager.ps1 is in the same
    echo directory as this launcher.
    echo.
    pause
    exit /b 1
)

:: Launch the PowerShell script with execution policy bypass
:: -NoProfile: Speeds up startup by not loading profile scripts
:: -ExecutionPolicy Bypass: Allows the script to run without changing system policy
:: -File: Specifies the script file to execute
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

:: Check for errors
if %errorLevel% neq 0 (
    echo.
    echo ============================================================
    echo  Error Detected
    echo ============================================================
    echo.
    echo The application encountered an error during execution.
    echo Error Code: %errorLevel%
    echo.
    echo Please check the following:
    echo  - You have administrator privileges
    echo  - PowerShell 5.1 or higher is installed
    echo  - The script file is not corrupted
    echo.
    pause
    exit /b %errorLevel%
)

:: Script completed successfully
exit /b 0