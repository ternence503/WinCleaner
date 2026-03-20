@echo off
chcp 65001 >/dev/null 2>&1
title WinCleaner v3.0

echo.
echo  Starting WinCleaner...
echo.

:: Get script directory
set "SCRIPT_DIR=%~dp0"
set "PS_FILE=%SCRIPT_DIR%WinCleaner.ps1"

:: Check if PowerShell script exists
if not exist "%PS_FILE%" (
    echo  Error: WinCleaner.ps1 not found
    echo  Make sure WinCleaner.ps1 is in the same folder as this file
    pause
    exit /b 1
)

:: Run PowerShell script as Administrator
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo  If the tool did not open, try right-clicking this .bat file
    echo  and selecting "Run as administrator"
    echo.
    pause
)
