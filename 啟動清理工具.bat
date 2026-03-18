@echo off
chcp 65001 >nul 2>&1
title WinCleaner - Windows 系統清理工具

echo.
echo  正在啟動清理工具... / Starting WinCleaner...
echo.

:: 取得此批次檔所在目錄
set "SCRIPT_DIR=%~dp0"
set "PS_FILE=%SCRIPT_DIR%WinCleaner.ps1"

:: 檢查 PowerShell 腳本是否存在
if not exist "%PS_FILE%" (
    echo  錯誤：找不到 WinCleaner.ps1
    echo  Error: WinCleaner.ps1 not found
    echo  請確認 WinCleaner.ps1 與此檔案在同一資料夾
    echo  Make sure WinCleaner.ps1 is in the same folder
    pause
    exit /b 1
)

:: 以系統管理員身份執行 PowerShell 腳本
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo  若工具未能開啟，請改用以下方式：
    echo  If the tool didn't open, try this instead:
    echo.
    echo  右鍵點擊此 .bat 檔案
    echo  Right-click this .bat file
    echo  選擇「以系統管理員身份執行」
    echo  Select "Run as administrator"
    echo.
    pause
)
