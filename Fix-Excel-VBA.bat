@echo off
title Excel VBA Access Setup
echo Excel VBA Access Setup Tool
echo ================================
echo.
echo This tool will help you configure Excel to allow VBA integration.
echo If you're getting "null-valued expression" errors, this should fix it.
echo.
pause

cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "scripts\Setup-ExcelVBAAccess.ps1"

echo.
echo Setup completed. Press any key to close...
pause >nul