@echo off
title Parking Report Generator - GUI

REM Change to script directory
cd /d "%~dp0"

REM Run PowerShell script with GUI (script handles its own exit)
powershell.exe -ExecutionPolicy Bypass -File "scripts\Get-ParkingReport_GUI_DE.ps1"