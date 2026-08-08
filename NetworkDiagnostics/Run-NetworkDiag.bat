@echo off
setlocal
title InfraEye - Network Diagnostics
color 0A

:: Ensure Administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator Privileges...
    powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

echo Starting Network Diagnostics...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0NetworkDiagnostics_Main.ps1"
