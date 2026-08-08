@echo off
setlocal
title InfraEye - Infrastructure Discovery
color 0A

:: Ensure Administrator privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator Privileges...
    powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

echo Starting Infrastructure Discovery (This may take a moment)...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0InfrastructureDiscovery_Main.ps1"
