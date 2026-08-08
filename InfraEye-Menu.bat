@echo off
setlocal EnableDelayedExpansion
title InfraEye IT Diagnostics - Master Menu
color 0B

:: Check for Administrator privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :menu
) else (
    echo Requesting Administrator Privileges...
    powershell -Command "Start-Process '%~dpnx0' -Verb RunAs"
    exit /b
)

:menu
cls
echo ==============================================================
echo                 INFRAEYE IT DIAGNOSTICS SUITE                 
echo ==============================================================
echo.
echo Please select a module to execute:
echo.
echo   [1] Device Health Diagnostics (Local Hardware/OS Check)
echo   [2] Network Diagnostics (Latency, IPs, Performance)
echo   [3] Infrastructure Discovery (Ping Sweep, Subnets, DHCP)
echo   [4] Python Unified Dashboard (Requires Python)
echo   [0] Exit
echo.
set /p choice="Enter your choice (0-4): "

if "!choice!"=="1" goto run_dh
if "!choice!"=="2" goto run_nd
if "!choice!"=="3" goto run_id
if "!choice!"=="4" goto run_py
if "!choice!"=="0" exit

echo Invalid choice. Please try again.
pause
goto menu

:run_dh
cls
echo Running Device Health Module...
cd /d "%~dp0DeviceHealth"
call Run-DeviceHealth.bat
pause
goto menu

:run_nd
cls
echo Running Network Diagnostics Module...
cd /d "%~dp0NetworkDiagnostics"
call Run-NetworkDiag.bat
pause
goto menu

:run_id
cls
echo Running Infrastructure Discovery Module...
cd /d "%~dp0InfrastructureDiscovery"
call Run-InfraDiscovery.bat
pause
goto menu

:run_py
cls
echo Running Python Unified Diagnostics...
cd /d "%~dp0"
python infra_eye_diagnostics.py
echo.
echo Process complete.
pause
goto menu
