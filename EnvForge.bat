@echo off
:: EnvForge - Windows Environment Variable Manager
:: Launches with admin prompt for full System variable editing

set PYTHON="C:\Users\gvaro\AppData\Local\Programs\Python\Python312\pythonw.exe"
set SCRIPT="%~dp0EnvForge.py"

:: Check if already admin
net session >nul 2>&1
if %errorlevel% == 0 (
    %PYTHON% %SCRIPT%
) else (
    :: Re-launch as admin
    powershell -Command "Start-Process '%PYTHON%' -ArgumentList '%SCRIPT%' -Verb RunAs"
)
