@echo off
echo ========================================
echo Password Configuration Utility
echo ========================================
echo.
echo This utility will update the password hash directly in the code files.
echo WARNING: This will modify the source code files!
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

REM Navigate to the script directory
cd /d "%~dp0"

REM Run the password configuration utility
"C:/Users/Arnav/Documents/TCoE/COREP Comparison/.venv/Scripts/python.exe" password_config.py

echo.
echo Press any key to exit...
pause >nul
