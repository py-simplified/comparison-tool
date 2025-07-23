@echo off
echo ========================================
echo Excel File Comparison Tool
echo ========================================
echo.
echo This tool requires a 4-digit password for security.
echo.

REM Navigate to the script directory
cd /d "%~dp0"

REM Activate virtual environment and run Python script
"C:/Users/Arnav/Documents/TCoE/COREP Comparison/.venv/Scripts/python.exe" test.py

echo.
echo Press any key to exit...
pause >nul
