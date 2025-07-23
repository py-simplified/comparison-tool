@echo off
echo ========================================
echo Advanced Excel File Comparison Tool
echo ========================================
echo.
echo This tool requires a 4-digit password for security.
echo.

REM Navigate to the script directory
cd /d "%~dp0"

REM Activate virtual environment and run advanced Python script
"C:/Users/Arnav/Documents/TCoE/COREP Comparison/.venv/Scripts/python.exe" excel_comparator_advanced.py

echo.
echo Press any key to exit...
pause >nul
