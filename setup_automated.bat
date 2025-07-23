@echo off
echo ========================================================
echo Excel Comparison Tool - Automated Setup
echo ========================================================
echo.
echo This script will automatically:
echo - Create a new project with all necessary files
echo - Set up Python virtual environment
echo - Install all dependencies
echo - Create executable files
echo - Generate documentation
echo.
echo Press any key to start the automated setup...
pause >nul

echo.
echo Starting automated setup...
echo.

REM Run the PowerShell setup script
powershell -ExecutionPolicy Bypass -File "setup_excel_comparator.ps1"

echo.
echo Setup completed! Check the output above for any errors.
pause
