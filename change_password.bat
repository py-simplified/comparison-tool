@echo off
echo ========================================
echo Password Change Utility
echo ========================================
echo.
echo Use this utility to change the 4-digit password.
echo You will need to know the current password.
echo.

REM Navigate to the script directory
cd /d "%~dp0"

REM Run the password change utility
"C:/Users/Arnav/Documents/TCoE/COREP Comparison/.venv/Scripts/python.exe" test.py --change-password

echo.
echo Press any key to exit...
pause >nul
