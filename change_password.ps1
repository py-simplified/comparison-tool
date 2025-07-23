# Password Change Utility - PowerShell Version

Write-Host "========================================" -ForegroundColor Green
Write-Host "Password Change Utility" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host

Write-Host "Use this utility to change the 4-digit password." -ForegroundColor Yellow
Write-Host "You will need to know the current password." -ForegroundColor Yellow
Write-Host

# Navigate to script directory
Set-Location $PSScriptRoot

# Check if virtual environment exists
if (-not (Test-Path ".venv")) {
    Write-Host "❌ Virtual environment not found!" -ForegroundColor Red
    Write-Host "Please run the setup script first." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Activate virtual environment and run password change utility
& ".\.venv\Scripts\Activate.ps1"
& ".\.venv\Scripts\python.exe" test.py --change-password

Write-Host
Read-Host "Press Enter to exit"
