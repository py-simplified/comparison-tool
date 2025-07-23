# Simple Excel Comparator Setup for Windows
# This creates a basic Excel comparison tool

$ErrorActionPreference = "Stop"

Write-Host "🚀 Excel Comparator Setup for Windows" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green

# Check Python
Write-Host "Checking Python..." -ForegroundColor Blue
try {
    $pythonVersion = & python --version 2>&1
    Write-Host "✅ Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Python not found. Please install Python first." -ForegroundColor Red
    exit 1
}

# Create project directory
$projectName = "excel_comparator_tool"
Write-Host "Creating project: $projectName" -ForegroundColor Blue

if (Test-Path $projectName) {
    $response = Read-Host "Directory exists. Remove it? (y/N)"
    if ($response -eq 'y') {
        Remove-Item $projectName -Recurse -Force
    } else {
        exit 0
    }
}

New-Item -ItemType Directory $projectName | Out-Null
Set-Location $projectName

# Create folders
Write-Host "Creating folders..." -ForegroundColor Blue
@("new", "prev", "template", "comparison_results", "logs") | ForEach-Object {
    New-Item -ItemType Directory $_ -Force | Out-Null
}

# Create virtual environment
Write-Host "Creating virtual environment..." -ForegroundColor Blue
& python -m venv venv
if (-not $?) {
    Write-Host "❌ Failed to create virtual environment" -ForegroundColor Red
    exit 1
}

# Activate and install packages
Write-Host "Installing packages..." -ForegroundColor Blue
& ".\venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
& ".\venv\Scripts\python.exe" -m pip install pandas openpyxl xlsxwriter numpy --quiet

if (-not $?) {
    Write-Host "❌ Failed to install packages" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Setup completed!" -ForegroundColor Green
Write-Host ""
Write-Host "📁 Project created in: $(Get-Location)" -ForegroundColor Cyan
Write-Host "📖 Next steps:" -ForegroundColor Yellow
Write-Host "   1. Copy your Excel files to new/, prev/, and template/ folders"
Write-Host "   2. Files must have the same names in all three folders"
Write-Host "   3. Activate venv: .\venv\Scripts\Activate.ps1"
Write-Host "   4. Run your comparison script"
Write-Host ""
Write-Host "🎉 Your Excel comparator environment is ready!" -ForegroundColor Green
