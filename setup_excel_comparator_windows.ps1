# Excel Comparator Setup Script - Windows PowerShell Version
# This script creates a complete Excel comparison tool setup
# Author: GitHub Copilot
# Date: July 23, 2025
# Version: 2.0 - Windows PowerShell

param(
    [switch]$Help
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Configuration
$PROJECT_NAME = "excel_comparator_simple"
$PYTHON_FILE = "excel_comparison_tool.py"
$VENV_NAME = "venv"

# Function to print colored output
function Write-Status {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Blue
}

function Write-Success {
    param([string]$Message)
    Write-Host "[SUCCESS] $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow
}

function Write-CustomError {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function Write-Header {
    param([string]$Message)
    Write-Host "[SETUP] $Message" -ForegroundColor Magenta
}

function Write-Feature {
    param([string]$Message)
    Write-Host "[FEATURE] $Message" -ForegroundColor Cyan
}

# Function to check if command exists
function Test-Command {
    param([string]$Command)
    try {
        $null = Get-Command $Command -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Function to create banner
function Write-Banner {
    Write-Host "================================================================" -ForegroundColor Magenta
    Write-Host "          Excel Comparator - Complete Setup Script" -ForegroundColor Magenta
    Write-Host "================================================================" -ForegroundColor Magenta
    Write-Host "🔧 This script will create a complete Excel comparison setup" -ForegroundColor Cyan
    Write-Host "📁 No manual configuration needed - everything automated!" -ForegroundColor Cyan
    Write-Host "🚀 Run once and start comparing Excel files immediately" -ForegroundColor Cyan
    Write-Host ""
}

# Function to activate virtual environment
function Invoke-VenvActivation {
    param([string]$VenvPath)
    
    if (Test-Path "$VenvPath\Scripts\Activate.ps1") {
        & "$VenvPath\Scripts\Activate.ps1"
        return $true
    } elseif (Test-Path "$VenvPath\Scripts\activate.bat") {
        & "$VenvPath\Scripts\activate.bat"
        return $true
    } else {
        Write-CustomError "Could not find activation script in $VenvPath"
        return $false
    }
}

# Function to deactivate virtual environment
function Invoke-VenvDeactivation {
    param([string]$VenvPath)
    
    if (Test-Path "$VenvPath\Scripts\deactivate.bat") {
        & "$VenvPath\Scripts\deactivate.bat"
    }
}

# Main setup function
function Start-Setup {
    Write-Banner
    
    # Check if Python is installed
    Write-Status "Checking Python installation..."
    $pythonCmd = $null
    $pythonVersion = $null
    
    if (Test-Command "python") {
        $pythonCmd = "python"
        $pythonVersion = & python --version 2>&1
        Write-Success "Python found: $pythonVersion"
    } elseif (Test-Command "py") {
        $pythonCmd = "py"
        $pythonVersion = & py --version 2>&1
        Write-Success "Python found: $pythonVersion"
    } else {
        Write-CustomError "Python is not installed. Please install Python 3.6+ and try again."
        Write-Host "Download from: https://www.python.org/downloads/"
        exit 1
    }
    
    # Check if pip is available
    Write-Status "Checking pip installation..."
    try {
        $null = & $pythonCmd -m pip --version 2>&1
        Write-Success "pip is available"
    } catch {
        Write-CustomError "pip is not available. Please install pip and try again."
        exit 1
    }
    
    # Create project directory
    Write-Header "Creating project directory: $PROJECT_NAME"
    if (Test-Path $PROJECT_NAME) {
        Write-Warning "Directory $PROJECT_NAME already exists."
        $response = Read-Host "Do you want to remove it and create fresh? (y/N)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            Remove-Item -Path $PROJECT_NAME -Recurse -Force
            Write-Success "Removed existing directory"
        } else {
            Write-CustomError "Setup cancelled by user"
            exit 1
        }
    }
    
    New-Item -ItemType Directory -Path $PROJECT_NAME | Out-Null
    Write-Success "Created directory: $PROJECT_NAME"
    
    Set-Location $PROJECT_NAME
    
    # Create folder structure
    Write-Header "Creating folder structure..."
    $folders = @("new", "prev", "template", "comparison_results", "logs")
    foreach ($folder in $folders) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    Write-Success "Created folders:"
    Write-Host "  📁 new/ - for new/current Excel files"
    Write-Host "  📁 prev/ - for previous Excel files"
    Write-Host "  📁 template/ - for template Excel files (formatting)"
    Write-Host "  📁 comparison_results/ - for output files"
    Write-Host "  📁 logs/ - for log files"
    
    # Create virtual environment
    Write-Header "Creating virtual environment..."
    if (Test-Path $VENV_NAME) {
        Write-Warning "Virtual environment already exists. Removing and recreating..."
        Remove-Item -Path $VENV_NAME -Recurse -Force
    }
    
    & $pythonCmd -m venv $VENV_NAME
    Write-Success "Virtual environment created: $VENV_NAME"
    
    # Activate virtual environment
    Write-Status "Activating virtual environment..."
    if (Invoke-VenvActivation -VenvPath $VENV_NAME) {
        Write-Success "Virtual environment activated"
    } else {
        Write-CustomError "Failed to activate virtual environment"
        exit 1
    }
    
    try {
        # Upgrade pip
        Write-Status "Upgrading pip..."
        & python -m pip install --upgrade pip --quiet
        Write-Success "pip upgraded to latest version"
        
        # Install required packages
        Write-Header "Installing required Python packages..."
        Write-Host "Installing packages: pandas, openpyxl, xlsxwriter, numpy"
        & python -m pip install "pandas>=1.3.0" "openpyxl>=3.0.0" "xlsxwriter>=3.0.0" "numpy>=1.21.0" --quiet
        Write-Success "All required packages installed successfully"
        
        # Create requirements.txt
        Write-Status "Creating requirements.txt..."
        @"
# Excel Comparison Tool Requirements
pandas>=1.3.0
openpyxl>=3.0.0
xlsxwriter>=3.0.0
numpy>=1.21.0
"@ | Out-File -FilePath "requirements.txt" -Encoding UTF8
        Write-Success "Created requirements.txt"
        
        # Create the main Python script
        Write-Header "Creating main Python script: $PYTHON_FILE"
        
        # Copy the Python template file if it exists, otherwise create it
        $templatePath = Join-Path $PSScriptRoot "excel_comparison_tool_template.py"
        if (Test-Path $templatePath) {
            Copy-Item $templatePath $PYTHON_FILE
            Write-Success "Copied Python script from template"
        } else {
            # Create a simplified version directly
            Write-Status "Creating Python script directly..."
            Create-PythonScript -FilePath $PYTHON_FILE
        }
        Write-Success "Created advanced Excel comparison script: $PYTHON_FILE"
        
        # Create execution scripts
        Write-Header "Creating execution scripts..."
        Create-ExecutionScripts
        
        # Create sample Excel files
        Write-Header "Creating sample Excel files for testing..."
        Create-SampleFiles
        
        # Create documentation
        Write-Header "Creating documentation..."
        Create-Documentation
        
        # Create requirements tracking
        & python -m pip freeze | Out-File -FilePath "requirements_full.txt" -Encoding UTF8
        Write-Success "Created complete requirements list"
        
    } finally {
        # Deactivate virtual environment
        Invoke-VenvDeactivation -VenvPath $VENV_NAME
    }
    
    # Create final status summary
    Create-StatusSummary -PythonVersion $pythonVersion
    
    # Final success message
    Show-FinalMessage
}

function Create-PythonScript {
    param([string]$FilePath)
    
    # Write Python script directly to file
@'
#!/usr/bin/env python3
"""
Excel Comparison Tool - Windows Version
Compares Excel files and highlights differences
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import os
import shutil
from pathlib import Path
import numpy as np
import json
from datetime import datetime
import sys

class ExcelComparator:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.new_folder = self.base_path / "new"
        self.prev_folder = self.base_path / "prev" 
        self.template_folder = self.base_path / "template"
        self.output_folder = self.base_path / "comparison_results"
        self.logs_folder = self.base_path / "logs"
        
        # Create output folders
        for folder in [self.output_folder, self.logs_folder]:
            folder.mkdir(exist_ok=True)
        
        # Styling
        self.red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        self.red_font = Font(color="FFFFFF", bold=True)
        
        self.log_file = self.logs_folder / f"comparison_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    def log_message(self, message, level="INFO"):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] [{level}] {message}"
        print(f"ℹ️  {message}")
        
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(log_entry + "\n")
    
    def get_matching_files(self):
        new_files = set(f.name for f in self.new_folder.glob("*.xlsx") if not f.name.startswith('~'))
        prev_files = set(f.name for f in self.prev_folder.glob("*.xlsx") if not f.name.startswith('~'))
        template_files = set(f.name for f in self.template_folder.glob("*.xlsx") if not f.name.startswith('~'))
        
        common_files = new_files.intersection(prev_files).intersection(template_files)
        return list(sorted(common_files))
    
    def compare_files(self, new_file, prev_file, template_file, output_file):
        # Copy template to output
        shutil.copy2(template_file, output_file)
        
        # Load workbooks
        new_wb = openpyxl.load_workbook(new_file, data_only=True)
        prev_wb = openpyxl.load_workbook(prev_file, data_only=True)
        output_wb = openpyxl.load_workbook(output_file)
        
        differences = 0
        
        # Compare sheets
        for sheet_name in new_wb.sheetnames:
            if sheet_name in prev_wb.sheetnames and sheet_name in output_wb.sheetnames:
                new_sheet = new_wb[sheet_name]
                prev_sheet = prev_wb[sheet_name]
                output_sheet = output_wb[sheet_name]
                
                max_row = max(new_sheet.max_row, prev_sheet.max_row)
                max_col = max(new_sheet.max_column, prev_sheet.max_column)
                
                for row in range(1, max_row + 1):
                    for col in range(1, max_col + 1):
                        new_cell = new_sheet.cell(row=row, column=col)
                        prev_cell = prev_sheet.cell(row=row, column=col)
                        output_cell = output_sheet.cell(row=row, column=col)
                        
                        if new_cell.value != prev_cell.value:
                            output_cell.fill = self.red_fill
                            output_cell.font = self.red_font
                            differences += 1
        
        output_wb.save(output_file)
        new_wb.close()
        prev_wb.close()
        output_wb.close()
        
        return differences
    
    def run_comparison(self):
        print("🚀 Excel Comparison Tool")
        print("=" * 50)
        
        matching_files = self.get_matching_files()
        
        if not matching_files:
            self.log_message("No matching Excel files found")
            print("❌ No files to compare. Add Excel files to all three folders.")
            return
        
        total_differences = 0
        
        for file_name in matching_files:
            print(f"📁 Processing: {file_name}")
            
            new_file = self.new_folder / file_name
            prev_file = self.prev_folder / file_name
            template_file = self.template_folder / file_name
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            name_parts = file_name.rsplit('.', 1)
            output_name = f"{name_parts[0]}_COMPARISON_{timestamp}.xlsx"
            output_file = self.output_folder / output_name
            
            try:
                differences = self.compare_files(new_file, prev_file, template_file, output_file)
                total_differences += differences
                
                if differences > 0:
                    print(f"   ✅ {differences} differences found and highlighted")
                else:
                    print(f"   ✅ No differences found")
                    
            except Exception as e:
                print(f"   ❌ Error: {str(e)}")
                self.log_message(f"Error processing {file_name}: {str(e)}", "ERROR")
        
        print(f"\n🎉 Comparison completed!")
        print(f"📊 Total differences found: {total_differences}")
        print(f"📁 Results saved in: {self.output_folder}")
        print(f"📄 Log file: {self.log_file}")

def main():
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        comparator = ExcelComparator(current_dir)
        comparator.run_comparison()
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
'@ | Out-File -FilePath $FilePath -Encoding UTF8
}

function Create-ExecutionScripts {
    # PowerShell run script
    $runScript = @'
# Excel Comparison Tool Runner - PowerShell

Write-Host "🚀 Starting Excel Comparison Tool..." -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Green

# Change to script directory
Set-Location $PSScriptRoot

# Check if virtual environment exists
if (-not (Test-Path "venv")) {
    Write-Host "❌ Virtual environment not found!" -ForegroundColor Red
    Write-Host "Please run the setup script first." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Activate virtual environment
Write-Host "🔧 Activating virtual environment..." -ForegroundColor Blue
if (Test-Path ".\venv\Scripts\Activate.ps1") {
    & ".\venv\Scripts\Activate.ps1"
} else {
    & ".\venv\Scripts\activate.bat"
}

# Check if main script exists
if (-not (Test-Path "excel_comparison_tool.py")) {
    Write-Host "❌ Main script not found!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Run the comparison
Write-Host "📊 Running Excel comparison..." -ForegroundColor Blue
& python excel_comparison_tool.py

# Deactivate virtual environment
if (Test-Path ".\venv\Scripts\deactivate.bat") {
    & ".\venv\Scripts\deactivate.bat"
}

Write-Host ""
Write-Host "✅ Comparison completed!" -ForegroundColor Green
Write-Host "📁 Check the comparison_results folder for output files." -ForegroundColor Cyan

# Wait for user input before closing
Read-Host "Press Enter to exit"
'@
    $runScript | Out-File -FilePath "run.ps1" -Encoding UTF8
    Write-Success "Created run.ps1"
    
    # Windows batch file
    $batchScript = @'
@echo off
echo 🚀 Starting Excel Comparison Tool...
echo ==================================

REM Change to script directory
cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "venv" (
    echo ❌ Virtual environment not found!
    echo Please run the setup script first.
    pause
    exit /b 1
)

REM Activate virtual environment
echo 🔧 Activating virtual environment...
call venv\Scripts\activate.bat

REM Check if main script exists
if not exist "excel_comparison_tool.py" (
    echo ❌ Main script not found!
    call venv\Scripts\deactivate.bat
    pause
    exit /b 1
)

REM Run the comparison
echo 📊 Running Excel comparison...
python excel_comparison_tool.py

REM Deactivate virtual environment
call venv\Scripts\deactivate.bat

echo.
echo ✅ Comparison completed!
echo 📁 Check the comparison_results folder for output files.
pause
'@
    $batchScript | Out-File -FilePath "run.bat" -Encoding ASCII
    Write-Success "Created run.bat"
}

function Create-SampleFiles {
@'
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import os

def create_sample_data():
    base_data = {
        'Account_ID': ['ACC001', 'ACC002', 'ACC003', 'ACC004', 'ACC005'],
        'Account_Name': ['Cash', 'Accounts Receivable', 'Inventory', 'Equipment', 'Accounts Payable'],
        'Q1_Amount': [50000, 120000, 80000, 200000, 45000],
        'Q2_Amount': [55000, 115000, 85000, 200000, 50000],
        'Status': ['Active', 'Active', 'Active', 'Active', 'Active']
    }
    
    modified_data = {
        'Account_ID': ['ACC001', 'ACC002', 'ACC003', 'ACC004', 'ACC005'],
        'Account_Name': ['Cash', 'Accounts Receivable', 'Inventory', 'Equipment', 'Accounts Payable'],
        'Q1_Amount': [52000, 120000, 82000, 200000, 45000],
        'Q2_Amount': [57000, 118000, 85000, 205000, 50000],
        'Status': ['Active', 'Active', 'Review', 'Active', 'Active']
    }
    
    return base_data, modified_data

def create_excel_file(data, filename, folder):
    wb = Workbook()
    ws = wb.active
    ws.title = "Financial_Data"
    
    df = pd.DataFrame(data)
    
    # Add headers
    for col, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    # Add data
    for row_idx, row_data in enumerate(df.values, 2):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    os.makedirs(folder, exist_ok=True)
    filepath = os.path.join(folder, filename)
    wb.save(filepath)
    print(f"Created: {filepath}")

def main():
    print("Creating sample Excel files...")
    base_data, modified_data = create_sample_data()
    
    create_excel_file(base_data, "sample_data.xlsx", "prev")
    create_excel_file(base_data, "sample_data.xlsx", "template")
    create_excel_file(modified_data, "sample_data.xlsx", "new")
    
    print("✅ Sample files created successfully!")

if __name__ == "__main__":
    main()
'@ | Out-File -FilePath "create_samples.py" -Encoding UTF8
    
    & python create_samples.py
    Remove-Item "create_samples.py"
    Write-Success "Sample Excel files created"
}

function Create-Documentation {
@'
# 📊 Excel Comparator Tool

## Quick Start

1. **Test with sample files:**
   - PowerShell: `.\run.ps1`
   - Command Prompt: `run.bat`

2. **Use with your files:**
   - Copy Excel files to `new/`, `prev/`, and `template/` folders
   - Files must have the same names in all folders
   - Run the tool again

## How it works

- Compares Excel files between `new/` and `prev/` folders
- Uses `template/` files for formatting
- Highlights differences in red
- Saves results to `comparison_results/` folder

## Folder Structure

```
excel_comparator_simple/
├── new/                    # Current Excel files
├── prev/                   # Previous Excel files  
├── template/               # Template files
├── comparison_results/     # Output files
├── logs/                   # Log files
└── venv/                   # Python environment
```

## Troubleshooting

- Ensure files exist in all three folders (new, prev, template)
- Files must have exactly the same names
- Use .xlsx format (not .xls)
- Close Excel before running the tool

🎉 Your tool is ready to use!
'@ | Out-File -FilePath "README.md" -Encoding UTF8
    Write-Success "Created README.md"
    
@'
🚀 EXCEL COMPARATOR - QUICK START

Your tool is ready!

STEP 1: Test with sample files
  PowerShell: .\run.ps1
  Command:    run.bat

STEP 2: Use with your files
  1. Copy Excel files to new/, prev/, template/
  2. Files must have same names in all folders
  3. Run the tool again

STEP 3: Check results
  - Look in comparison_results/ folder
  - Red highlighting shows differences

That's it!
'@ | Out-File -FilePath "QUICK_START.txt" -Encoding UTF8
    Write-Success "Created QUICK_START.txt"
}

function Create-StatusSummary {
    param([string]$PythonVersion)
    
    $currentPath = Get-Location
    $setupStatus = @"
=================================================================
EXCEL COMPARATOR SETUP - COMPLETION STATUS
=================================================================

✅ SETUP COMPLETED SUCCESSFULLY!

Created: $(Get-Date)
Python Version: $PythonVersion
Project Location: $currentPath

📁 FOLDERS CREATED:
  ✅ new/ - for current Excel files
  ✅ prev/ - for previous Excel files
  ✅ template/ - for template Excel files
  ✅ comparison_results/ - for output files
  ✅ logs/ - for log files
  ✅ venv/ - Python virtual environment

📄 FILES CREATED:
  ✅ excel_comparison_tool.py - Main script
  ✅ run.ps1 - PowerShell runner
  ✅ run.bat - Batch runner
  ✅ README.md - Documentation

🧪 SAMPLE FILES:
  ✅ sample_data.xlsx (in all folders)

🚀 READY TO USE:
  Run: .\run.ps1 or run.bat

=================================================================
"@
    $setupStatus | Out-File -FilePath "SETUP_STATUS.txt" -Encoding UTF8
}

function Show-FinalMessage {
    $currentPath = Get-Location
    
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Magenta
    Write-Feature "🎉 SETUP COMPLETED SUCCESSFULLY!"
    Write-Host "================================================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Success "Your Excel Comparator is ready!"
    Write-Host ""
    Write-Host "📁 Project location: $currentPath" -ForegroundColor White
    Write-Host "🚀 To get started:" -ForegroundColor White
    Write-Host "   PowerShell: .\run.ps1" -ForegroundColor Cyan
    Write-Host "   Command:    run.bat" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "📋 What's included:" -ForegroundColor White
    Write-Host "   ✅ Complete Excel comparison tool" -ForegroundColor Green
    Write-Host "   ✅ Virtual environment with dependencies" -ForegroundColor Green
    Write-Host "   ✅ Sample files for testing" -ForegroundColor Green
    Write-Host "   ✅ Execution scripts for Windows" -ForegroundColor Green
    Write-Host "   ✅ Documentation" -ForegroundColor Green
    Write-Host ""
    Write-Host "🎯 Test with sample files first!" -ForegroundColor Yellow
    Write-Host "📖 See README.md for complete instructions" -ForegroundColor Cyan
    Write-Host ""
    Write-Feature "Happy Excel comparing! 🎉"
}

# Show help if requested
if ($Help) {
    Write-Host "Excel Comparator Setup Script - Windows PowerShell Version"
    Write-Host "Usage: .\setup_excel_comparator_windows.ps1 [-Help]"
    Write-Host ""
    Write-Host "This script creates a complete Excel comparison tool setup."
    Write-Host "It will create a new directory with everything needed."
    Write-Host ""
    Write-Host "After setup, navigate to the created directory and run:"
    Write-Host "  .\run.ps1 (PowerShell) or run.bat (Command Prompt)"
    exit 0
}

# Run the main function
try {
    Start-Setup
} catch {
    Write-CustomError "Setup failed with error: $($_.Exception.Message)"
    Write-Host "Full error details:" -ForegroundColor Red
    Write-Host $_.Exception -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
