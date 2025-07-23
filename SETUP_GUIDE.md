# Excel Comparison Tool - Complete Automation Setup Guide

This guide provides multiple automated setup scripts that will create the entire Excel comparison tool project from scratch.

## 🚀 Quick Setup Options

### Option 1: Automated Windows Setup (Recommended)
```batch
# Double-click this file for fully automated setup
setup_automated.bat
```

### Option 2: PowerShell Setup
```powershell
# Run this in PowerShell
.\setup_excel_comparator.ps1
```

### Option 3: Bash Setup (Linux/Mac/WSL)
```bash
# Run this in Bash
chmod +x setup_excel_comparator.sh
./setup_excel_comparator.sh
```

## 📋 What These Scripts Do

The automated setup scripts will:

1. **Create Project Structure**
   - Create new project directory
   - Set up folder structure (new/, prev/, template/, comparison_results/)

2. **Python Environment Setup**
   - Create Python virtual environment (.venv)
   - Install all required packages (pandas, openpyxl, xlsxwriter, numpy, pyinstaller)

3. **Generate Python Code**
   - Create `test.py` (basic comparison script)
   - Create `excel_comparator_advanced.py` (advanced with reporting)

4. **Create Execution Scripts**
   - Windows batch files (.bat)
   - PowerShell scripts (.ps1)
   - Bash scripts (.sh)

5. **Build Executable**
   - Create standalone .exe file using PyInstaller
   - No Python installation required for end users

6. **Generate Documentation**
   - README.md with complete usage instructions
   - LICENSE file
   - requirements.txt
   - .gitignore for version control

7. **Initialize Git Repository**
   - Set up Git repository
   - Make initial commit
   - Ready for GitHub upload

## 📁 Generated Project Structure

After running the setup, you'll get:

```
excel-comparison-tool/
├── 📁 new/                    # Place new Excel files here
├── 📁 prev/                   # Place previous Excel files here
├── 📁 template/               # Place template Excel files here
├── 📁 comparison_results/     # Output folder (auto-created)
├── 📁 dist/                   # Executable files
│   └── ExcelComparator.exe    # Standalone executable
├── 📁 .venv/                  # Python virtual environment
├── 📄 test.py                 # Basic comparison script
├── 📄 excel_comparator_advanced.py  # Advanced version
├── 📄 run_comparison.bat      # Windows batch runner
├── 📄 run_advanced_comparison.bat
├── 📄 run_comparison.sh       # Unix/Linux runner
├── 📄 run_advanced_comparison.sh
├── 📄 run_comparison.ps1      # PowerShell runner
├── 📄 run_advanced_comparison.ps1
├── 📄 requirements.txt        # Python dependencies
├── 📄 README.md              # Documentation
├── 📄 LICENSE                # MIT License
└── 📄 .gitignore             # Git ignore rules
```

## 🔧 Prerequisites

Before running the setup scripts, ensure you have:

1. **Python 3.8+** installed and in PATH
   ```bash
   python --version
   ```

2. **Git** (optional, for version control)
   ```bash
   git --version
   ```

3. **PowerShell** (for Windows PowerShell setup)

## 🎯 Usage After Setup

Once setup is complete, you can use the tool in several ways:

### Method 1: Executable (Easiest)
1. Place Excel files in `new/`, `prev/`, and `template/` folders
2. Double-click `dist/ExcelComparator.exe`

### Method 2: Batch Files (Windows)
1. Place Excel files in appropriate folders
2. Double-click `run_comparison.bat` or `run_advanced_comparison.bat`

### Method 3: PowerShell Scripts
1. Place Excel files in appropriate folders
2. Right-click → "Run with PowerShell" on `.ps1` files

### Method 4: Command Line
```bash
# Activate virtual environment
.venv/Scripts/activate    # Windows
source .venv/bin/activate # Linux/Mac

# Run comparison
python test.py                        # Basic
python excel_comparator_advanced.py  # Advanced
```

## 🎨 Features of Generated Tool

- ✅ **Multi-file Support**: Compares multiple Excel files at once
- ✅ **All Worksheets**: Processes every tab in each Excel file
- ✅ **Template Preservation**: Maintains original formatting from template files
- ✅ **Visual Highlighting**: Red highlighting for all differences
- ✅ **Numerical Calculations**: Shows `new_value - old_value` for numbers
- ✅ **Comprehensive Reports**: Detailed summary reports (advanced version)
- ✅ **Error Handling**: Robust error handling and logging
- ✅ **Cross-platform**: Works on Windows, Linux, and macOS

## 🔄 Version Control Ready

The generated project is ready for Git version control:

```bash
# The setup script already does this, but you can also:
git remote add origin https://github.com/yourusername/excel-comparison-tool.git
git branch -M main
git push -u origin main
```

## 🆘 Troubleshooting

### Common Issues:

1. **Python not found**
   - Install Python 3.8+ from python.org
   - Ensure Python is in system PATH

2. **Permission errors**
   - Run PowerShell as Administrator
   - Enable script execution: `Set-ExecutionPolicy RemoteSigned`

3. **PyInstaller fails**
   - Ensure virtual environment is activated
   - Try: `pip install --upgrade pyinstaller`

4. **Excel files not found**
   - Check file extensions (.xlsx only)
   - Ensure matching filenames in all three folders

## 📞 Support

If you encounter issues:
1. Check the generated README.md in your project
2. Review error messages in terminal/command prompt
3. Ensure all prerequisites are installed
4. Try running individual components manually

## 🎉 Success!

Once setup completes successfully, you'll have a fully functional Excel comparison tool ready for use or distribution!
