#!/bin/bash

# Excel Comparison Tool - Complete Setup Script
# This script automates the entire process of creating the Excel comparison tool

set -e  # Exit on any error

# Color codes for better output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

print_step() {
    echo -e "${BLUE}[STEP]${NC} $1"
}

# Function to check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Main setup function
main() {
    echo "=========================================="
    echo "Excel Comparison Tool - Complete Setup"
    echo "=========================================="
    echo

    # Check prerequisites
    print_step "Checking prerequisites..."
    
    if ! command_exists python3; then
        print_error "Python 3 is not installed. Please install Python 3.8+ first."
        exit 1
    fi
    
    if ! command_exists git; then
        print_warning "Git is not installed. Git functionality will be skipped."
    fi

    # Get project directory from user or use default
    read -p "Enter project directory name (default: excel-comparison-tool): " PROJECT_DIR
    PROJECT_DIR=${PROJECT_DIR:-excel-comparison-tool}

    # Create project directory
    print_step "Creating project directory: $PROJECT_DIR"
    mkdir -p "$PROJECT_DIR"
    cd "$PROJECT_DIR"

    # Create folder structure
    print_step "Creating folder structure..."
    mkdir -p new prev template comparison_results

    # Create virtual environment
    print_step "Creating Python virtual environment..."
    python3 -m venv .venv
    
    # Activate virtual environment
    print_step "Activating virtual environment..."
    source .venv/bin/activate

    # Create requirements.txt
    print_step "Creating requirements.txt..."
    cat > requirements.txt << 'EOF'
pandas>=1.3.0
openpyxl>=3.0.0
xlsxwriter>=3.0.0
numpy>=1.21.0
pyinstaller>=5.0.0
EOF

    # Install dependencies
    print_step "Installing Python dependencies..."
    pip install --upgrade pip
    pip install -r requirements.txt

    # Create main comparison script
    print_step "Creating main comparison script (test.py)..."
    cat > test.py << 'EOF'
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import os
import shutil
from pathlib import Path
import numpy as np

class ExcelComparator:
    def __init__(self, base_path):
        """
        Initialize the Excel Comparator
        
        Args:
            base_path (str): Base directory containing new, prev, and template folders
        """
        self.base_path = Path(base_path)
        self.new_folder = self.base_path / "new"
        self.prev_folder = self.base_path / "prev"
        self.template_folder = self.base_path / "template"
        self.output_folder = self.base_path / "comparison_results"
        
        # Create output folder if it doesn't exist
        self.output_folder.mkdir(exist_ok=True)
        
        # Define red fill for highlighting differences
        self.red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        self.red_font = Font(color="FFFFFF", bold=True)
    
    def get_matching_files(self):
        """
        Get list of Excel files that exist in both new and prev folders
        
        Returns:
            list: List of file names present in both folders
        """
        new_files = set(f.name for f in self.new_folder.glob("*.xlsx"))
        prev_files = set(f.name for f in self.prev_folder.glob("*.xlsx"))
        template_files = set(f.name for f in self.template_folder.glob("*.xlsx"))
        
        # Get files that exist in all three folders
        common_files = new_files.intersection(prev_files).intersection(template_files)
        
        if not common_files:
            print("No matching Excel files found in all three folders (new, prev, template)")
            return []
        
        print(f"Found {len(common_files)} matching files:")
        for file in common_files:
            print(f"  - {file}")
        
        return list(common_files)
    
    def is_numeric(self, value):
        """
        Check if a value is numeric (int, float, or can be converted to float)
        
        Args:
            value: The value to check
            
        Returns:
            bool: True if numeric, False otherwise
        """
        if pd.isna(value) or value is None or value == "":
            return False
        
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    def compare_sheets(self, new_file, prev_file, template_file, output_file):
        """
        Compare all sheets in the Excel files and highlight differences
        
        Args:
            new_file (Path): Path to new Excel file
            prev_file (Path): Path to previous Excel file
            template_file (Path): Path to template Excel file
            output_file (Path): Path to output Excel file
        """
        # Copy template file to output location to preserve formatting
        shutil.copy2(template_file, output_file)
        
        # Load workbooks
        try:
            new_wb = openpyxl.load_workbook(new_file, data_only=True)
            prev_wb = openpyxl.load_workbook(prev_file, data_only=True)
            output_wb = openpyxl.load_workbook(output_file)
        except Exception as e:
            print(f"Error loading workbooks: {e}")
            return
        
        # Get common sheet names
        new_sheets = set(new_wb.sheetnames)
        prev_sheets = set(prev_wb.sheetnames)
        template_sheets = set(output_wb.sheetnames)
        
        common_sheets = new_sheets.intersection(prev_sheets).intersection(template_sheets)
        
        if not common_sheets:
            print(f"No common sheets found in {new_file.name}")
            return
        
        print(f"\nProcessing {len(common_sheets)} sheets in {new_file.name}:")
        
        differences_found = False
        
        for sheet_name in common_sheets:
            print(f"  Comparing sheet: {sheet_name}")
            
            try:
                new_sheet = new_wb[sheet_name]
                prev_sheet = prev_wb[sheet_name]
                output_sheet = output_wb[sheet_name]
                
                sheet_differences = self.compare_single_sheet(
                    new_sheet, prev_sheet, output_sheet, sheet_name
                )
                
                if sheet_differences:
                    differences_found = True
                    
            except Exception as e:
                print(f"    Error comparing sheet {sheet_name}: {e}")
                continue
        
        # Save the output file
        try:
            output_wb.save(output_file)
            if differences_found:
                print(f"  ✓ Differences found and highlighted in: {output_file.name}")
            else:
                print(f"  ✓ No differences found in: {output_file.name}")
        except Exception as e:
            print(f"  ✗ Error saving output file: {e}")
        finally:
            new_wb.close()
            prev_wb.close()
            output_wb.close()
    
    def compare_single_sheet(self, new_sheet, prev_sheet, output_sheet, sheet_name):
        """
        Compare a single sheet and highlight differences
        
        Args:
            new_sheet: New sheet object
            prev_sheet: Previous sheet object
            output_sheet: Output sheet object
            sheet_name: Name of the sheet
            
        Returns:
            bool: True if differences were found, False otherwise
        """
        differences_found = False
        differences_count = 0
        
        # Get the maximum row and column from both sheets
        max_row = max(new_sheet.max_row, prev_sheet.max_row)
        max_col = max(new_sheet.max_column, prev_sheet.max_column)
        
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                try:
                    # Get cell values
                    new_cell = new_sheet.cell(row=row, column=col)
                    prev_cell = prev_sheet.cell(row=row, column=col)
                    output_cell = output_sheet.cell(row=row, column=col)
                    
                    new_value = new_cell.value
                    prev_value = prev_cell.value
                    
                    # Skip if both cells are empty
                    if (new_value is None or new_value == "") and (prev_value is None or prev_value == ""):
                        continue
                    
                    # Check if values are different
                    if new_value != prev_value:
                        # Check if both values are numeric
                        if self.is_numeric(new_value) and self.is_numeric(prev_value):
                            try:
                                new_num = float(new_value) if new_value is not None else 0
                                prev_num = float(prev_value) if prev_value is not None else 0
                                difference = new_num - prev_num
                                
                                # Update the output cell with the difference
                                output_cell.value = difference
                                
                                # Apply red highlighting
                                output_cell.fill = self.red_fill
                                output_cell.font = self.red_font
                                
                                differences_found = True
                                differences_count += 1
                                
                            except (ValueError, TypeError):
                                # If conversion fails, treat as text difference
                                output_cell.value = new_value
                                output_cell.fill = self.red_fill
                                output_cell.font = self.red_font
                                differences_found = True
                                differences_count += 1
                        
                        elif self.is_numeric(new_value) and not self.is_numeric(prev_value):
                            # New value is numeric, old is not
                            try:
                                new_num = float(new_value)
                                output_cell.value = new_num  # Show the new numeric value
                                output_cell.fill = self.red_fill
                                output_cell.font = self.red_font
                                differences_found = True
                                differences_count += 1
                            except (ValueError, TypeError):
                                pass
                        
                        elif not self.is_numeric(new_value) and self.is_numeric(prev_value):
                            # Old value was numeric, new is not
                            output_cell.value = new_value  # Show the new non-numeric value
                            output_cell.fill = self.red_fill
                            output_cell.font = self.red_font
                            differences_found = True
                            differences_count += 1
                        
                        else:
                            # Both are non-numeric but different
                            output_cell.value = new_value
                            output_cell.fill = self.red_fill
                            output_cell.font = self.red_font
                            differences_found = True
                            differences_count += 1
                
                except Exception as e:
                    print(f"    Error processing cell {row},{col}: {e}")
                    continue
        
        if differences_count > 0:
            print(f"    Found {differences_count} differences in sheet '{sheet_name}'")
        
        return differences_found
    
    def run_comparison(self):
        """
        Run the complete comparison process for all matching files
        """
        print("Starting Excel Comparison Process...")
        print("="*50)
        
        matching_files = self.get_matching_files()
        
        if not matching_files:
            print("No files to compare. Exiting.")
            return
        
        print(f"\nProcessing {len(matching_files)} files...")
        print("="*50)
        
        for file_name in matching_files:
            print(f"\nProcessing file: {file_name}")
            print("-" * 40)
            
            new_file = self.new_folder / file_name
            prev_file = self.prev_folder / file_name
            template_file = self.template_folder / file_name
            
            # Create output filename with timestamp
            name_parts = file_name.rsplit('.', 1)
            output_name = f"{name_parts[0]}_COMPARISON.xlsx"
            output_file = self.output_folder / output_name
            
            try:
                self.compare_sheets(new_file, prev_file, template_file, output_file)
            except Exception as e:
                print(f"  ✗ Error processing {file_name}: {e}")
                continue
        
        print("\n" + "="*50)
        print("Comparison process completed!")
        print(f"Results saved in: {self.output_folder}")
        print("\nOutput files contain:")
        print("- Original formatting from template files")
        print("- Differences highlighted in RED")
        print("- For numeric differences: shows (new value - old value)")
        print("- For non-numeric differences: shows the new value")


def main():
    """
    Main function to run the Excel comparison
    """
    # Get the current directory (where the script is located)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("Excel File Comparison Tool")
    print("="*30)
    print(f"Working directory: {current_dir}")
    print()
    
    # Initialize and run the comparator
    comparator = ExcelComparator(current_dir)
    comparator.run_comparison()


if __name__ == "__main__":
    main()
EOF

    # Create advanced comparison script
    print_step "Creating advanced comparison script..."
    create_advanced_script

    # Create batch files for Windows
    print_step "Creating Windows batch files..."
    create_batch_files

    # Create shell scripts for Unix/Linux/Mac
    print_step "Creating shell scripts..."
    create_shell_scripts

    # Create .gitignore
    print_step "Creating .gitignore..."
    create_gitignore

    # Create README.md
    print_step "Creating README.md..."
    create_readme

    # Create LICENSE
    print_step "Creating LICENSE..."
    create_license

    # Create executable using PyInstaller
    print_step "Creating executable file..."
    create_executable

    # Initialize Git repository if Git is available
    if command_exists git; then
        print_step "Initializing Git repository..."
        git init
        git add .
        git commit -m "Initial commit: Excel Comparison Tool"
        print_status "Git repository initialized with initial commit."
    fi

    # Final summary
    echo
    echo "=========================================="
    echo "Setup Complete!"
    echo "=========================================="
    echo
    print_status "Project created successfully in: $(pwd)"
    echo
    echo "What was created:"
    echo "├── Python Scripts:"
    echo "│   ├── test.py (basic comparison)"
    echo "│   └── excel_comparator_advanced.py (advanced with reports)"
    echo "├── Executable:"
    echo "│   └── dist/ExcelComparator.exe"
    echo "├── Batch Files (Windows):"
    echo "│   ├── run_comparison.bat"
    echo "│   └── run_advanced_comparison.bat"
    echo "├── Shell Scripts (Unix/Linux/Mac):"
    echo "│   ├── run_comparison.sh"
    echo "│   └── run_advanced_comparison.sh"
    echo "├── Folders:"
    echo "│   ├── new/ (place new Excel files here)"
    echo "│   ├── prev/ (place previous Excel files here)"
    echo "│   ├── template/ (place template Excel files here)"
    echo "│   └── comparison_results/ (output folder)"
    echo "└── Documentation:"
    echo "    ├── README.md"
    echo "    ├── LICENSE"
    echo "    └── requirements.txt"
    echo
    echo "To use the tool:"
    echo "1. Place Excel files in new/, prev/, and template/ folders"
    echo "2. Run: ./run_comparison.sh or double-click run_comparison.bat"
    echo "3. Or use the executable: ./dist/ExcelComparator.exe"
    echo
    print_status "Ready to use!"
}

# Function to create advanced script
create_advanced_script() {
    cat > excel_comparator_advanced.py << 'EOF'
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
import os
import shutil
from pathlib import Path
import numpy as np
from datetime import datetime
import json

class ExcelComparatorAdvanced:
    def __init__(self, base_path):
        """
        Initialize the Advanced Excel Comparator
        
        Args:
            base_path (str): Base directory containing new, prev, and template folders
        """
        self.base_path = Path(base_path)
        self.new_folder = self.base_path / "new"
        self.prev_folder = self.base_path / "prev"
        self.template_folder = self.base_path / "template"
        self.output_folder = self.base_path / "comparison_results"
        
        # Create output folder if it doesn't exist
        self.output_folder.mkdir(exist_ok=True)
        
        # Define red fill for highlighting differences
        self.red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        self.red_font = Font(color="FFFFFF", bold=True)
        
        # Initialize summary data
        self.summary_data = {
            "timestamp": datetime.now().isoformat(),
            "files_processed": [],
            "total_differences": 0,
            "errors": []
        }
    
    def get_matching_files(self):
        """
        Get list of Excel files that exist in both new and prev folders
        
        Returns:
            list: List of file names present in both folders
        """
        new_files = set(f.name for f in self.new_folder.glob("*.xlsx"))
        prev_files = set(f.name for f in self.prev_folder.glob("*.xlsx"))
        template_files = set(f.name for f in self.template_folder.glob("*.xlsx"))
        
        # Get files that exist in all three folders
        common_files = new_files.intersection(prev_files).intersection(template_files)
        
        if not common_files:
            print("No matching Excel files found in all three folders (new, prev, template)")
            return []
        
        print(f"Found {len(common_files)} matching files:")
        for file in common_files:
            print(f"  - {file}")
        
        return list(common_files)
    
    def is_numeric(self, value):
        """
        Check if a value is numeric (int, float, or can be converted to float)
        
        Args:
            value: The value to check
            
        Returns:
            bool: True if numeric, False otherwise
        """
        if pd.isna(value) or value is None or value == "":
            return False
        
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    def compare_sheets(self, new_file, prev_file, template_file, output_file):
        """
        Compare all sheets in the Excel files and highlight differences
        
        Args:
            new_file (Path): Path to new Excel file
            prev_file (Path): Path to previous Excel file
            template_file (Path): Path to template Excel file
            output_file (Path): Path to output Excel file
            
        Returns:
            dict: Summary of comparison results
        """
        # Copy template file to output location to preserve formatting
        shutil.copy2(template_file, output_file)
        
        file_summary = {
            "filename": new_file.name,
            "sheets_processed": 0,
            "total_differences": 0,
            "sheet_details": {},
            "errors": []
        }
        
        # Load workbooks
        try:
            new_wb = openpyxl.load_workbook(new_file, data_only=True)
            prev_wb = openpyxl.load_workbook(prev_file, data_only=True)
            output_wb = openpyxl.load_workbook(output_file)
        except Exception as e:
            error_msg = f"Error loading workbooks: {e}"
            print(error_msg)
            file_summary["errors"].append(error_msg)
            return file_summary
        
        # Get common sheet names
        new_sheets = set(new_wb.sheetnames)
        prev_sheets = set(prev_wb.sheetnames)
        template_sheets = set(output_wb.sheetnames)
        
        common_sheets = new_sheets.intersection(prev_sheets).intersection(template_sheets)
        
        if not common_sheets:
            error_msg = f"No common sheets found in {new_file.name}"
            print(error_msg)
            file_summary["errors"].append(error_msg)
            return file_summary
        
        print(f"\nProcessing {len(common_sheets)} sheets in {new_file.name}:")
        
        differences_found = False
        
        for sheet_name in common_sheets:
            print(f"  Comparing sheet: {sheet_name}")
            
            try:
                new_sheet = new_wb[sheet_name]
                prev_sheet = prev_wb[sheet_name]
                output_sheet = output_wb[sheet_name]
                
                sheet_result = self.compare_single_sheet(
                    new_sheet, prev_sheet, output_sheet, sheet_name
                )
                
                file_summary["sheet_details"][sheet_name] = sheet_result
                file_summary["total_differences"] += sheet_result["differences_count"]
                file_summary["sheets_processed"] += 1
                
                if sheet_result["differences_found"]:
                    differences_found = True
                    
            except Exception as e:
                error_msg = f"Error comparing sheet {sheet_name}: {e}"
                print(f"    {error_msg}")
                file_summary["errors"].append(error_msg)
                continue
        
        # Save the output file
        try:
            output_wb.save(output_file)
            if differences_found:
                print(f"  ✓ {file_summary['total_differences']} differences found and highlighted in: {output_file.name}")
            else:
                print(f"  ✓ No differences found in: {output_file.name}")
                
            file_summary["status"] = "completed"
            
        except Exception as e:
            error_msg = f"Error saving output file: {e}"
            print(f"  ✗ {error_msg}")
            file_summary["errors"].append(error_msg)
            file_summary["status"] = "failed"
        finally:
            new_wb.close()
            prev_wb.close()
            output_wb.close()
            
        return file_summary
    
    def compare_single_sheet(self, new_sheet, prev_sheet, output_sheet, sheet_name):
        """
        Compare a single sheet and highlight differences
        
        Args:
            new_sheet: New sheet object
            prev_sheet: Previous sheet object
            output_sheet: Output sheet object
            sheet_name: Name of the sheet
            
        Returns:
            dict: Summary of sheet comparison
        """
        sheet_result = {
            "differences_found": False,
            "differences_count": 0,
            "numeric_differences": 0,
            "text_differences": 0,
            "cells_compared": 0
        }
        
        # Get the maximum row and column from both sheets
        max_row = max(new_sheet.max_row, prev_sheet.max_row)
        max_col = max(new_sheet.max_column, prev_sheet.max_column)
        
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                try:
                    # Get cell values
                    new_cell = new_sheet.cell(row=row, column=col)
                    prev_cell = prev_sheet.cell(row=row, column=col)
                    output_cell = output_sheet.cell(row=row, column=col)
                    
                    new_value = new_cell.value
                    prev_value = prev_cell.value
                    
                    sheet_result["cells_compared"] += 1
                    
                    # Skip if both cells are empty
                    if (new_value is None or new_value == "") and (prev_value is None or prev_value == ""):
                        continue
                    
                    # Check if values are different
                    if new_value != prev_value:
                        # Check if both values are numeric
                        if self.is_numeric(new_value) and self.is_numeric(prev_value):
                            try:
                                new_num = float(new_value) if new_value is not None else 0
                                prev_num = float(prev_value) if prev_value is not None else 0
                                difference = new_num - prev_num
                                
                                # Update the output cell with the difference
                                output_cell.value = difference
                                
                                # Apply red highlighting
                                output_cell.fill = self.red_fill
                                output_cell.font = self.red_font
                                
                                sheet_result["differences_found"] = True
                                sheet_result["differences_count"] += 1
                                sheet_result["numeric_differences"] += 1
                                
                            except (ValueError, TypeError):
                                # If conversion fails, treat as text difference
                                output_cell.value = new_value
                                output_cell.fill = self.red_fill
                                output_cell.font = self.red_font
                                sheet_result["differences_found"] = True
                                sheet_result["differences_count"] += 1
                                sheet_result["text_differences"] += 1
                        
                        elif self.is_numeric(new_value) and not self.is_numeric(prev_value):
                            # New value is numeric, old is not
                            try:
                                new_num = float(new_value)
                                output_cell.value = new_num  # Show the new numeric value
                                output_cell.fill = self.red_fill
                                output_cell.font = self.red_font
                                sheet_result["differences_found"] = True
                                sheet_result["differences_count"] += 1
                                sheet_result["numeric_differences"] += 1
                            except (ValueError, TypeError):
                                pass
                        
                        elif not self.is_numeric(new_value) and self.is_numeric(prev_value):
                            # Old value was numeric, new is not
                            output_cell.value = new_value  # Show the new non-numeric value
                            output_cell.fill = self.red_fill
                            output_cell.font = self.red_font
                            sheet_result["differences_found"] = True
                            sheet_result["differences_count"] += 1
                            sheet_result["text_differences"] += 1
                        
                        else:
                            # Both are non-numeric but different
                            output_cell.value = new_value
                            output_cell.fill = self.red_fill
                            output_cell.font = self.red_font
                            sheet_result["differences_found"] = True
                            sheet_result["differences_count"] += 1
                            sheet_result["text_differences"] += 1
                
                except Exception as e:
                    print(f"    Error processing cell {row},{col}: {e}")
                    continue
        
        if sheet_result["differences_count"] > 0:
            print(f"    Found {sheet_result['differences_count']} differences in sheet '{sheet_name}' " +
                  f"({sheet_result['numeric_differences']} numeric, {sheet_result['text_differences']} text)")
        
        return sheet_result
    
    def generate_summary_report(self):
        """
        Generate a detailed summary report of the comparison process
        """
        summary_file = self.output_folder / "comparison_summary.json"
        report_file = self.output_folder / "comparison_report.txt"
        
        # Save JSON summary
        with open(summary_file, 'w') as f:
            json.dump(self.summary_data, f, indent=2)
        
        # Generate text report
        with open(report_file, 'w') as f:
            f.write("EXCEL COMPARISON SUMMARY REPORT\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"Comparison completed at: {self.summary_data['timestamp']}\n")
            f.write(f"Total files processed: {len(self.summary_data['files_processed'])}\n")
            f.write(f"Total differences found: {self.summary_data['total_differences']}\n\n")
            
            if self.summary_data['errors']:
                f.write("ERRORS ENCOUNTERED:\n")
                f.write("-" * 20 + "\n")
                for error in self.summary_data['errors']:
                    f.write(f"- {error}\n")
                f.write("\n")
            
            f.write("FILE DETAILS:\n")
            f.write("-" * 20 + "\n")
            for file_data in self.summary_data['files_processed']:
                f.write(f"\nFile: {file_data['filename']}\n")
                f.write(f"  Status: {file_data.get('status', 'unknown')}\n")
                f.write(f"  Sheets processed: {file_data['sheets_processed']}\n")
                f.write(f"  Total differences: {file_data['total_differences']}\n")
                
                if file_data['sheet_details']:
                    f.write("  Sheet breakdown:\n")
                    for sheet_name, details in file_data['sheet_details'].items():
                        f.write(f"    {sheet_name}: {details['differences_count']} differences " +
                               f"({details['numeric_differences']} numeric, {details['text_differences']} text)\n")
                
                if file_data['errors']:
                    f.write("  Errors:\n")
                    for error in file_data['errors']:
                        f.write(f"    - {error}\n")
        
        print(f"\nSummary report saved to: {report_file}")
        print(f"Detailed JSON data saved to: {summary_file}")
    
    def run_comparison(self):
        """
        Run the complete comparison process for all matching files
        """
        print("Starting Advanced Excel Comparison Process...")
        print("="*50)
        
        matching_files = self.get_matching_files()
        
        if not matching_files:
            print("No files to compare. Exiting.")
            return
        
        print(f"\nProcessing {len(matching_files)} files...")
        print("="*50)
        
        for file_name in matching_files:
            print(f"\nProcessing file: {file_name}")
            print("-" * 40)
            
            new_file = self.new_folder / file_name
            prev_file = self.prev_folder / file_name
            template_file = self.template_folder / file_name
            
            # Create output filename with timestamp
            name_parts = file_name.rsplit('.', 1)
            output_name = f"{name_parts[0]}_COMPARISON.xlsx"
            output_file = self.output_folder / output_name
            
            try:
                file_summary = self.compare_sheets(new_file, prev_file, template_file, output_file)
                self.summary_data['files_processed'].append(file_summary)
                self.summary_data['total_differences'] += file_summary['total_differences']
                
                # Add any file errors to global errors
                if file_summary['errors']:
                    self.summary_data['errors'].extend(file_summary['errors'])
                    
            except Exception as e:
                error_msg = f"Error processing {file_name}: {e}"
                print(f"  ✗ {error_msg}")
                self.summary_data['errors'].append(error_msg)
                continue
        
        print("\n" + "="*50)
        print("Advanced comparison process completed!")
        print(f"Results saved in: {self.output_folder}")
        print(f"Total differences found: {self.summary_data['total_differences']}")
        print("\nOutput files contain:")
        print("- Original formatting from template files")
        print("- Differences highlighted in RED")
        print("- For numeric differences: shows (new value - old value)")
        print("- For non-numeric differences: shows the new value")
        
        # Generate summary report
        self.generate_summary_report()


def main():
    """
    Main function to run the Excel comparison
    """
    # Get the current directory (where the script is located)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("Advanced Excel File Comparison Tool")
    print("="*35)
    print(f"Working directory: {current_dir}")
    print()
    
    # Initialize and run the comparator
    comparator = ExcelComparatorAdvanced(current_dir)
    comparator.run_comparison()


if __name__ == "__main__":
    main()
EOF
}

# Function to create batch files for Windows
create_batch_files() {
    # Basic comparison batch file
    cat > run_comparison.bat << 'EOF'
@echo off
echo ========================================
echo Excel File Comparison Tool
echo ========================================
echo.
echo Starting comparison process...
echo.

REM Navigate to the script directory
cd /d "%~dp0"

REM Activate virtual environment and run Python script
call .venv\Scripts\activate.bat
python test.py

echo.
echo Comparison completed!
echo Check the 'comparison_results' folder for output files.
echo.
pause
EOF

    # Advanced comparison batch file
    cat > run_advanced_comparison.bat << 'EOF'
@echo off
echo ========================================
echo Advanced Excel File Comparison Tool
echo ========================================
echo.
echo Starting advanced comparison process with detailed reporting...
echo.

REM Navigate to the script directory
cd /d "%~dp0"

REM Activate virtual environment and run advanced Python script
call .venv\Scripts\activate.bat
python excel_comparator_advanced.py

echo.
echo Advanced comparison completed!
echo Check the 'comparison_results' folder for:
echo - Excel files with highlighted differences
echo - comparison_report.txt (detailed summary)
echo - comparison_summary.json (machine-readable data)
echo.
pause
EOF

    # Make batch files executable
    chmod +x *.bat 2>/dev/null || true
}

# Function to create shell scripts for Unix/Linux/Mac
create_shell_scripts() {
    # Basic comparison shell script
    cat > run_comparison.sh << 'EOF'
#!/bin/bash

echo "========================================"
echo "Excel File Comparison Tool"
echo "========================================"
echo
echo "Starting comparison process..."
echo

# Navigate to the script directory
cd "$(dirname "$0")"

# Activate virtual environment and run Python script
source .venv/bin/activate
python test.py

echo
echo "Comparison completed!"
echo "Check the 'comparison_results' folder for output files."
echo
EOF

    # Advanced comparison shell script
    cat > run_advanced_comparison.sh << 'EOF'
#!/bin/bash

echo "========================================"
echo "Advanced Excel File Comparison Tool"
echo "========================================"
echo
echo "Starting advanced comparison process with detailed reporting..."
echo

# Navigate to the script directory
cd "$(dirname "$0")"

# Activate virtual environment and run advanced Python script
source .venv/bin/activate
python excel_comparator_advanced.py

echo
echo "Advanced comparison completed!"
echo "Check the 'comparison_results' folder for:"
echo "- Excel files with highlighted differences"
echo "- comparison_report.txt (detailed summary)"
echo "- comparison_summary.json (machine-readable data)"
echo
EOF

    # Make shell scripts executable
    chmod +x *.sh
}

# Function to create .gitignore
create_gitignore() {
    cat > .gitignore << 'EOF'
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
pip-wheel-metadata/
share/python-wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST

# Virtual Environment
.venv/
venv/
ENV/
env/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
.DS_Store?
._*
.Spotlight-V100
.Trashes
ehthumbs.db
Thumbs.db

# Excel temporary files
~$*.xlsx
~$*.xls

# Output files (you may want to exclude these from version control)
comparison_results/
*.tmp

# Logs
*.log

# PyInstaller
*.spec
build/
dist/
EOF
}

# Function to create README
create_readme() {
    cat > README.md << 'EOF'
# Excel Comparison Tool 📊

An intelligent Python automation tool that compares Excel files between two versions and highlights differences with visual formatting. Perfect for financial reports, data analysis, and document version control.

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-.xlsx-green)](https://www.microsoft.com/en-us/microsoft-365/excel)

## ✨ Features

- 🔄 **Automated Comparison**: Compares all matching Excel files between `new` and `prev` folders
- 🎨 **Template-based Output**: Uses files from `template` folder to maintain original formatting
- 📋 **All Tabs Support**: Compares all worksheets/tabs in each Excel file
- 🧮 **Numerical Difference Calculation**: For numerical data, shows the difference (new value - old value)
- 🎯 **Visual Highlighting**: Differences are highlighted in red background with white bold text
- 🛡️ **Comprehensive Error Handling**: Handles various data types and edge cases
- 📊 **Detailed Reporting**: Advanced version includes comprehensive reports and statistics
- 💻 **Executable Version**: Standalone .exe file for easy distribution

## 🚀 Quick Start

### Using the Executable (Easiest)
1. Download and extract the project
2. Place Excel files in `new/`, `prev/`, and `template/` folders
3. Double-click `dist/ExcelComparator.exe`

### Using Scripts
**Windows:**
- Double-click `run_comparison.bat` (basic)
- Double-click `run_advanced_comparison.bat` (with reports)

**Mac/Linux:**
```bash
./run_comparison.sh          # Basic comparison
./run_advanced_comparison.sh # Advanced with reports
```

### Using Python Directly
```bash
source .venv/bin/activate    # Activate virtual environment
python test.py               # Basic comparison
python excel_comparator_advanced.py  # Advanced version
```

## 📁 Project Structure

```
excel-comparison-tool/
├── 📁 new/                    # Place new/current Excel files here
├── 📁 prev/                   # Place previous Excel files here
├── 📁 template/               # Place template Excel files here
├── 📁 comparison_results/     # Output folder (auto-created)
├── 📁 dist/                   # Executable files
│   └── ExcelComparator.exe    # Standalone executable
├── 📄 test.py                 # Main comparison script
├── 📄 excel_comparator_advanced.py  # Advanced version
├── 📄 run_comparison.bat/.sh  # Easy execution scripts
├── 📄 run_advanced_comparison.bat/.sh
└── 📄 requirements.txt        # Dependencies
```

## 🔧 How It Works

1. **📁 File Matching**: Finds Excel files that exist in all three folders
2. **📋 Template Copying**: Preserves original formatting from template
3. **📊 Sheet Comparison**: Compares each worksheet individually
4. **🔍 Cell Analysis**: 
   - Numeric differences: Shows `new_value - old_value`
   - Text differences: Shows new value
   - Highlights all differences in red
5. **💾 Output Generation**: Creates comparison files with "_COMPARISON" suffix

## 📈 Output

- **📍 Location**: `comparison_results/` folder
- **📝 Files**: 
  - `*_COMPARISON.xlsx` - Excel files with highlighted differences
  - `comparison_report.txt` - Human-readable summary
  - `comparison_summary.json` - Machine-readable data
- **🎨 Formatting**: Red background, white bold text for differences

## 🛠️ Development Setup

If you want to modify or rebuild the project:

```bash
# Clone and setup
git clone <repository-url>
cd excel-comparison-tool

# Create virtual environment
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the automated setup script
./setup_excel_comparator.sh
```

## 📦 Building Executable

To create your own executable:

```bash
pyinstaller --onefile --windowed --name ExcelComparator excel_comparator_advanced.py
```

## 🤝 Contributing

1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏷️ Tags

`excel` `comparison` `automation` `python` `data-analysis` `xlsx` `openpyxl` `pandas` `pyinstaller`

---

⭐ **Star this repository if you find it helpful!** ⭐
EOF
}

# Function to create LICENSE
create_license() {
    cat > LICENSE << 'EOF'
MIT License

Copyright (c) 2025 Excel Comparison Tool

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
EOF
}

# Function to create executable
create_executable() {
    print_status "Building executable with PyInstaller..."
    
    # Create a spec file for better control
    cat > ExcelComparator.spec << 'EOF'
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel_comparator_advanced.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'openpyxl', 'xlsxwriter', 'numpy'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelComparator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
EOF

    # Build the executable
    pyinstaller ExcelComparator.spec --clean --noconfirm
    
    if [ -f "dist/ExcelComparator" ] || [ -f "dist/ExcelComparator.exe" ]; then
        print_status "Executable created successfully in dist/ folder"
    else
        print_warning "Executable creation may have failed. Check dist/ folder."
    fi
    
    # Clean up build files
    rm -rf build/
    rm -f ExcelComparator.spec
}

# Run the main function
main "$@"
EOF
