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
-  **Multiple Execution Options**: Batch files, PowerShell scripts, and cross-platform support

## 🏗️ Project Structure

```
excel-comparison-tool/
├── 📁 new/                    # Folder containing new/current Excel files
├── 📁 prev/                   # Folder containing previous Excel files  
├── 📁 template/               # Folder containing template Excel files (for formatting)
├── 📁 comparison_results/     # Output folder (created automatically)
├── 📄 test.py                 # Main comparison script
├── 📄 excel_comparator_advanced.py  # Advanced version with detailed reporting
├── 📄 run_comparison.bat      # Windows batch file for easy execution
├── 📄 run_advanced_comparison.bat   # Batch file for advanced version
├── 📄 requirements.txt        # Python dependencies
└── 📄 README.md              # This file
```

## 🚀 Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/excel-comparison-tool.git
cd excel-comparison-tool
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Prepare Your Files
Place your Excel files in the appropriate folders:
- `new/`: Current/latest versions
- `prev/`: Previous versions  
- `template/`: Template files for formatting

### 4. Run the Comparison

**Option A: Using Batch Files (Windows)**
- Double-click `run_comparison.bat` for basic comparison
- Double-click `run_advanced_comparison.bat` for detailed reporting
**Option B: Using PowerShell Scripts**
- Right-click → "Run with PowerShell" on `.ps1` files

**Option C: Command Line**
```bash
python test.py                           # Basic comparison
python excel_comparator_advanced.py     # Advanced with reports
```

## 📋 Requirements

Create a `requirements.txt` file with:
```
pandas>=1.3.0
openpyxl>=3.0.0
xlsxwriter>=3.0.0
numpy>=1.21.0
```

## 🔧 How It Works

1. **📁 File Matching**: Finds Excel files (.xlsx) that exist in all three folders (new, prev, template)
2. **📋 Template Copying**: Creates output files by copying from template folder to preserve formatting
3. **📊 Sheet-by-Sheet Comparison**: Compares each worksheet individually
4. **🔍 Cell-by-Cell Analysis**: 
   - For numerical differences: Calculates `new_value - old_value`
   - For text differences: Shows the new value
   - Highlights all differences in red
5. **💾 Result Generation**: Saves comparison results with "_COMPARISON" suffix

## 📈 Output

- **📍 Location**: `comparison_results/` folder
- **📝 Naming**: Original filename + `_COMPARISON.xlsx`
- **🎨 Formatting**: Maintains original template formatting
- **🔴 Highlighting**: Red background with white bold text for differences
- **📊 Reports**: Summary reports (advanced version only)

## 🎯 Data Type Handling

- **🔢 Numeric Values**: Shows difference calculation (new - old)
- **📝 Text Values**: Shows new value when different
- **⬜ Empty Cells**: Properly handles null/empty values
- **🔄 Mixed Types**: Handles conversion between numeric and text data

## 📊 Advanced Features

The advanced version (`excel_comparator_advanced.py`) includes:

- 📈 **Detailed Statistics**: Counts of numeric vs text differences
- 📋 **Comprehensive Reports**: Text and JSON format summaries
- 🕒 **Timestamp Tracking**: When comparisons were performed
- 🔍 **Error Logging**: Detailed error tracking and reporting
- 📊 **Sheet-level Breakdown**: Statistics for each worksheet

## 🛠️ Error Handling

The tool includes comprehensive error handling for:
- ❌ Missing files or folders
- 💥 Corrupted Excel files
- 🔄 Data type conversion issues
- 💾 Memory limitations for large files
- 🔒 Permission issues

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏷️ Tags

`excel` `comparison` `automation` `python` `data-analysis` `reporting` `xlsx` `openpyxl` `pandas`

## 📞 Support

If you encounter any issues or have questions, please [open an issue](https://github.com/yourusername/excel-comparison-tool/issues) on GitHub.

---

⭐ **Star this repository if you find it helpful!** ⭐
