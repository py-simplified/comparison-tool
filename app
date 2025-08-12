#!/usr/bin/env python3
"""
Excel File Comparison Tool for COREP Reports

This tool compares Excel files between 'new' and 'prev' folders, focusing on numerical differences.
It ignores timestamps and generates detailed comparison reports with highlighted differences.

Features:
- Compares all tabs in Excel files
- Highlights numerical differences in red
- Generates summary report with detailed difference information
- Handles .xls (Excel 97-2003) format
- Ignores date/time stamps in filename matching

Author: Python Expert Assistant
Date: July 23, 2025
"""

import os
import re
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import xlrd
from xlutils.copy import copy
import xlwt
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class ExcelComparer:
    """
    A comprehensive Excel file comparison tool specifically designed for COREP reports.
    """

    def __init__(self, new_folder, prev_folder, output_base_folder="comparison_results"):
        """
        Initialize the Excel comparer.

        Args:
            new_folder (str): Path to folder containing new Excel files
            prev_folder (str): Path to folder containing previous Excel files
            output_base_folder (str): Base path for saving comparison results
        """
        self.new_folder = new_folder
        self.prev_folder = prev_folder
        
        # Create timestamped output folder
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_folder = os.path.join(output_base_folder, f"comparison_{timestamp}")
        
        self.summary_data = []
        self.comparison_summary = []  # Track all comparisons (with and without differences)

        # Create output folder if it doesn't exist
        os.makedirs(self.output_folder, exist_ok=True)
        logger.info(f"Results will be saved to: {self.output_folder}")

        # Define fill styles for highlighting differences
        self.red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        self.bold_font = Font(bold=True)

    def normalize_filename(self, filename):
        """
        Normalize filename by removing:
        - the date (DD-MM-YYYY) after the first hyphen
        - the 14-digit timestamp just before the last part

        Args:
            filename (str): Original filename

        Returns:
            str: Normalized filename for matching
        """
        # Remove file extension
        name_without_ext = os.path.splitext(filename)[0].strip()

        # Split by ' - ' delimiter (consistent with the filenames shown)
        parts = name_without_ext.split(' - ')

        cleaned_parts = []

        for _, part in enumerate(parts):
            # Remove part if it's a date (e.g., 31-08-2024)
            if re.match(r"\d{2}-\d{2}-\d{4}$", part):
                continue

            # Remove part if it's a 14-digit timestamp (e.g., 20241105230128)
            if re.match(r"\d{14}$", part):
                continue

            cleaned_parts.append(part.strip())

        return ' - '.join(cleaned_parts).lower()

    def find_matching_files(self):
        """
        Find matching files between new and prev folders based on normalized names.

        Returns:
            list: List of tuples (new_file, prev_file) for matching pairs
        """
        new_files = [f for f in os.listdir(self.new_folder) if f.endswith('.xls')]
        prev_files = [f for f in os.listdir(self.prev_folder) if f.endswith('.xls')]

        # Create mapping of normalized names to actual filenames
        new_mapping = {self.normalize_filename(f): f for f in new_files}
        prev_mapping = {self.normalize_filename(f): f for f in prev_files}

        matching_pairs = []
        for norm_name in new_mapping:
            if norm_name in prev_mapping:
                matching_pairs.append((new_mapping[norm_name], prev_mapping[norm_name]))
            else:
                logger.warning(f"No matching file found for: {new_mapping[norm_name]}")

        logger.info(f"Found {len(matching_pairs)} matching file pairs")
        return matching_pairs

    def read_excel_file(self, file_path):
        """
        Read Excel file and return data for all sheets.

        Args:
            file_path (str): Path to Excel file

        Returns:
            dict: Dictionary with sheet names as keys and DataFrames as values
        """
        try:
            # Read all sheets from the Excel file without treating first row as header
            # This ensures row numbering matches the actual Excel file
            sheets_data = pd.read_excel(file_path, sheet_name=None, engine='xlrd', header=None)
            logger.info(f"Successfully read {len(sheets_data)} sheets from {os.path.basename(file_path)}")
            return sheets_data
        except Exception as e:
            logger.error(f"Error reading {file_path}: {str(e)}")
            return {}

    def is_numeric_value(self, value):
        """
        Check if a value is numeric (int, float, or numeric string).

        Args:
            value: Value to check

        Returns:
            bool: True if value is numeric
        """
        if pd.isna(value):
            return False

        if isinstance(value, (int, float)):
            return not pd.isna(value)

        if isinstance(value, str):
            try:
                float(value.replace(',', '').replace('$', '').replace('%', ''))
                return True
            except ValueError:
                return False

        return False

    def convert_to_numeric(self, value):
        """
        Convert value to numeric, handling various formats.

        Args:
            value: Value to convert

        Returns:
            float: Numeric value or NaN if conversion fails
        """
        if pd.isna(value):
            return np.nan

        if isinstance(value, (int, float)):
            return float(value)

        if isinstance(value, str):
            try:
                # Remove common formatting characters
                clean_value = value.replace(',', '').replace('$', '').replace('%', '').strip()
                return float(clean_value)
            except ValueError:
                return np.nan

        return np.nan

    def compare_dataframes(self, df1, df2, sheet_name, file_name):
        """
        Compare two DataFrames and identify numerical differences.

        Args:
            df1 (pd.DataFrame): DataFrame from new file
            df2 (pd.DataFrame): DataFrame from previous file
            sheet_name (str): Name of the sheet being compared
            file_name (str): Name of the file being compared

        Returns:
            list: List of differences found
        """
        differences = []

        # Ensure both DataFrames have the same shape by padding with NaN
        max_rows = max(len(df1), len(df2))
        max_cols = max(len(df1.columns), len(df2.columns))

        # Resize DataFrames to same dimensions
        df1 = df1.reindex(range(max_rows), fill_value=np.nan)
        df2 = df2.reindex(range(max_rows), fill_value=np.nan)

        # Ensure same column count
        while len(df1.columns) < max_cols:
            df1[f'col_{len(df1.columns)}'] = np.nan
        while len(df2.columns) < max_cols:
            df2[f'col_{len(df2.columns)}'] = np.nan

        # Compare each cell
        for row_idx in range(max_rows):
            for col_idx in range(max_cols):
                try:
                    val1 = df1.iloc[row_idx, col_idx] if row_idx < len(df1) else np.nan
                    val2 = df2.iloc[row_idx, col_idx] if row_idx < len(df2) else np.nan

                    # Only compare if either value is numeric-like
                    if self.is_numeric_value(val1) or self.is_numeric_value(val2):
                        num1 = self.convert_to_numeric(val1)
                        num2 = self.convert_to_numeric(val2)

                        # Check for differences (considering NaN values)
                        if pd.isna(num1) and pd.isna(num2):
                            continue  # Both are NaN, no difference
                        elif pd.isna(num1) or pd.isna(num2):
                            # One is NaN, the other is not
                            difference = {
                                'file': file_name,
                                'sheet': sheet_name,
                                'row': row_idx + 1,  # 1-based indexing for Excel
                                'column': col_idx + 1,
                                'column_letter': self.get_column_letter(col_idx + 1),
                                'new_value': val1 if not pd.isna(num1) else 'Empty',
                                'prev_value': val2 if not pd.isna(num2) else 'Empty',
                                'difference': 'N/A'
                            }
                            differences.append(difference)
                        elif abs(num1 - num2) > 1e-10:  # tolerance for float compare
                            difference = {
                                'file': file_name,
                                'sheet': sheet_name,
                                'row': row_idx + 1,
                                'column': col_idx + 1,
                                'column_letter': self.get_column_letter(col_idx + 1),
                                'new_value': num1,
                                'prev_value': num2,
                                'difference': num1 - num2
                            }
                            differences.append(difference)
                except Exception as e:
                    logger.warning(f"Error comparing cell ({row_idx}, {col_idx}) in {sheet_name}: {str(e)}")
                    continue

        return differences

    def get_column_letter(self, col_num):
        """
        Convert column number to Excel column letter (A, B, C, ..., AA, AB, etc.).

        Args:
            col_num (int): Column number (1-based)

        Returns:
            str: Excel column letter
        """
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result

    def create_highlighted_file(self, original_file, differences, output_path):
        """
        Create a copy of the original file with differences highlighted in red.

        Args:
            original_file (str): Path to original Excel file
            differences (list): List of differences to highlight
            output_path (str): Path to save the highlighted file
        """
        try:
            logger.info(f"Creating highlighted .xls file with {len(differences)} differences")
            
            # Read the original Excel file
            workbook = xlrd.open_workbook(original_file, formatting_info=True)
            new_workbook = copy(workbook)

            # Create style for highlighting with red background
            style = xlwt.XFStyle()
            pattern = xlwt.Pattern()
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
            style.pattern = pattern
            
            # Make text bold and white for better visibility
            font = xlwt.Font()
            font.bold = True
            font.colour_index = xlwt.Style.colour_map['white']
            style.font = font

            # Group differences by sheet
            sheet_differences = {}
            for diff in differences:
                sheet_name = str(diff['sheet'])
                if sheet_name not in sheet_differences:
                    sheet_differences[sheet_name] = []
                sheet_differences[sheet_name].append(diff)

            logger.info(f"Differences grouped by sheet: {list(sheet_differences.keys())}")

            # Apply highlighting to each sheet
            for sheet_name, sheet_diffs in sheet_differences.items():
                try:
                    # Find the sheet by name
                    sheet_index = None
                    for idx, name in enumerate(workbook.sheet_names()):
                        if str(name) == sheet_name:
                            sheet_index = idx
                            break
                    
                    if sheet_index is None:
                        logger.warning(f"Sheet '{sheet_name}' not found in workbook")
                        continue
                    
                    worksheet = new_workbook.get_sheet(sheet_index)
                    original_sheet = workbook.sheet_by_index(sheet_index)
                    
                    logger.info(f"Processing sheet '{sheet_name}' (index {sheet_index}) with {len(sheet_diffs)} differences")

                    for diff in sheet_diffs:
                        row = diff['row'] - 1  # Convert to zero-based indexing
                        col = diff['column'] - 1
                        
                        try:
                            # Get the original cell value
                            cell_value = original_sheet.cell_value(row, col)
                            
                            # Write the cell with highlighting style
                            worksheet.write(row, col, cell_value, style)
                            
                            logger.info(f"✅ Highlighted cell {diff['column_letter']}{diff['row']} in sheet '{sheet_name}' (value: {cell_value})")
                            
                        except Exception as cell_error:
                            logger.error(f"❌ Error highlighting cell {diff['column_letter']}{diff['row']}: {str(cell_error)}")

                except Exception as sheet_error:
                    logger.error(f"❌ Error processing sheet {sheet_name}: {str(sheet_error)}")

            # Save highlighted file
            new_workbook.save(output_path)
            logger.info(f"✅ Highlighted .xls file saved: {output_path}")

        except Exception as e:
            logger.error(f"❌ Error creating highlighted file for {original_file}: {str(e)}")
            # Don't raise the exception, just log it so the process continues
            return False
        
        return True

    def compare_files(self, new_file, prev_file):
        """
        Compare two workbooks sheet-by-sheet and highlight differences on the new file.

        Args:
            new_file (str): Filename in new folder
            prev_file (str): Filename in prev folder

        Returns:
            list: List of all differences found
        """
        logger.info(f"Comparing {new_file} with {prev_file}")

        new_path = os.path.join(self.new_folder, new_file)
        prev_path = os.path.join(self.prev_folder, prev_file)

        # Read both files
        new_data = self.read_excel_file(new_path)
        prev_data = self.read_excel_file(prev_path)

        if not new_data or not prev_data:
            logger.warning(f"Could not read one or both files: {new_file}, {prev_file}")
            # Still track this comparison
            self.comparison_summary.append({
                'new_file': new_file,
                'prev_file': prev_file,
                'status': 'Error - Could not read files',
                'total_differences': 0,
                'sheets_compared': 0,
                'sheets_with_differences': 0
            })
            return []

        all_differences = []
        sheets_compared = 0
        sheets_with_differences = 0

        # Get all sheet names from both files
        all_sheets = set(new_data.keys()) | set(prev_data.keys())

        for sheet_name in all_sheets:
            if sheet_name in new_data and sheet_name in prev_data:
                sheets_compared += 1
                differences = self.compare_dataframes(
                    new_data[sheet_name],
                    prev_data[sheet_name],
                    sheet_name,
                    new_file
                )
                all_differences.extend(differences)
                if differences:
                    sheets_with_differences += 1
                logger.info(f"Found {len(differences)} differences in sheet '{sheet_name}'")
            else:
                logger.warning(f"Sheet '{sheet_name}' not found in both files")

        # Track this comparison summary
        comparison_status = "No differences found" if len(all_differences) == 0 else f"{len(all_differences)} differences found"
        self.comparison_summary.append({
            'new_file': new_file,
            'prev_file': prev_file,
            'status': comparison_status,
            'total_differences': len(all_differences),
            'sheets_compared': sheets_compared,
            'sheets_with_differences': sheets_with_differences
        })

        # Create highlighted version of the new file
        if all_differences:
            highlighted_filename = f"highlighted_{new_file}"
            highlighted_path = os.path.join(self.output_folder, highlighted_filename)
            success = self.create_highlighted_file(new_path, all_differences, highlighted_path)
            
            if success:
                logger.info(f"✅ Successfully created highlighted file: {highlighted_filename}")
            else:
                logger.warning(f"⚠️ Failed to create highlighted file for: {new_file}")

        return all_differences

    def create_summary_report(self):
        """
        Create a comprehensive summary report of all differences found.
        """
        # Create summary DataFrame for detailed differences
        summary_df = pd.DataFrame(self.summary_data) if self.summary_data else pd.DataFrame()

        # Create comparison overview DataFrame
        comparison_df = pd.DataFrame(self.comparison_summary)

        # Create Excel workbook for summary
        summary_path = os.path.join(self.output_folder, "comparison_summary.xlsx")

        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
            # Write comparison overview first
            if not comparison_df.empty:
                comparison_df.to_excel(writer, sheet_name='Comparison Overview', index=False)
                
                # Get the workbook and worksheet for overview
                workbook = writer.book
                overview_ws = writer.sheets['Comparison Overview']
                
                # Auto-adjust column widths for overview
                for column in overview_ws.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except Exception:
                            pass
                    adjusted_width = min(max_length + 2, 60)
                    overview_ws.column_dimensions[column_letter].width = adjusted_width

                # Add formatting to header row
                for cell in overview_ws[1]:
                    cell.font = self.bold_font
                    cell.fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")  # Light green
                
                # Color code the status cells
                for row_idx, row in enumerate(comparison_df.itertuples(index=False), start=2):
                    status_cell = overview_ws[f'C{row_idx}']  # Status column
                    if 'No differences found' in str(row.status):
                        status_cell.fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")  # Light green
                    elif 'differences found' in str(row.status):
                        status_cell.fill = PatternFill(start_color="FFFFAAAA", end_color="FFFFAAAA", fill_type="solid")  # Light red

            # Write detailed differences summary (if any)
            if not summary_df.empty:
                summary_df.to_excel(writer, sheet_name='Detailed Differences', index=False)

                # Get the worksheet for detailed differences
                detail_ws = writer.sheets['Detailed Differences']

                # Auto-adjust column widths for detailed differences
                for column in detail_ws.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except Exception:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    detail_ws.column_dimensions[column_letter].width = adjusted_width

                # Add formatting to header row
                for cell in detail_ws[1]:
                    cell.font = self.bold_font
                    cell.fill = PatternFill(start_color="FFCCCCCC", end_color="FFCCCCCC", fill_type="solid")

        logger.info(f"Summary report created: {summary_path}")

        # Create statistics summary
        self.create_statistics_summary(summary_df, comparison_df)

    def create_statistics_summary(self, summary_df, comparison_df):
        """
        Create a statistics summary of the comparison results.

        Args:
            summary_df (pd.DataFrame): Summary DataFrame with all differences
            comparison_df (pd.DataFrame): Comparison overview DataFrame
        """
        stats_path = os.path.join(self.output_folder, "comparison_statistics.txt")

        with open(stats_path, 'w') as f:
            f.write("COREP Comparison Statistics\n")
            f.write("=" * 50 + "\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

            # Overall comparison stats
            total_files_compared = len(comparison_df) if not comparison_df.empty else 0
            files_with_differences = len(comparison_df[comparison_df['total_differences'] > 0]) if not comparison_df.empty else 0
            files_without_differences = total_files_compared - files_with_differences
            
            f.write(f"Files Compared: {total_files_compared}\n")
            f.write(f"Files with Differences: {files_with_differences}\n")
            f.write(f"Files with No Differences: {files_without_differences}\n\n")

            # Detailed comparison results
            f.write("Comparison Results by File:\n")
            f.write("-" * 40 + "\n")
            if not comparison_df.empty:
                for _, row in comparison_df.iterrows():
                    f.write(f"File: {row['new_file']}\n")
                    f.write(f"  Status: {row['status']}\n")
                    f.write(f"  Sheets Compared: {row['sheets_compared']}\n")
                    f.write(f"  Sheets with Differences: {row['sheets_with_differences']}\n")
                    f.write(f"  Total Differences: {row['total_differences']}\n\n")

            # Overall differences stats
            total_differences = len(summary_df) if not summary_df.empty else 0
            f.write(f"Total Differences Found: {total_differences}\n\n")

            if total_differences > 0:
                # File-wise breakdown
                f.write("File-wise Breakdown:\n")
                f.write("-" * 30 + "\n")
                file_counts = summary_df['file'].value_counts()
                for file_name, count in file_counts.items():
                    f.write(f"{file_name}: {count} differences\n")

                # Sheet-wise breakdown
                f.write("\nSheet-wise Breakdown:\n")
                f.write("-" * 30 + "\n")
                sheet_counts = summary_df['sheet'].value_counts()
                for sheet_name, count in sheet_counts.items():
                    f.write(f"{sheet_name}: {count} differences\n")

                # Largest differences
                f.write("\nTop 10 Largest Differences:\n")
                f.write("-" * 30 + "\n")
                numeric_diffs = summary_df[summary_df['difference'] != 'N/A'].copy()
                if not numeric_diffs.empty:
                    numeric_diffs['abs_difference'] = pd.to_numeric(
                        numeric_diffs['difference'], errors='coerce'
                    ).abs()
                    top_diffs = numeric_diffs.nlargest(10, 'abs_difference')
                    for _, row in top_diffs.iterrows():
                        f.write(
                            f"File: {row['file']}, Sheet: {row['sheet']}, "
                            f"Cell: {row['column_letter']}{row['row']}, "
                            f"Difference: {row['difference']}\n"
                        )
            else:
                f.write("✅ No differences found in any compared files!\n")

        logger.info(f"Statistics summary created: {stats_path}")

    def run_comparison(self):
        """
        Run the complete comparison process.
        """
        logger.info("Starting Excel file comparison process")

        # Find matching files
        matching_pairs = self.find_matching_files()

        if not matching_pairs:
            logger.warning("No matching file pairs found")
            return

        # Compare each pair
        for new_file, prev_file in matching_pairs:
            differences = self.compare_files(new_file, prev_file)
            self.summary_data.extend(differences)

        # Create summary report
        self.create_summary_report()

        logger.info(f"Comparison complete. Found {len(self.summary_data)} total differences.")
        logger.info(f"Results saved in: {self.output_folder}")

        # Create/update comparison index
        self.update_comparison_index()

    def update_comparison_index(self):
        """
        Create or update an index file that tracks all comparison runs.
        """
        try:
            index_path = os.path.join(os.path.dirname(self.output_folder), "comparison_index.txt")
            
            # Create index entry for this run
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            folder_name = os.path.basename(self.output_folder)
            total_differences = len(self.summary_data)
            total_files = len(self.comparison_summary)
            files_with_diffs = len([c for c in self.comparison_summary if c['total_differences'] > 0])
            
            index_entry = (f"{timestamp} | {folder_name} | "
                          f"Files: {total_files} | Differences: {total_differences} | "
                          f"Files with diffs: {files_with_diffs}\n")
            
            # Append to index file
            with open(index_path, 'a', encoding='utf-8') as f:
                if not os.path.exists(index_path) or os.path.getsize(index_path) == 0:
                    f.write("COREP Comparison Index\n")
                    f.write("=" * 80 + "\n")
                    f.write("Timestamp           | Folder               | Files | Diffs | Files w/ Diffs\n")
                    f.write("-" * 80 + "\n")
                f.write(index_entry)
            
            logger.info(f"Updated comparison index: {index_path}")
            
        except Exception as e:
            logger.warning(f"Could not update comparison index: {str(e)}")


def main():
    """
    Main function to run the Excel comparison tool.
    """
    # Define folder paths
    new_folder = "new"
    prev_folder = "prev"
    output_base_folder = "comparison_results"

    # Create comparer instance (will create timestamped subfolder automatically)
    comparer = ExcelComparer(new_folder, prev_folder, output_base_folder)

    # Run comparison
    comparer.run_comparison()

    print("\n" + "="*60)
    print("COREP Excel Comparison Tool - Execution Complete")
    print("="*60)
    print(f"Results saved in: {comparer.output_folder}")
    print("Generated files:")
    print("- comparison_summary.xlsx: Detailed differences report with overview")
    print("- comparison_statistics.txt: Summary statistics")
    print("- highlighted_*.xls: Original files with differences highlighted in red")
    print("="*60)
    print(f"Folder structure: {output_base_folder}/comparison_YYYYMMDD_HHMMSS/")
    print("="*60)


if __name__ == "__main__":
    main()
