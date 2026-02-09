"""
EXCEL FILE COMPARATOR - MAIN CONFIGURATION (MULTI-SHEET VERSION)
================================================================

This is the ONLY file you need to edit!
Just update the settings below and run this file.

All the complex comparison logic is in separate modules.
"""

from pathlib import Path
import pandas as pd

# Import our modules
from file_reader import read_folder_files, read_excel_with_formulas, strip_dataframe_whitespace
from file_comparison import compare_single_file
from report_generator import generate_excel_report
from column_processor import remove_blank_columns, align_columns_to_reference
from sheet_handler import (
    read_all_sheets, 
    read_specific_sheets, 
    read_excel_sheets_with_formulas,
    compare_sheet_structures,
    get_sheet_selection,
    format_sheet_info
)


# ============================================================================
# USER CONFIGURATION - EDIT THIS SECTION ONLY
# ============================================================================

# ========== FOLDER PATHS ==========
Folder_A = r"C:\Users\Nattapong.Sroysamutr\Downloads\ISR"      # Source of truth
Folder_B = r"C:\Users\Nattapong.Sroysamutr\Downloads\DCODE"    # To compare
output_path = r"C:\Users\Nattapong.Sroysamutr\Downloads"       # Save location

# ========== COMPARISON OPTIONS ==========
STRIP_WHITESPACE = True      # True = Ignore leading/trailing spaces
CASE_SENSITIVE = True        # True = Case matters ("Hello" ≠ "hello")
NUMERIC_TOLERANCE = None     # None = Exact match | 0.01 = Allow ±0.01 difference
COMPARE_FORMULAS = False     # True = Compare Excel formulas (=SUM...)
SUMMARY_ONLY = False         # True = Only counts | False = Full details

# ========== ADVANCED COLUMN HANDLING ==========
DELETE_BLANK_COLUMNS = True  # True = Remove completely empty columns before comparison
AUTO_ALIGN_COLUMNS = True    # True = Reorder Folder B columns to match Folder A

# ========== MULTI-SHEET HANDLING ==========
COMPARE_ALL_SHEETS = True    # True = Compare all sheets in each Excel file
                             # False = Compare only specific sheets (see SPECIFIC_SHEETS)

SPECIFIC_SHEETS = None       # None = Use COMPARE_ALL_SHEETS setting
                             # ["Sheet1", "Data"] = Only compare these sheets
                             # Note: Only used if COMPARE_ALL_SHEETS = False

# ========== OUTPUT FILE NAME ==========
OUTPUT_FILENAME = "ISR_vs_DCODE_comparison.xlsx"


# ============================================================================
# COMPARISON LOGIC - DON'T EDIT BELOW THIS LINE
# ============================================================================

def run_comparison():
    """Main function to run the comparison with multi-sheet support"""
    
    # Convert paths
    folder_a_path = Path(Folder_A)
    folder_b_path = Path(Folder_B)
    output_file_path = Path(output_path) / OUTPUT_FILENAME
    
    # Configuration dictionary
    config = {
        'strip_whitespace': STRIP_WHITESPACE,
        'case_sensitive': CASE_SENSITIVE,
        'numeric_tolerance': NUMERIC_TOLERANCE,
        'compare_formulas': COMPARE_FORMULAS,
        'summary_only': SUMMARY_ONLY,
        'delete_blank_columns': DELETE_BLANK_COLUMNS,
        'auto_align_columns': AUTO_ALIGN_COLUMNS,
        'compare_all_sheets': COMPARE_ALL_SHEETS,
        'specific_sheets': SPECIFIC_SHEETS
    }
    
    # Storage for results
    results = []
    summary = {}
    alignment_warnings = []
    sheet_structure_warnings = []
    
    # Print header
    print("\n" + "=" * 70)
    print("EXCEL FILE COMPARATOR (MULTI-SHEET)")
    print("=" * 70)
    print(f"Folder A (Source of Truth): {folder_a_path}")
    print(f"Folder B (To Compare):      {folder_b_path}")
    print(f"Whitespace Stripping:       {'ENABLED' if STRIP_WHITESPACE else 'DISABLED'}")
    print(f"Case Sensitive:             {'YES' if CASE_SENSITIVE else 'NO'}")
    print(f"Numeric Tolerance:          {NUMERIC_TOLERANCE if NUMERIC_TOLERANCE else 'DISABLED'}")
    print(f"Compare Formulas:           {'YES' if COMPARE_FORMULAS else 'NO'}")
    print(f"Summary Only:               {'YES' if SUMMARY_ONLY else 'NO'}")
    print(f"Delete Blank Columns:       {'YES' if DELETE_BLANK_COLUMNS else 'NO'}")
    print(f"Auto-Align Columns:         {'YES' if AUTO_ALIGN_COLUMNS else 'NO'}")
    print(f"Compare All Sheets:         {'YES' if COMPARE_ALL_SHEETS else 'NO'}")
    if not COMPARE_ALL_SHEETS and SPECIFIC_SHEETS:
        print(f"Specific Sheets:            {SPECIFIC_SHEETS}")
    print()
    
    # STEP 1: Read Folder A files
    print("=" * 70)
    print("STEP 1: Reading Folder A files (Source of Truth)")
    print("=" * 70)
    
    folder_a_data = read_folder_files(
        folder_a_path, 
        strip_whitespace=STRIP_WHITESPACE,
        compare_formulas=COMPARE_FORMULAS,
        delete_blank_columns=DELETE_BLANK_COLUMNS,
        compare_all_sheets=COMPARE_ALL_SHEETS,
        specific_sheets=SPECIFIC_SHEETS
    )
    
    print(f"\nTotal Folder A files loaded: {len(folder_a_data)}")
    _print_configuration_summary()
    print()
    
    # STEP 2: Compare Folder B files
    print("=" * 70)
    print("STEP 2: Comparing Folder B files against Folder A")
    print("=" * 70)
    
    # Get Folder B files
    folder_b_files = [f.name for f in folder_b_path.iterdir() 
                      if f.is_file() and f.suffix.lower() in ['.xlsx', '.xls']]
    
    for filename in sorted(folder_b_files):
        print(f"\nComparing: {filename}")
        
        # Check if corresponding Folder A file exists
        if filename not in folder_a_data:
            summary[filename] = {
                'total_differences': 1,
                'shape_mismatch': 0,
                'column_mismatch': 0,
                'value_mismatch': 0,
                'type_mismatch': 0,
                'formula_mismatch': 0,
                'missing_value': 0
            }
            if not SUMMARY_ONLY:
                results.append({
                    'File Name': filename,
                    'Sheet Name': 'N/A',
                    'Row': 'N/A',
                    'Column': 'N/A',
                    'Issue Type': 'Missing Folder A File',
                    'Folder A Value (Correct)': 'File not found in Folder A',
                    'Folder B Value (Incorrect)': 'File exists in Folder B',
                    'Details': 'This Folder B file has no corresponding Folder A file'
                })
            print(f"  ✗ No matching Folder A file found!")
            continue
        
        # Get Folder A sheets
        folder_a_file_data = folder_a_data[filename]
        sheets_A = folder_a_file_data['sheets']
        
        # Read Folder B file sheets
        try:
            if COMPARE_FORMULAS:
                sheets_to_read = None if COMPARE_ALL_SHEETS else SPECIFIC_SHEETS
                sheets_B_raw = read_excel_sheets_with_formulas(folder_b_path / filename, sheets_to_read)
            else:
                if COMPARE_ALL_SHEETS:
                    sheets_B_temp = read_all_sheets(folder_b_path / filename)
                    sheets_B_raw = {name: {'dataframe': df, 'formulas': None} 
                                   for name, df in sheets_B_temp.items()}
                else:
                    sheets_to_read = SPECIFIC_SHEETS if SPECIFIC_SHEETS else list(sheets_A.keys())
                    sheets_B_temp = read_specific_sheets(folder_b_path / filename, sheets_to_read)
                    sheets_B_raw = {name: {'dataframe': df, 'formulas': None} 
                                   for name, df in sheets_B_temp.items()}
            
        except Exception as e:
            summary[filename] = {'total_differences': 1, 'shape_mismatch': 0, 'column_mismatch': 0, 
                               'value_mismatch': 0, 'type_mismatch': 0, 'formula_mismatch': 0, 'missing_value': 0}
            if not SUMMARY_ONLY:
                results.append({
                    'File Name': filename,
                    'Sheet Name': 'N/A',
                    'Row': 'N/A',
                    'Column': 'N/A',
                    'Issue Type': 'Folder B File Read Error',
                    'Folder A Value (Correct)': 'N/A',
                    'Folder B Value (Incorrect)': 'Cannot read file',
                    'Details': f'Error: {str(e)}'
                })
            print(f"  ✗ Error reading Folder B file: {str(e)}")
            continue
        
        # Check sheet structure
        sheet_comparison = compare_sheet_structures(sheets_A, sheets_B_raw, filename)
        
        if not sheet_comparison['sheets_match']:
            warning = f"⚠️  {filename}: Sheet structure differs"
            if sheet_comparison['sheets_only_in_A']:
                warning += f" | Missing in B: {sheet_comparison['sheets_only_in_A']}"
            if sheet_comparison['sheets_only_in_B']:
                warning += f" | Extra in B: {sheet_comparison['sheets_only_in_B']}"
            sheet_structure_warnings.append(warning)
            print(warning)
        
        # Initialize summary for this file
        file_summary = {
            'total_differences': 0,
            'shape_mismatch': 0,
            'column_mismatch': 0,
            'value_mismatch': 0,
            'type_mismatch': 0,
            'formula_mismatch': 0,
            'missing_value': 0
        }
        
        # Compare each common sheet
        for sheet_name in sheet_comparison['common_sheets']:
            print(f"  Sheet: '{sheet_name}'")
            
            # Get sheet data
            sheet_A_data = sheets_A[sheet_name]
            sheet_B_data = sheets_B_raw[sheet_name]
            
            df_A = sheet_A_data['dataframe']
            df_B = sheet_B_data['dataframe']
            formulas_A = sheet_A_data.get('formulas')
            formulas_B = sheet_B_data.get('formulas')
            
            # Strip whitespace if enabled
            if STRIP_WHITESPACE:
                df_B = strip_dataframe_whitespace(df_B)
            
            # Remove blank columns if enabled
            if DELETE_BLANK_COLUMNS:
                df_B, removed_cols_B = remove_blank_columns(df_B, f"{filename} - {sheet_name}")
            
            # Auto-align columns if enabled
            if AUTO_ALIGN_COLUMNS:
                df_B, alignment_info = align_columns_to_reference(df_A, df_B, f"{filename} - {sheet_name}")
                
                if alignment_info['was_aligned']:
                    warning_msg = f"⚠️  {filename} - {sheet_name}: Columns auto-aligned"
                    if alignment_info.get('reordered'):
                        warning_msg += " (reordered)"
                    if alignment_info.get('missing_in_target'):
                        warning_msg += f" | Missing: {alignment_info['missing_in_target']}"
                    if alignment_info.get('extra_in_target'):
                        warning_msg += f" | Extra: {alignment_info['extra_in_target']}"
                    
                    alignment_warnings.append(warning_msg)
                    print(f"    {warning_msg}")
            
            # Create sheet-specific summary
            sheet_summary = {
                'total_differences': 0,
                'shape_mismatch': 0,
                'column_mismatch': 0,
                'value_mismatch': 0,
                'type_mismatch': 0,
                'formula_mismatch': 0,
                'missing_value': 0
            }
            
            # Compare the sheet (with sheet_name in results)
            sheet_results = []
            compare_single_file(
                df_A, df_B, filename,
                formulas_A=formulas_A,
                formulas_B=formulas_B,
                config=config,
                results_list=sheet_results,
                summary_dict=sheet_summary
            )
            
            # Add sheet name to each result
            for result in sheet_results:
                result['Sheet Name'] = sheet_name
                results.append(result)
            
            # Aggregate to file summary
            for key in file_summary:
                file_summary[key] += sheet_summary[key]
        
        summary[filename] = file_summary
    
    # STEP 3: Generate report
    generate_excel_report(results, summary, output_file_path, summary_only=SUMMARY_ONLY)
    
    # Print warnings
    if sheet_structure_warnings:
        print("\n" + "=" * 70)
        print("SHEET STRUCTURE WARNINGS")
        print("=" * 70)
        for warning in sheet_structure_warnings:
            print(warning)
    
    if alignment_warnings:
        print("\n" + "=" * 70)
        print("COLUMN ALIGNMENT WARNINGS")
        print("=" * 70)
        for warning in alignment_warnings:
            print(warning)
    
    print("\n" + "=" * 70)
    print("COMPARISON COMPLETE!")
    print("=" * 70)


def _print_configuration_summary():
    """Print enabled configuration options"""
    if STRIP_WHITESPACE:
        print("Note: Whitespace stripping is ENABLED")
    if not CASE_SENSITIVE:
        print("Note: Case-insensitive comparison is ENABLED")
    if NUMERIC_TOLERANCE is not None:
        print(f"Note: Numeric tolerance is {NUMERIC_TOLERANCE}")
    if COMPARE_FORMULAS:
        print("Note: Formula comparison is ENABLED")
    if SUMMARY_ONLY:
        print("Note: Summary-only mode is ENABLED")
    if DELETE_BLANK_COLUMNS:
        print("Note: Blank column deletion is ENABLED")
    if AUTO_ALIGN_COLUMNS:
        print("Note: Auto column alignment is ENABLED")
    if COMPARE_ALL_SHEETS:
        print("Note: Comparing ALL sheets in each file")
    elif SPECIFIC_SHEETS:
        print(f"Note: Comparing specific sheets: {SPECIFIC_SHEETS}")


# ============================================================================
# RUN THE COMPARISON
# ============================================================================

if __name__ == "__main__":
    run_comparison()


# ============================================================================
# USAGE EXAMPLES - MULTI-SHEET
# ============================================================================

"""
MULTI-SHEET EXAMPLES:

1. Compare ALL sheets in files:
   COMPARE_ALL_SHEETS = True
   SPECIFIC_SHEETS = None

2. Compare only specific sheets:
   COMPARE_ALL_SHEETS = False
   SPECIFIC_SHEETS = ["Sheet1", "Data", "Summary"]

3. Compare first sheet only (default Excel behavior):
   COMPARE_ALL_SHEETS = False
   SPECIFIC_SHEETS = None  # Will compare only first sheet

EXAMPLE SCENARIOS:

Scenario A: Files with multiple data sheets
   File structure: Sales.xlsx has sheets ["Q1", "Q2", "Q3", "Q4"]
   Setting: COMPARE_ALL_SHEETS = True
   Result: Compares all 4 quarters

Scenario B: Files with data + documentation sheets
   File structure: Report.xlsx has sheets ["Data", "Charts", "README"]
   Setting: COMPARE_ALL_SHEETS = False
           SPECIFIC_SHEETS = ["Data", "Charts"]
   Result: Compares only Data and Charts, ignores README

Scenario C: Single sheet files (like before)
   File structure: Simple.xlsx has sheet ["Sheet1"]
   Setting: COMPARE_ALL_SHEETS = True
   Result: Compares Sheet1 (works same as before)
"""
