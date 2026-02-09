"""
EXCEL FILE COMPARATOR - MAIN CONFIGURATION
===========================================

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
Folder_A = r"path\to\Folder_A"      # Source of truth
Folder_B = r"path\to\Folder_B"    # To compare
output_path = r"Path\to\Output"       # Save location

# ========== COMPARISON OPTIONS ==========
STRIP_WHITESPACE = True      # True = Ignore leading/trailing spaces
CASE_SENSITIVE = True        # True = Case matters ("Hello" ≠ "hello")
NUMERIC_TOLERANCE = None     # None = Exact match | 0.01 = Allow ±0.01 difference
COMPARE_FORMULAS = False     # True = Compare Excel formulas (=SUM...)
SUMMARY_ONLY = False         # True = Only counts | False = Full details

# ========== ADVANCED COLUMN HANDLING ==========
DELETE_BLANK_COLUMNS = True  # True = Remove completely empty columns before comparison
                             # False = Keep all columns (will flag blank column differences)

AUTO_ALIGN_COLUMNS = True    # True = Reorder Folder B columns to match Folder A
                             # False = Compare columns in their original order
                             # Note: Only works if same columns exist (just different order)

# ========== MULTI-SHEET HANDLING ==========
COMPARE_ALL_SHEETS = True    # True = Compare all sheets in each Excel file
                             # False = Compare only specific sheets (see SPECIFIC_SHEETS)

SPECIFIC_SHEETS = None       # None = Use COMPARE_ALL_SHEETS setting
                             # ["Sheet1", "Data", "Summary"] = Only compare these sheets
                             # Note: Only used if COMPARE_ALL_SHEETS = False

# ========== OUTPUT FILE NAME ==========
OUTPUT_FILENAME = "Folder_A_vs_Folder_B_comparison.xlsx"


# ============================================================================
# COMPARISON LOGIC - DON'T EDIT BELOW THIS LINE
# ============================================================================

def run_comparison():
    """Main function to run the comparison"""
    
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
    alignment_warnings = []  # Track column alignment warnings
    sheet_structure_warnings = []  # Track sheet structure warnings
    
    # Print header
    print("\n" + "=" * 70)
    print("EXCEL FILE COMPARATOR")
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
        
        # Initialize summary for this file
        summary[filename] = {
            'total_differences': 0,
            'shape_mismatch': 0,
            'column_mismatch': 0,
            'value_mismatch': 0,
            'type_mismatch': 0,
            'formula_mismatch': 0,
            'missing_value': 0
        }
        
        # Check if corresponding Folder A file exists
        if filename not in folder_a_data:
            if not SUMMARY_ONLY:
                results.append({
                    'File Name': filename,
                    'Row': 'N/A',
                    'Column': 'N/A',
                    'Issue Type': 'Missing Folder A File',
                    'Folder A Value (Correct)': 'File not found in Folder A',
                    'Folder B Value (Incorrect)': 'File exists in Folder B',
                    'Details': 'This Folder B file has no corresponding Folder A file'
                })
            summary[filename]['total_differences'] += 1
            print(f"No matching Folder A file found!")
            continue
        
        # Get Folder A data
        folder_a_file_data = folder_a_data[filename]
        df_A = folder_a_file_data['dataframe']
        formulas_A = folder_a_file_data.get('formulas')
        
        # Read Folder B file
        try:
            if COMPARE_FORMULAS:
                df_B, formulas_B = read_excel_with_formulas(folder_b_path / filename)
            else:
                df_B = pd.read_excel(folder_b_path / filename)
                formulas_B = None
            
            # Strip whitespace if enabled
            if STRIP_WHITESPACE:
                df_B = strip_dataframe_whitespace(df_B)
            
            # Remove blank columns if enabled
            removed_cols_B = []
            if DELETE_BLANK_COLUMNS:
                df_B, removed_cols_B = remove_blank_columns(df_B, filename)
            
            # Auto-align columns if enabled
            alignment_info = None
            if AUTO_ALIGN_COLUMNS:
                df_B, alignment_info = align_columns_to_reference(df_A, df_B, filename)
                
                # Add warning if columns were aligned
                if alignment_info['was_aligned']:
                    warning_msg = f"⚠️  {filename}: Columns were auto-aligned"
                    if alignment_info.get('reordered'):
                        warning_msg += " (reordered to match Folder A)"
                    if alignment_info.get('missing_in_target'):
                        warning_msg += f" | Missing: {alignment_info['missing_in_target']}"
                    if alignment_info.get('extra_in_target'):
                        warning_msg += f" | Extra (removed): {alignment_info['extra_in_target']}"
                    
                    alignment_warnings.append({
                        'filename': filename,
                        'message': warning_msg,
                        'details': alignment_info
                    })
                    print(warning_msg)
                
        except Exception as e:
            if not SUMMARY_ONLY:
                results.append({
                    'File Name': filename,
                    'Row': 'N/A',
                    'Column': 'N/A',
                    'Issue Type': 'Folder B File Read Error',
                    'Folder A Value (Correct)': 'N/A',
                    'Folder B Value (Incorrect)': 'Cannot read file',
                    'Details': f'Error: {str(e)}'
                })
            summary[filename]['total_differences'] += 1
            print(f"Error reading Folder B file: {str(e)}")
            continue
        
        # Compare the files
        compare_single_file(
            df_A, df_B, filename,
            formulas_A=formulas_A,
            formulas_B=formulas_B,
            config=config,
            results_list=results,
            summary_dict=summary[filename]
        )
    
    # STEP 3: Generate report
    generate_excel_report(results, summary, output_file_path, summary_only=SUMMARY_ONLY)
    
    # Print alignment warnings if any
    if alignment_warnings:
        print("\n" + "=" * 70)
        print("COLUMN ALIGNMENT WARNINGS")
        print("=" * 70)
        for warning in alignment_warnings:
            print(warning['message'])
        print("\nNote: These files had columns that were reordered or adjusted.")
        print("The data comparison proceeded with aligned columns.")
    
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


# ============================================================================
# RUN THE COMPARISON
# ============================================================================

if __name__ == "__main__":
    run_comparison()


# ============================================================================
# COMMON CONFIGURATIONS - REFERENCE
# ============================================================================

"""
1. STRICT MODE : Catch everything that differs, including whitespace and case
   STRIP_WHITESPACE = False
   CASE_SENSITIVE = True
   NUMERIC_TOLERANCE = None
   → Perfect for exact data validation

2. RELAXED MODE : Ignore formatting and minor numeric differences
   STRIP_WHITESPACE = True
   CASE_SENSITIVE = False
   NUMERIC_TOLERANCE = 0.01
   → Good for comparing data from different systems

3. FORMULA VALIDATION: in case there need to check if Excel formulas are identical, not just results
   COMPARE_FORMULAS = True
   → Check if Excel formulas match, not just results

4. QUICK HEALTH CHECK:
   SUMMARY_ONLY = True
   → Just see which files differ, no details

5. RECOMMENDED : Defult settings to catch common issues
   STRIP_WHITESPACE = True      # Fix your whitespace issue
   CASE_SENSITIVE = True         # Keep case checking
   NUMERIC_TOLERANCE = None      # Exact numbers
   SUMMARY_ONLY = False          # See all details

KEY FEATURES EXPLAINED:

DELETE_BLANK_COLUMNS:
- Scenario: Folder A has columns [Name, Age, City]
           Folder B has columns [Name, BLANK, Age, City]
- With DELETE_BLANK_COLUMNS = True:
  → Removes BLANK column from both folders
  → Compares [Name, Age, City] vs [Name, Age, City]
  → No column mismatch error!
  → Logs warning about removed columns

- With DELETE_BLANK_COLUMNS = False:
  → Compares [Name, Age, City] vs [Name, BLANK, Age, City]
  → Flags column count mismatch
  → Flags position mismatches for Age and City

AUTO_ALIGN_COLUMNS:
- Scenario: Folder A has columns [Name, Age, City]
           Folder B has columns [City, Name, Age]
- With AUTO_ALIGN_COLUMNS = True:
  → Reorders Folder B to [Name, Age, City]
  → Compares data correctly aligned
  → Logs warning about reordering
  → Data comparison proceeds normally

- With AUTO_ALIGN_COLUMNS = False:
  → Compares [Name, Age, City] vs [City, Name, Age]
  → Flags every cell as mismatched (wrong positions)

COMBINED POWER:
DELETE_BLANK_COLUMNS = True + AUTO_ALIGN_COLUMNS = True
→ Perfect for messy data where columns are reordered AND have blanks
→ Focuses comparison on actual data differences, not formatting
"""
