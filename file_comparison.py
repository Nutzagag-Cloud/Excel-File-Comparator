"""
File Comparison Module
Handles comparison of individual files (dataframes)
"""

from comparison_engine import values_match, analyze_mismatch, format_value


def compare_single_file(df_A, df_B, filename, formulas_A=None, formulas_B=None, 
                       config=None, results_list=None, summary_dict=None):
    """
    Compare two dataframes (representing same file from Folder A and Folder B)
    
    Args:
        df_A: DataFrame from Folder A (correct)
        df_B: DataFrame from Folder B (to check)
        filename: Name of the file being compared
        formulas_A: Formulas from Folder A (if compare_formulas enabled)
        formulas_B: Formulas from Folder B (if compare_formulas enabled)
        config: Configuration dict with comparison settings
        results_list: List to append detailed results to
        summary_dict: Dict to store summary counts
    """
    if config is None:
        config = {}
    
    differences_found = 0
    
    # 1. SHAPE CHECK
    if df_A.shape != df_B.shape:
        add_result(results_list, filename, 'N/A', 'N/A', 'Shape Mismatch',
                   f"{df_A.shape[0]} rows × {df_A.shape[1]} cols",
                   f"{df_B.shape[0]} rows × {df_B.shape[1]} cols",
                   'File dimensions do not match', config.get('summary_only', False))
        summary_dict['shape_mismatch'] += 1
        differences_found += 1
    
    # 2. COLUMN COUNT CHECK
    if len(df_A.columns) != len(df_B.columns):
        add_result(results_list, filename, 'Header', 'N/A', 'Column Count Mismatch',
                   f"{len(df_A.columns)} columns",
                   f"{len(df_B.columns)} columns",
                   'Number of columns differs', config.get('summary_only', False))
        summary_dict['column_mismatch'] += 1
        differences_found += 1
    
    # 3. COLUMN NAMES CHECK
    max_cols = max(len(df_A.columns), len(df_B.columns))
    for col_idx in range(max_cols):
        col_A = df_A.columns[col_idx] if col_idx < len(df_A.columns) else '[MISSING]'
        col_B = df_B.columns[col_idx] if col_idx < len(df_B.columns) else '[MISSING]'
        
        if str(col_A) != str(col_B):
            from comparison_engine import find_string_difference
            add_result(results_list, filename, 'Header', f'Column {col_idx + 1}', 
                      'Column Name Mismatch', str(col_A), str(col_B),
                      find_string_difference(str(col_A), str(col_B), config.get('case_sensitive', True)),
                      config.get('summary_only', False))
            summary_dict['column_mismatch'] += 1
            differences_found += 1
    
    # 4. ROW COUNT CHECK
    if len(df_A) != len(df_B):
        add_result(results_list, filename, 'N/A', 'N/A', 'Row Count Mismatch',
                   f"{len(df_A)} rows", f"{len(df_B)} rows",
                   'Number of data rows differs', config.get('summary_only', False))
        summary_dict['shape_mismatch'] += 1
        differences_found += 1
    
    # 5. CELL-BY-CELL COMPARISON
    min_rows = min(len(df_A), len(df_B))
    min_cols = min(len(df_A.columns), len(df_B.columns))
    
    for row_idx in range(min_rows):
        for col_idx in range(min_cols):
            col_name = df_A.columns[col_idx]
            val_A = df_A.iloc[row_idx, col_idx]
            val_B = df_B.iloc[row_idx, col_idx]
            
            # Check formulas if enabled
            if config.get('compare_formulas', False) and formulas_A and formulas_B:
                formula_A = formulas_A.get((row_idx, col_idx))
                formula_B = formulas_B.get((row_idx, col_idx))
                
                if formula_A != formula_B:
                    add_result(results_list, filename, row_idx + 2, col_name, 'Formula Mismatch',
                              formula_A or '[No formula]', formula_B or '[No formula]',
                              'Excel formulas are different', config.get('summary_only', False))
                    summary_dict['formula_mismatch'] += 1
                    differences_found += 1
                    continue
            
            # Check if values match
            if not values_match(val_A, val_B, 
                              case_sensitive=config.get('case_sensitive', True),
                              numeric_tolerance=config.get('numeric_tolerance')):
                issue_type, details = analyze_mismatch(val_A, val_B, 
                                                      numeric_tolerance=config.get('numeric_tolerance'),
                                                      case_sensitive=config.get('case_sensitive', True))
                
                add_result(results_list, filename, row_idx + 2, col_name, issue_type,
                          format_value(val_A), format_value(val_B), details,
                          config.get('summary_only', False))
                
                # Update summary counts
                if 'Type' in issue_type:
                    summary_dict['type_mismatch'] += 1
                elif 'Missing' in issue_type or 'Empty' in issue_type:
                    summary_dict['missing_value'] += 1
                else:
                    summary_dict['value_mismatch'] += 1
                
                differences_found += 1
    
    # Update total
    summary_dict['total_differences'] = differences_found
    
    if differences_found == 0:
        print(f"  ✓ Perfect match! No differences found.")
    else:
        print(f"  ✗ Found {differences_found} differences")
    
    return differences_found


def add_result(results_list, filename, row, column, issue_type, val_A, val_B, details, summary_only):
    """Add a comparison result (only if not in summary_only mode)"""
    if not summary_only and results_list is not None:
        results_list.append({
            'File Name': filename,
            'Row': row,
            'Column': column,
            'Issue Type': issue_type,
            'Folder A Value (Correct)': val_A,
            'Folder B Value (Incorrect)': val_B,
            'Details': details
        })
