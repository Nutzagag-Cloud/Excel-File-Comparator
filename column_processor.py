"""
Column Processor Module
Handles blank column removal and column alignment/reordering
"""

import pandas as pd


def remove_blank_columns(df, filename=""):
    """
    Remove columns that are completely blank (all NaN or empty strings)
    
    Args:
        df: DataFrame to process
        filename: Name of file (for logging)
    
    Returns:
        tuple: (cleaned_df, removed_columns_list)
    """
    removed_columns = []
    
    for col in df.columns:
        # Check if column is completely blank
        if df[col].isna().all() or (df[col].astype(str).str.strip() == '').all():
            removed_columns.append(col)
    
    if removed_columns:
        df_cleaned = df.drop(columns=removed_columns)
        print(f"  → Removed {len(removed_columns)} blank column(s) from {filename}: {removed_columns}")
        return df_cleaned, removed_columns
    
    return df, []


def align_columns_to_reference(df_reference, df_to_align, filename=""):
    """
    Reorder columns in df_to_align to match df_reference column order
    Also handles missing/extra columns
    
    Args:
        df_reference: DataFrame with correct column order (Folder A)
        df_to_align: DataFrame to reorder (Folder B)
        filename: Name of file (for logging)
    
    Returns:
        tuple: (aligned_df, alignment_info_dict)
    """
    alignment_info = {
        'was_aligned': False,
        'original_order': list(df_to_align.columns),
        'reference_order': list(df_reference.columns),
        'missing_in_target': [],
        'extra_in_target': [],
        'reordered': False
    }
    
    reference_cols = list(df_reference.columns)
    target_cols = list(df_to_align.columns)
    
    # Check if columns are already in the same order
    if reference_cols == target_cols:
        return df_to_align, alignment_info
    
    # Check if same columns but different order
    if set(reference_cols) == set(target_cols):
        # Reorder to match reference
        df_aligned = df_to_align[reference_cols]
        alignment_info['was_aligned'] = True
        alignment_info['reordered'] = True
        print(f"  → Reordered columns in {filename} to match Folder A")
        return df_aligned, alignment_info
    
    # Find missing and extra columns
    missing_cols = [col for col in reference_cols if col not in target_cols]
    extra_cols = [col for col in target_cols if col not in reference_cols]
    
    alignment_info['missing_in_target'] = missing_cols
    alignment_info['extra_in_target'] = extra_cols
    
    # Create aligned dataframe
    df_aligned = pd.DataFrame()
    
    # Add columns from reference in the correct order
    for col in reference_cols:
        if col in df_to_align.columns:
            df_aligned[col] = df_to_align[col]
        else:
            # Column missing in target - add as NaN
            df_aligned[col] = pd.NA
    
    # Note: Extra columns in target are dropped
    
    if missing_cols or extra_cols:
        alignment_info['was_aligned'] = True
        print(f"  → Aligned columns in {filename}:")
        if missing_cols:
            print(f"    - Missing in Folder B (added as empty): {missing_cols}")
        if extra_cols:
            print(f"    - Extra in Folder B (removed): {extra_cols}")
    
    return df_aligned, alignment_info


def get_column_mapping(df_A_cols, df_B_cols):
    """
    Get mapping of how columns differ between two dataframes
    
    Returns:
        dict with analysis of column differences
    """
    mapping = {
        'exact_match': list(df_A_cols) == list(df_B_cols),
        'same_columns_different_order': set(df_A_cols) == set(df_B_cols) and list(df_A_cols) != list(df_B_cols),
        'missing_in_B': [col for col in df_A_cols if col not in df_B_cols],
        'extra_in_B': [col for col in df_B_cols if col not in df_A_cols],
        'position_changes': []
    }
    
    # Find position changes for common columns
    if mapping['same_columns_different_order']:
        for i, col in enumerate(df_A_cols):
            if col in df_B_cols:
                pos_in_B = list(df_B_cols).index(col)
                if i != pos_in_B:
                    mapping['position_changes'].append({
                        'column': col,
                        'position_in_A': i,
                        'position_in_B': pos_in_B
                    })
    
    return mapping


def analyze_column_structure(df_A, df_B, filename):
    """
    Analyze and report on column structure differences
    
    Returns:
        dict with detailed analysis
    """
    analysis = {
        'filename': filename,
        'folder_A_columns': list(df_A.columns),
        'folder_B_columns': list(df_B.columns),
        'blank_columns_A': [],
        'blank_columns_B': [],
        'mapping': get_column_mapping(df_A.columns, df_B.columns)
    }
    
    # Find blank columns
    for col in df_A.columns:
        if df_A[col].isna().all() or (df_A[col].astype(str).str.strip() == '').all():
            analysis['blank_columns_A'].append(col)
    
    for col in df_B.columns:
        if df_B[col].isna().all() or (df_B[col].astype(str).str.strip() == '').all():
            analysis['blank_columns_B'].append(col)
    
    return analysis
