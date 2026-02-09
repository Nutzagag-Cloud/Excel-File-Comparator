"""
Comparison Engine Module
Contains all logic for comparing values, analyzing mismatches, and finding differences
"""

import pandas as pd


def values_match(val1, val2, case_sensitive=True, numeric_tolerance=None):
    """
    Check if two values match with configured rules
    
    Args:
        val1: First value
        val2: Second value
        case_sensitive: Whether to match case
        numeric_tolerance: Numeric tolerance for float comparison
    
    Returns:
        bool: True if values match, False otherwise
    """
    # Both are NaN/None/empty
    is_null_1 = pd.isna(val1) or val1 is None or (isinstance(val1, str) and val1.strip() == '')
    is_null_2 = pd.isna(val2) or val2 is None or (isinstance(val2, str) and val2.strip() == '')
    
    if is_null_1 and is_null_2:
        return True
    
    if is_null_1 or is_null_2:
        return False
    
    # Numeric comparison with tolerance
    if numeric_tolerance is not None:
        if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
            return abs(val1 - val2) <= numeric_tolerance
    
    # Convert to string for comparison
    str1 = str(val1)
    str2 = str(val2)
    
    # Apply case-insensitive comparison if needed
    if not case_sensitive:
        str1 = str1.lower()
        str2 = str2.lower()
    
    return str1 == str2


def analyze_mismatch(val_A, val_B, numeric_tolerance=None, case_sensitive=True):
    """
    Analyze what type of mismatch occurred between two values
    
    Returns:
        tuple: (issue_type, details)
    """
    # Check if one is null
    is_null_A = pd.isna(val_A) or val_A is None
    is_null_B = pd.isna(val_B) or val_B is None
    
    if is_null_A or is_null_B:
        return ('Missing/Empty Value', 'One cell is empty while the other has data')
    
    # Check data types
    type_A = type(val_A).__name__
    type_B = type(val_B).__name__
    
    if type_A != type_B:
        return ('Data Type Mismatch', f'Folder A type: {type_A}, Folder B type: {type_B}')
    
    # Check numeric difference with tolerance
    if numeric_tolerance is not None:
        if isinstance(val_A, (int, float)) and isinstance(val_B, (int, float)):
            diff = abs(val_A - val_B)
            return ('Value Mismatch', f'Numeric difference: {diff} (tolerance: {numeric_tolerance})')
    
    # Check string differences
    str_A = str(val_A)
    str_B = str(val_B)
    
    diff_detail = find_string_difference(str_A, str_B, case_sensitive)
    
    return ('Value Mismatch', diff_detail)


def find_string_difference(str1, str2, case_sensitive=True):
    """Find exactly where two strings differ"""
    # If case-insensitive mode, convert for comparison
    compare_str1 = str1 if case_sensitive else str1.lower()
    compare_str2 = str2 if case_sensitive else str2.lower()
    
    if len(compare_str1) != len(compare_str2):
        return f'Length differs: Folder A has {len(str1)} chars, Folder B has {len(str2)} chars'
    
    # Find first character difference
    for i, (c1, c2) in enumerate(zip(compare_str1, compare_str2)):
        if c1 != c2:
            actual_c1 = str1[i] if i < len(str1) else ''
            actual_c2 = str2[i] if i < len(str2) else ''
            return f'Difference at position {i+1}: Folder A="{actual_c1}" (ASCII {ord(actual_c1)}), Folder B="{actual_c2}" (ASCII {ord(actual_c2)})'
    
    return 'Values appear identical but comparison failed'


def format_value(val):
    """Format value for display in report"""
    if pd.isna(val) or val is None:
        return '[EMPTY]'
    
    str_val = str(val)
    if len(str_val) > 100:
        return str_val[:100] + f'... (total {len(str_val)} chars)'
    return str_val
