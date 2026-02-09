"""
Report Generator Module
Handles creation of Excel reports (summary and detailed)
"""

import pandas as pd
from openpyxl.styles import PatternFill, Font


def generate_excel_report(results_list, summary_dict, output_path, summary_only=False):
    """
    Generate Excel report with summary and detailed sheets
    
    Args:
        results_list: List of detailed comparison results
        summary_dict: Dictionary with summary counts per file
        output_path: Full path to output Excel file
        summary_only: Whether to generate only summary
    """
    print("\n" + "=" * 70)
    print("STEP 3: Generating Excel Report")
    print("=" * 70)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Always generate summary sheet
        generate_summary_sheet(writer, summary_dict)
        
        # Generate detailed sheet only if not summary_only mode
        if not summary_only and results_list:
            generate_detailed_sheet(writer, results_list)
    
    print(f"\n✓ Report saved: {output_path}")
    
    if summary_only:
        print(f"✓ Summary report generated for {len(summary_dict)} files")
    else:
        print(f"✓ Total differences logged: {len(results_list)}")


def generate_summary_sheet(writer, summary_dict):
    """Generate summary sheet with counts per file"""
    summary_data = []
    
    for filename, counts in summary_dict.items():
        summary_data.append({
            'File Name': filename,
            'Total Differences': counts['total_differences'],
            'Shape Mismatches': counts['shape_mismatch'],
            'Column Mismatches': counts['column_mismatch'],
            'Value Mismatches': counts['value_mismatch'],
            'Type Mismatches': counts['type_mismatch'],
            'Formula Mismatches': counts['formula_mismatch'],
            'Missing Values': counts['missing_value'],
            'Status': '✓ MATCH' if counts['total_differences'] == 0 else '✗ DIFFER'
        })
    
    df_summary = pd.DataFrame(summary_data)
    df_summary.to_excel(writer, sheet_name='Summary', index=False)
    
    # Format summary sheet
    worksheet = writer.sheets['Summary']
    format_worksheet(worksheet, header_color='1F4E78')


def generate_detailed_sheet(writer, results_list):
    """Generate detailed differences sheet"""
    df_report = pd.DataFrame(results_list)
    df_report = df_report.sort_values(['File Name', 'Row', 'Column'])
    df_report.to_excel(writer, sheet_name='Detailed Differences', index=False)
    
    # Format detailed sheet
    worksheet = writer.sheets['Detailed Differences']
    format_worksheet(worksheet, header_color='1F4E78')


def format_worksheet(worksheet, header_color='1F4E78'):
    """Apply formatting to worksheet"""
    # Format header row
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        
        adjusted_width = min(max_length + 3, 60)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Freeze header row
    worksheet.freeze_panes = 'A2'
