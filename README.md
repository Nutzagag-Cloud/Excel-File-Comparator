# Excel File Comparator - Modular Version

A powerful, modular Python tool for comparing Excel files between two folders with advanced features Wahahaha.

## 📁 Project Structure

```
file_comparator_project/
│
├── main.py                    # ⭐ USER CONFIG - Edit this file only!
├── file_reader.py             # Reads Excel files and preprocesses data
├── comparison_engine.py       # Core comparison logic (values, types, strings)
├── file_comparison.py         # File-level comparison orchestration
├── column_processor.py        # Handles blank columns & column alignment
├── report_generator.py        # Generates Excel reports
└── README.md                  # This file
```

## 🚀 Quick Start

1. **Edit `main.py`** - Only file you need to touch!
   ```python
   # Set your folder paths
   Folder_A = r"path\to\source_of_truth"
   Folder_B = r"path\to\files_to_check"
   
   # Configure options
   STRIP_WHITESPACE = True
   DELETE_BLANK_COLUMNS = True
   AUTO_ALIGN_COLUMNS = True
   ```

2. **Install dependencies:**
   ```bash
   pip install pandas openpyxl
   ```

3. **Run:**
   ```bash
   python main.py
   ```

## ⚙️ Configuration Options

### Basic Options

| Option | Default | Description |
|--------|---------|-------------|
| `STRIP_WHITESPACE` | `True` | Remove leading/trailing spaces before comparison |
| `CASE_SENSITIVE` | `True` | Whether "Hello" ≠ "hello" |
| `NUMERIC_TOLERANCE` | `None` | Allow small number differences (e.g., `0.01`) |
| `COMPARE_FORMULAS` | `False` | Compare Excel formulas instead of values |
| `SUMMARY_ONLY` | `False` | Generate only summary counts (faster) |

### Advanced Column Handling ⭐ NEW

| Option | Default | Description |
|--------|---------|-------------|
| `DELETE_BLANK_COLUMNS` | `True` | Remove completely empty columns before comparison |
| `AUTO_ALIGN_COLUMNS` | `True` | Automatically reorder Folder B columns to match Folder A |

## 📊 Features Explained

### 1. DELETE_BLANK_COLUMNS

**Problem:**
```
Folder A: [Name, Age, City]
Folder B: [Name, <BLANK>, Age, City]
```

**Solution:**
```python
DELETE_BLANK_COLUMNS = True
```

**Result:**
- Removes blank columns from both folders
- Compares actual data: `[Name, Age, City]` vs `[Name, Age, City]`
- Logs warning about removed columns
- No false column mismatch errors!

### 2. AUTO_ALIGN_COLUMNS

**Problem:**
```
Folder A: [Name, Age, City]
Folder B: [City, Name, Age]  # Same columns, wrong order!
```

**Solution:**
```python
AUTO_ALIGN_COLUMNS = True
```

**Result:**
- Reorders Folder B to match Folder A: `[Name, Age, City]`
- Compares data correctly aligned
- Logs warning about reordering
- Focuses on actual data differences, not column order!

### 3. Combined Power

```python
DELETE_BLANK_COLUMNS = True
AUTO_ALIGN_COLUMNS = True
```

**Handles messy data:**
```
Folder A: [ID, Name, Age, City]
Folder B: [City, <BLANK>, Name, ID, Age]
```

**Processing:**
1. Removes blank column from B: `[City, Name, ID, Age]`
2. Reorders to match A: `[ID, Name, Age, City]`
3. Compares clean data ✓

## 📋 Output

The tool generates an Excel file with two sheets:

### Summary Sheet
| File Name | Total Differences | Shape Mismatches | Column Mismatches | Value Mismatches | Status |
|-----------|-------------------|------------------|-------------------|------------------|--------|
| file1.xlsx | 0 | 0 | 0 | 0 | ✓ MATCH |
| file2.xlsx | 5 | 0 | 0 | 5 | ✗ DIFFER |

### Detailed Differences Sheet (if `SUMMARY_ONLY = False`)
| File Name | Row | Column | Issue Type | Folder A Value | Folder B Value | Details |
|-----------|-----|--------|------------|----------------|----------------|---------|
| file2.xlsx | 3 | Age | Value Mismatch | 25 | 26 | Values differ |
| file2.xlsx | 5 | Name | Value Mismatch | John | Jon | Difference at position 3 |

## 🎯 Common Use Cases

### Use Case 1: Data Validation
```python
STRIP_WHITESPACE = True       # Handle spaces
DELETE_BLANK_COLUMNS = True   # Remove empty columns
AUTO_ALIGN_COLUMNS = True     # Fix column order
CASE_SENSITIVE = True         # Keep strict
SUMMARY_ONLY = False          # See all details
```

### Use Case 2: Quick Health Check
```python
SUMMARY_ONLY = True           # Just show counts
DELETE_BLANK_COLUMNS = True   # Ignore blanks
AUTO_ALIGN_COLUMNS = True     # Auto-fix order
```

### Use Case 3: Formula Validation
```python
COMPARE_FORMULAS = True       # Check formulas
DELETE_BLANK_COLUMNS = True   # Clean data
```

### Use Case 4: Ultra-Strict Mode
```python
STRIP_WHITESPACE = False      # Spaces matter
CASE_SENSITIVE = True         # Case matters
DELETE_BLANK_COLUMNS = False  # Flag everything
AUTO_ALIGN_COLUMNS = False    # No auto-fixes
NUMERIC_TOLERANCE = None      # Exact numbers
```

## 🔧 Module Descriptions

### main.py ⭐ 
**What users edit.** Contains all configuration options and runs the comparison.

### file_reader.py
Reads Excel files, extracts formulas, strips whitespace, removes blank columns.

### comparison_engine.py
Core comparison logic - determines if two values match based on configured rules.

### file_comparison.py
Orchestrates file-level comparison, tracks differences, updates summary counts.

### column_processor.py
Handles column operations:
- Detects and removes blank columns
- Aligns/reorders columns to match reference
- Provides column structure analysis

### report_generator.py
Creates formatted Excel reports with summary and detailed sheets.

## 💡 Tips

1. **Start with relaxed settings** to see if data matches:
   ```python
   DELETE_BLANK_COLUMNS = True
   AUTO_ALIGN_COLUMNS = True
   STRIP_WHITESPACE = True
   ```

2. **Use SUMMARY_ONLY for large datasets** to get quick overview

3. **Check alignment warnings** - if columns were reordered, review the changes

4. **Gradually tighten rules** once structure is fixed:
   ```python
   # After alignment is good, check exact values
   NUMERIC_TOLERANCE = None
   CASE_SENSITIVE = True
   ```

## 📝 Example Output

```
======================================================================
EXCEL FILE COMPARATOR
======================================================================
Folder A (Source of Truth): C:\You
Folder B (To Compare):      C:\Your enemies
Delete Blank Columns:       YES
Auto-Align Columns:         YES

======================================================================
STEP 1: Reading Folder A files (Source of Truth)
======================================================================
✓ file1.xlsx
  → Removed 1 blank column(s): ['Unnamed: 5']
  - Columns: ['ID', 'Name', 'Coin', 'City']
  - Shape: 100 rows × 4 columns

======================================================================
STEP 2: Comparing Folder B files against Folder A
======================================================================

Comparing: file1.xlsx
  → Removed 2 blank column(s): ['BLANK', 'Unnamed: 7']
  → Reordered columns to match Folder A
⚠️  file1.xlsx: Columns were auto-aligned (reordered to match Folder A)
  ✓ Perfect match! No differences found.

======================================================================
COLUMN ALIGNMENT WARNINGS
======================================================================
⚠️  file1.xlsx: Columns were auto-aligned (reordered to match Folder A)

Note: These files had columns that were reordered or adjusted.
The data comparison proceeded with aligned columns.

✓ Report saved: C:\comparison_report.xlsx
✓ Total differences logged: 0

======================================================================
COMPARISON COMPLETE!
======================================================================
```

## 🤝 Contributing

To add new features, create a new module file and import functions in `main.py`.

## 📄 License - No We do not have any License it open source and Freeeeeee

Free to use and modify for your needs! Please contract Na.sroysamutr@gmail.com if you need Help Thank!!!
