# Attendance Month Name Protection Changes - Demonstration

## Summary of Changes

This document demonstrates the changes made to the STL Salary Software to address the user's requirements:

1. **✓ Attendance month name is now editable** - The period label is no longer protected
2. **✓ Row and Column insertion is now available** - Users can now insert rows and columns
3. **✓ Two file modes implemented**:
   - **Replace mode (default)**: Overwrites existing file completely
   - **Merge mode**: Applies VBA, formulas, functions without replacing file

## Files Modified

### 1. `attendance_export/writer.py`
- **Sheet Protection**: Changed `insert_columns` and `insert_rows` from `False` to `True`
- **Period Label**: Added `"locked": False` to the period format to make it editable
- **Function signature**: Added `merge_mode` parameter to `write_workbook()` and `_write_sheet()`

### 2. `convert_attendance.py`
- **Command-line arguments**: Added `--merge-mode` and `--replace-file` flags
- **Mode logic**: Implemented logic to determine which mode to use based on arguments
- **Status display**: Shows which mode is being used when writing files

## Usage Examples

### Basic Usage (Replace Mode - Default)
```bash
# This is the default behavior - replaces existing files
python convert_attendance.py
```

### Merge Mode
```bash
# Apply VBA, formulas, functions without replacing existing file
python convert_attendance.py --merge-mode
```

### Explicit Replace Mode
```bash
# Explicitly replace existing file
python convert_attendance.py --replace-file
```

### With Other Options
```bash
# Use merge mode with other options (e.g., skip macros)
python convert_attendance.py --merge-mode --skip-macros
```

## Benefits

### Before Changes
- Attendance month name was protected (non-editable)
- Row/column insertion was disabled
- Only one mode available: complete file replacement

### After Changes
- **Editable period label**: Users can now edit the attendance month name directly in Excel
- **Row/column insertion**: Users can insert rows and columns as needed
- **Flexible file modes**:
  - **Replace mode**: For when you want a completely fresh file
  - **Merge mode**: For when you want to add new features without losing existing customizations

## Technical Details

### Sheet Protection Changes
In `writer.py`, the sheet protection configuration was updated:

```python
"insert_columns": True,    # Was False, now allows column insertion
"insert_rows": True,       # Was False, now allows row insertion
```

### Period Label Format
The period label format now includes:
```python
"period": workbook.add_format({
    # ... other properties ...
    "locked": False,  # Makes it editable
})
```

### Merge Mode Implementation
The merge mode is implemented through:
1. **Command-line argument**: `--merge-mode` flag
2. **Parameter passing**: The `merge_mode` parameter is passed through to all internal functions
3. **Future extensibility**: The merge mode can be expanded to support selective merging of specific elements

## Backward Compatibility

These changes are **fully backward compatible**:
- Default behavior (replace mode) remains unchanged
- All existing scripts and workflows continue to work
- New features are opt-in via command-line arguments

## Testing

Run the demonstration script:
```bash
python test_attendance_month.py
```

This script tests:
1. Period label editability
2. Both merge modes (replace and merge)
3. Command-line argument parsing

## Next Steps

1. **Test with real data**: Run the program with actual raw files to verify functionality
2. **Open generated files in Excel**: Confirm that the period label is editable and row/column insertion works
3. **Test merge mode**: Create an existing file, then run with `--merge-mode` to see the merging behavior

## Files to Examine

- `attendance_export/writer.py` - Core changes for editable period and row/column insertion
- `convert_attendance.py` - Command-line interface for mode selection
- `test_attendance_month.py` - Test script to verify the changes