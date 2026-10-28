# convert_attendance.py

Convert Matrix jobcard raw files into the Spring Trade attendance workbook.

## Overview

This Python program reads raw Matrix jobcard HTML files and generates an Excel attendance workbook with:
- Attendance data and formulas
- Sub-section filtering via dropdown
- VBA macros (optional) for duty transfer, print-time column hiding, and footer signatures
- Professional formatting for payroll reporting

## Prerequisites

### Python Packages
Install required packages via pip:
```bash
pip install xlsxwriter openpyxl pandas pythoncom pywin32 beautifulsoup4 lxml
```

### Microsoft Excel
- **Required for VBA macros** - the program uses `win32com.client` to add VBA projects
- The workbook functions without Excel, but macros won't be added if Excel/VBA access is unavailable
- Enable "Trust access to the VBA project object model" in Excel > Options > Trust Center if needed

## Repository Structure

```
.
├── convert_attendance.py    # Main script
├── run.bat                  # Double-click launcher (Windows)
├── raw_files/               # Place raw jobcard HTML files here
├── attendance_export/       # Internal package (do not modify)
│   ├── __init__.py
│   ├── classify.py
│   ├── constants.py
│   ├── parser.py
│   ├── vba.py
│   └── writer.py
├── output/                  # Generated output files
├── nevertouch/              # Company logo (embedded at build time)
├── Format.xls               # Reference format file
└── requirements.txt         # Python dependencies
```

## Running the Program

### One-Click Run (Windows)

Double-click **`run.bat`** in the project folder. It:

- always runs from the project folder (works even with "Run as administrator")
- finds Python automatically (`python`, falling back to `py`)
- runs the full conversion and prints the result summary in the console
- lists the newest files in `output/` and stays open ("Press any key") so
  nothing is missed

Command-line options pass through, e.g. `run.bat --skip-macros`.

> Tip: close any `Salary_Attendance_*.xlsm` you have open in Excel before
> running — Excel locks the file and the macro step can't overwrite it while
> it is open (the .xlsx is still regenerated).

### Basic Usage (with VBA macros)
```bash
python convert_attendance.py
```
This generates `output/Salary_Attendance_<Month>_<Year>.xlsm` with all macros.

### Without VBA macros
```bash
python convert_attendance.py --skip-macros
```
Generates a regular `.xlsx` file without VBA. Runs without Excel installed.

### With custom output path
```bash
python convert_attendance.py --output /path/to/my_output/
```

### With custom raw directory
```bash
python convert_attendance.py --raw-dir /path/to/my_raw_files/
```

## Command-Line Arguments

| Argument | Description | Default |
|----------|-------------|---------|
| `--raw-dir` | Directory containing raw jobcard files | `raw_files/` |
| `--output` | Output file or folder path | `output/` |
| `--skip-macros` | Skip VBA macro generation | (run with macros) |

## Input Files

Place your raw Matrix jobcard HTML export files in the `raw_files/` directory. The program supports:
- `.xls`, `.xlsx`, `.html`, `.htm` files
- Files must contain jobcard tables with employee data, dates, and attendance codes

Expected raw file structure:
- Company/address/metadata at the top
- "Sub Section:" labels
- "SL" row indicating date headers
- Worker rows with attendance codes (P, A, W, etc.)
- Daily OT hours and totals rows

## Output Files

### With VBA (default):
- `output/Salary_Attendance_August_2026.xlsm` - Macro-enabled workbook
  - All attendance data and formulas
  - VBA macros included:
    - **Sub-section filtering**: Dropdown in A4 to filter by subsection
    - **Duty transfer**: Move worker duties between dates
    - **Print-time hiding**: Designation column hides on print
    - **Footer signatures**: "Prepared by", "Checked by", "Approved by"
  - User form for duty transfer (Alt+F8 or button)

### Without VBA (`--skip-macros`):
- `output/Salary_Attendance_August_2026.xlsx` - Regular Excel workbook
  - All data and formulas work
  - Sub-section filtering via AutoFilter (A4 dropdown)
  - No VBA macros

## Using the Generated Workbook

### Sub-Section Filtering
- Row 4, Column A (A4) has a dropdown filter
- Select a subsection name to show only that group
- Select "All" to show everyone
- Multiple subsections can be selected (joined with "+" sign)
- The SubSection Filter column is visible on screen but outside the print area

### Duty Transfer (**.xlsm files only**)
Access via **Alt+F8**:
- `ShowDutyTransferForm` - opens the duty transfer form
- `UndoLastTransfer` - undoes the last duty transfer

The form lets you:
- Choose a subsection (or "All")
- Pick a source date (duty must be present - code P or worked weekend W)
- Pick a target date (must show A for the worker)
- Move qualifying workers' duties
- Undo the last batch transfer

### Printing
- Designation column (D) is hidden automatically before printing
- Column D is hidden on paper but visible on screen
- Footer signatures appear at the bottom of each printed page
- Table fits to pages horizontally, continues as needed vertically

### Company logo watermark

The company logo (`nevertouch/nevertouch.png`) is printed on every page at the
top-right of the header band. The company/address/period rows keep their
original compact heights; the logo spans them plus the empty right side of
the Sub Section row, so the info text spacing and the number of data rows per
page stay close to the original layout. The logo is embedded at a fixed
standard size (128×128 px) inside the print area, so it never overflows or
underflows and scales with the sheet's fit-to-width print scaling.

The logo is embedded into the workbook package itself — it is **not linked** to
`nevertouch.png`. Once a `.xlsx`/`.xlsm` is generated, the logo remains in the
file even if `nevertouch.png` is deleted or replaced later.

### Footer Signatures
- Three default signatures: "Prepared by", "Checked by", "Approved by"
- Names can be edited on the "Lists" sheet in Excel
- Changes refresh footers without touching code
- Each name prints as "Name: __________" (22 underscores) at 9pt

### Formulas
- Monthly summary formulas (P, A, OT, EOT, TOT, CL, SL, EL, H) use COUNTIF/SUM
- These recalculate when Excel recalculates
- The file also caches formula results so numbers show immediately

## First-Time Usage Steps

1. **Place raw files** in the `raw_files/` directory
2. **Run the program**:
   ```bash
   python convert_attendance.py
   ```
3. **Find the output** in the `output/` folder
4. **Open in Excel** - enable content if prompted (for .xlsm files)
5. **Use the features**:
   - Filter by subsection using A4 dropdown
   - (With .xlsm) Use duty transfer form via Alt+F8
   - Print with hidden Designation column

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "No raw files found" | Ensure files are in `raw_files/` directory |
| "Could not read date columns" | Check HTML files contain proper jobcard tables with "SL" row |
| "Macros were not added" | Read the printed **Reason** line - it names the exact fix: install `pywin32`, enable *"Trust access to the VBA project object model"* in Excel's Trust Center, or start/repair Excel. The `.xlsx` still works with AutoFilter. |
| VBA compilation errors | Enable "Trust access to the VBA project object model" in Excel |
| Duty transfer form won't open | Ensure the .xlsm file is saved, not just .xlsx |
| Footer signatures not appearing | Edit names on the "Lists" sheet and reopen the file |

## Project Conventions

- **0-indexed rows/columns** in Python code, **1-indexed** in Excel/A1 notation
- **SHEET_ORDER** determines worksheet creation order: LO_PO, FIXED, PRODUCTION
- **SUMMARY_KEYS** = ["P", "A", "OT", "EOT", "TOT", "CL", "SL", "EL", "H"]
- **DATA_START_ROW** = 5 (0-based, Excel row 6 = first worker data)
- **HEADER_ROWS** = 5 (rows 1-5: company, address, period, subsection, table header)
- Print settings: A4 landscape, fit-to-width, 80% zoom, hidden gridlines

## Development

### Adding New Features
- Changes to `constants.py` affect sheet names, headers, and formatting
- `writer.py` handles all Excel generation - modify `_write_sheet()` for new layouts
- `vba.py` contains all VBA code - modify `attach_filter_macros()` for new macros
- `parser.py` parses raw HTML files - modify `_parse_section_table()` for new formats

### Testing
- Run `python convert_attendance.py` to generate output
- Open the `.xlsm` file in Excel to verify macros
- Test filtering, duty transfer, and printing
- Use `--skip-macros` to test without Excel

## License

This project is for internal Spring Trade attendance reporting.