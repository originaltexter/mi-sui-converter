# Michigan MiUI SUI Wage Report Converter

Converts a **QuickBooks Desktop** SUI 1028 export (`.xlsm`) into the
delimited CSV format required by Michigan's
[MiUI unemployment insurance system](https://miui.michigan.gov)
for quarterly wage reporting.

## Features

- GUI — no command line needed
- Auto-fills employer info (name, address, EAN) directly from the QB export
- Detects the reporting quarter automatically
- Handles single-location employers (unit `000`)
- Excludes employees with no wages in the quarter
- Derives 12th-of-month employment indicators from pay period data

## Requirements

- Windows
- Python 3.8+ with `pandas` and `openpyxl`  
  *(not needed if using the pre-built `.exe`)*

## Usage

### Option A — Use the pre-built .EXE
Download the EXE from the Releases page

### Option B — Run the Python script directly
pip install pandas openpyxl
python sui_converter.py

### Option C — Build a standalone .exe
1. Install Python 3.8+ from [python.org](https://python.org) (check *Add to PATH*)
2. Place `sui_converter.py` and `build_exe.bat` in the same folder
3. Double-click `build_exe.bat`
4. Find `SUI_Converter.exe` in the `dist\` subfolder — distribute this file alone

## Quarterly due dates

| Quarter | Months | Due |
|---------|--------|-----|
| Q1 | Jan–Mar | April 25 |
| Q2 | Apr–Jun | July 25 |
| Q3 | Jul–Sep | October 25 |
| Q4 | Oct–Dec | January 25 |

## Notes

- Wages are sourced from the **State Report** sheet (QB's SUI-subject wages for the quarter)
- Employees with $0 wages for the quarter are excluded automatically
- Unit number is hardcoded to `000` (single-location employer)
- Obligation Gross Wages and Visa Wage Indicator use inactive defaults per current MiUI spec

