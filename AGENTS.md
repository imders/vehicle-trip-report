# AGENTS.md — Vehicle Trip Report Automation

## Project Context
This is a Python data-processing script that automates analysis of
vehicle surveillance camera logs. It reads an Excel file, detects
vehicle "trips" by time-gap analysis, and produces a formatted report.

## Tech Stack
- Language: Python 3.10+
- Core libraries: pandas >= 2.0, openpyxl
- No Windows-specific dependencies (no xlwings, win32com, xlrd)
- Input: .xlsx file with sheets «Выгрузка» and «База для расчета»
- Output: formatted .xlsx report file

## Code Standards
- Type hints required on ALL function signatures
- Google-style docstrings on all public functions
- NO .iterrows() — use only vectorized pandas operations
- NO stubs, NO TODO, NO pass — all code must be production-ready
- All string constants (column names, sheet names) as module-level variables

## Error Handling Rules
- Never crash silently — always log with logging module (INFO level)
- Unparseable dates → NaT → skip row with WARNING log
- Missing Excel sheet → raise KeyError with descriptive message
- FileNotFoundError → catch at main() level, print user-friendly message

## Naming Conventions
- Functions: snake_case, verbs (load_data, assign_trips, export_report)
- Constants: UPPER_SNAKE_CASE
- DataFrames: df_ prefix (df_raw, df_trips, df_base)

## Architecture Rules
- All logic split into pure functions (no God-function main)
- main() is orchestrator only — no business logic inside
- Script must run via: python script.py input.xlsx