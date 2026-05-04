# Excel Spec Compare

A simple Python tool for comparing two Excel spec files and creating a colored Excel report.

## What It Does

It compares:

- Added sheets, columns, and rows
- Removed sheets, columns, and rows
- Changed values

The Excel report is color coded:

- Green: added
- Red: removed
- Yellow: changed
- Orange: warning

For structure/data type specs, it compares rows by:

```text
Structure Name + Field Name
```

This means row order changes will not create noisy false differences.

## Project Structure

```text
ExcelSpecComparePython/
  compare_excel_specs.py
  requirements.txt
  README.md
  examples/
    create_struct_spec_sample.py
    README.md
```

The main code is only:

```text
compare_excel_specs.py
```

## Setup on a New Machine

Clone the repository, then run:

```powershell
cd ExcelSpecComparePython
python -m pip install -r requirements.txt
```

If Windows does not recognize `python`, reinstall Python and enable:

```text
Add python.exe to PATH
```

## Run with Real Files

```powershell
python compare_excel_specs.py "C:\path\old_spec.xlsx" "C:\path\new_spec.xlsx" -o "C:\path\diff_report.xlsx"
```

For structure specs, this is the recommended command:

```powershell
python compare_excel_specs.py "C:\path\old_spec.xlsx" "C:\path\new_spec.xlsx" --profile struct-spec -o "C:\path\diff_report.xlsx"
```

If your key columns have custom names:

```powershell
python compare_excel_specs.py "old_spec.xlsx" "new_spec.xlsx" --key-column "Structure Name" --key-column "Field Name" -o "diff_report.xlsx"
```

## Adapting to Your Own Spec Format

Use this table when you receive a new spec format and need to decide what to change.

| Situation | What to do | Example |
| --- | --- | --- |
| Your key columns are different | Do not edit code. Pass `--key-column` for each key column. | `--key-column "Struct" --key-column "Member"` |
| Your header row is not row 1 | Pass `--header-row`. | `--header-row 3` |
| You want to compare by exact cell position | Use cell mode. | `--mode cell` |
| You want the tool to auto-detect a new structure column name | Add the name to `STRUCTURE_ALIASES` in `compare_excel_specs.py`. | Add `"class name"` |
| You want the tool to auto-detect a new field column name | Add the name to `FIELD_ALIASES` in `compare_excel_specs.py`. | Add `"property name"` |
| You want to change report columns | Edit `REPORT_COLUMNS` in `compare_excel_specs.py`. | Add or remove output column names |
| You want to change report colors | Edit `ADDED_FILL`, `REMOVED_FILL`, `CHANGED_FILL`, or `WARNING_FILL`. | Change Excel color hex code |
| You want to ignore some columns | Not supported yet. Add an ignore-column option or remove those columns before comparing. | `Updated Date`, `Author`, `Comment` |
| Your file has duplicate keys | Check orange warning rows in the report. Duplicate keys should usually be fixed in the spec. | Same `Structure Name + Field Name` appears twice |

## Run the Example

```powershell
cd ExcelSpecComparePython
python examples\create_struct_spec_sample.py
python compare_excel_specs.py examples\struct_spec_old.xlsx examples\struct_spec_new.xlsx --profile struct-spec -o examples\struct_diff_report.xlsx
```

Then open:

```text
examples\struct_diff_report.xlsx
```

## Options

```text
--profile struct-spec        Compare by Structure Name + Field Name
--mode auto|keyed|cell       Default: auto
--key-column "Column Name"   Use custom key column; can be repeated
--header-row 1               Header row number
--ignore-case                Ignore case differences
--no-trim                    Do not trim whitespace
-o report.xlsx               Output report file
```
