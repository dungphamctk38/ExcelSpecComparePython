# Example

Create sample Excel files:

```powershell
cd ExcelSpecComparePython
python examples\create_struct_spec_sample.py
```

Compare them:

```powershell
python compare_excel_specs.py examples\struct_spec_old.xlsx examples\struct_spec_new.xlsx --profile struct-spec -o examples\struct_diff_report.xlsx
```

Open:

```text
examples\struct_diff_report.xlsx
```

