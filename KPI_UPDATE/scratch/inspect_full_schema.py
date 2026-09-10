import openpyxl
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

f_nt7 = r"d:\MEDIGO\KPI_UPDATE\goc\NHÀ THUỐC THÁNG 7 2026.xlsx"
f_nt8 = r"d:\MEDIGO\KPI_UPDATE\goc\NHÀ THUỐC THÁNG 8 2026.xlsx"

for tag, fpath in [("T7", f_nt7), ("T8", f_nt8)]:
    print(f"\n==========================================================================")
    print(f"                         FULL SCHEMA OF {tag}")
    print(f"==========================================================================")
    wb = openpyxl.load_workbook(fpath, data_only=True)
    for sname in wb.sheetnames:
        ws = wb[sname]
        print(f"\n==================== SHEET: '{sname}' ({ws.max_row} rows, {ws.max_column} cols) ====================")
        # Find header rows (rows 1-3)
        for r in range(1, min(5, ws.max_row + 1)):
            row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            non_empty = [(c, v) for c, v in enumerate(row_vals, 1) if v is not None]
            if non_empty:
                print(f"Row {r:2d}:")
                for c, v in non_empty:
                    print(f"  [{c:2d} | {openpyxl.utils.get_column_letter(c):>3s}]: {repr(v)}")
        # Print first 2 data rows
        print("Sample data row:")
        for r in range(min(5, ws.max_row + 1), min(7, ws.max_row + 1)):
            vals = [ws.cell(r, c).value for c in range(1, min(25, ws.max_column + 1))]
            print(f"  Data Row {r:2d}: {vals}")
