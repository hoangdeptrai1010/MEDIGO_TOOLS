# -*- coding: utf-8 -*-
import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING ALL SHEETS IN baocaokpi_thang8_hoanthien.xlsx ===")
wb_kpi = openpyxl.load_workbook('baocaokpi_thang8_hoanthien.xlsx', data_only=False)
for s in wb_kpi.sheetnames:
    ws = wb_kpi[s]
    # check first 5 rows
    print(f"\nSheet: {s} (max_row={ws.max_row}, max_col={ws.max_column})")
    for r in range(1, min(5, ws.max_row + 1)):
        row_vals = [f"Col{c}:{ws.cell(r, c).value}" for c in range(1, min(15, ws.max_column + 1)) if ws.cell(r, c).value is not None]
        if row_vals:
            print(f"  Row {r}: {row_vals[:8]}")

print("\n=== CHECKING ALL SHEETS IN thang8/bangluong_thang8_hoanthien.xlsx ===")
wb_bl = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=False)
for s in wb_bl.sheetnames:
    ws = wb_bl[s]
    print(f"\nSheet: {s} (max_row={ws.max_row}, max_col={ws.max_column})")
    for r in range(1, min(5, ws.max_row + 1)):
        row_vals = [f"Col{c}:{ws.cell(r, c).value}" for c in range(1, min(15, ws.max_column + 1)) if ws.cell(r, c).value is not None]
        if row_vals:
            print(f"  Row {r}: {row_vals[:8]}")
