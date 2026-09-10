import sys, os, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

print("=== 1. LISTING ALL FILES IN thang7/Pharmacy_Retail_store_KPIs_July_2026 ===")
for root, dirs, files in os.walk('thang7/Pharmacy_Retail_store_KPIs_July_2026'):
    for f in files:
        p = os.path.join(root, f)
        print(f"  {p} ({os.path.getsize(p)} bytes)")

print("\n=== 2. INSPECTING thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx ===")
wb_bl7 = openpyxl.load_workbook('thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)
print("Sheets in BẢNG LƯƠNG target:", wb_bl7.sheetnames)

if 'Thưởng CK' in wb_bl7.sheetnames:
    ws = wb_bl7['Thưởng CK']
    print(f"\n--- Sheet 'Thưởng CK' (Formulas) ---")
    for r in range(1, min(20, ws.max_row + 1)):
        row_vals = [f"Col {openpyxl.utils.get_column_letter(c)}: {repr(ws.cell(r, c).value)}" for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None]
        if row_vals:
            print(f"Row {r:2d}: {', '.join(row_vals)}")

# Now data_only values
wb_bl7_val = openpyxl.load_workbook('thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)
if 'Thưởng CK' in wb_bl7_val.sheetnames:
    ws_v = wb_bl7_val['Thưởng CK']
    print(f"\n--- Sheet 'Thưởng CK' (Values) ---")
    for r in range(1, min(25, ws_v.max_row + 1)):
        row_vals = [f"{openpyxl.utils.get_column_letter(c)}={repr(ws_v.cell(r, c).value)}" for c in range(1, ws_v.max_column + 1) if ws_v.cell(r, c).value is not None]
        if row_vals:
            print(f"Row {r:2d}: {', '.join(row_vals)}")

print("\n=== 3. INSPECTING thang7/target/NHÀ THUỐC THÁNG 7 2026.xlsx ===")
wb_nt7 = openpyxl.load_workbook('thang7/target/NHÀ THUỐC THÁNG 7 2026.xlsx', data_only=False)
print("Sheets in NHÀ THUỐC target:", wb_nt7.sheetnames)
wb_nt7_val = openpyxl.load_workbook('thang7/target/NHÀ THUỐC THÁNG 7 2026.xlsx', data_only=True)

for sname in wb_nt7.sheetnames:
    ws_f = wb_nt7[sname]
    ws_v = wb_nt7_val[sname]
    print(f"\n--- NHÀ THUỐC THÁNG 7 -> Sheet: {sname} (max_row={ws_f.max_row}, max_col={ws_f.max_column}) ---")
    for r in range(1, min(10, ws_f.max_row + 1)):
        row_vals = [f"{openpyxl.utils.get_column_letter(c)}={repr(ws_v.cell(r, c).value)}" for c in range(1, min(20, ws_f.max_column + 1)) if ws_v.cell(r, c).value is not None]
        if row_vals:
            print(f"Row {r:2d}: {', '.join(row_vals)}")

