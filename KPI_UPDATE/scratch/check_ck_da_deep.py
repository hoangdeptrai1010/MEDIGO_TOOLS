import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

print("=== 1. JULY BẢNG LƯƠNG THÁNG 7 2026.xlsx ===")
wb_j = openpyxl.load_workbook('thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)
ws_j_ck = wb_j['Thưởng CK']
print("--- July Thưởng CK non-zero rows ---")
for r in range(1, ws_j_ck.max_row + 1):
    vals = [ws_j_ck.cell(r, c).value for c in range(1, 15)]
    if any(v is not None and v != 0 for v in vals):
        print(f"Row {r:2d}: {vals}")

ws_j_da = wb_j['Dự án']
print("\n--- July Dự án non-zero rows ---")
for r in range(1, ws_j_da.max_row + 1):
    vals = [ws_j_da.cell(r, c).value for c in range(1, 15)]
    if any(v is not None and v != 0 for v in vals):
        print(f"Row {r:2d}: {vals}")

print("\n=== 2. AUGUST ORIGINAL BẢNG LƯƠNG THÁNG 8 2026.xlsx ===")
wb_a_orig = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_a_ck = wb_a_orig['Thưởng CK']
print("--- Aug Orig Thưởng CK non-zero rows ---")
for r in range(1, ws_a_ck.max_row + 1):
    vals = [ws_a_ck.cell(r, c).value for c in range(1, 15)]
    if any(v is not None and v != 0 for v in vals):
        print(f"Row {r:2d}: {vals}")

ws_a_da = wb_a_orig['Dự án']
print("\n--- Aug Orig Dự án non-zero rows ---")
for r in range(1, ws_a_da.max_row + 1):
    vals = [ws_a_da.cell(r, c).value for c in range(1, 15)]
    if any(v is not None and v != 0 for v in vals):
        print(f"Row {r:2d}: {vals}")

print("\n=== 3. AUGUST CURRENT bangluong_thang8_hoanthien.xlsx ===")
wb_a_curr = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_c_ck = wb_a_curr['Thưởng CK']
print("--- Aug Curr Thưởng CK non-zero rows ---")
for r in range(1, ws_c_ck.max_row + 1):
    vals = [ws_c_ck.cell(r, c).value for c in range(1, 15)]
    if any(v is not None and v != 0 for v in vals):
        print(f"Row {r:2d}: {vals}")

ws_c_da = wb_a_curr['Dự án']
print("\n--- Aug Curr Dự án non-zero rows ---")
for r in range(1, ws_c_da.max_row + 1):
    vals = [ws_c_da.cell(r, c).value for c in range(1, 15)]
    if any(v is not None and v != 0 for v in vals):
        print(f"Row {r:2d}: {vals}")
