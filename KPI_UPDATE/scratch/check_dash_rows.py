import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_gc = wb['Giờ công']
ws_nc = wb['Ngày công']

print("--- Inspect rows with '-' in Giờ công (first 10) ---")
count = 0
for r in range(2, 50):
    val_c = ws_gc.cell(r, 3).value
    if val_c == '-' or val_c == '- ':
        count += 1
        print(r, [ws_gc.cell(r, c).value for c in range(1, 6)])

print(f"\n--- Inspect rows with '-' in Ngày công (first 10) ---")
count_nc = 0
for r in range(2, 50):
    val_d = ws_nc.cell(r, 4).value
    if val_d == '-' or val_d == '- ':
        count_nc += 1
        print(r, [ws_nc.cell(r, c).value for c in range(1, 6)])
