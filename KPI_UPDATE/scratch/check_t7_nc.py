import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb7 = openpyxl.load_workbook('thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)
ws7_nc = wb7['Ngày công']
print("--- Month 7 Template 'Ngày công' rows 1 to 5 ---")
for r in range(1, 6):
    print(r, [ws7_nc.cell(r, c).value for c in range(1, 10)])
