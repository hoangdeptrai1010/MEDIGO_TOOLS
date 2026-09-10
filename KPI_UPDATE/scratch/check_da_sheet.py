import openpyxl

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_da = wb['Dự án']

print("=== Dự án sheet in bangluong_thang8_hoanthien.xlsx (sample rows 2-15) ===")
for r in range(2, 16):
    print([ws_da.cell(r, c).value for c in range(1, 13)])
