import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')
wb = openpyxl.load_workbook('thang8/tinhcongnhungthuongchia.xlsx', data_only=False)
ws = wb['BẢNG LƯƠNG']

for r in range(2, 10):
    c_j = ws.cell(r, 10)
    print(f"Row {r}: value={c_j.value}, number_format={c_j.number_format}")
