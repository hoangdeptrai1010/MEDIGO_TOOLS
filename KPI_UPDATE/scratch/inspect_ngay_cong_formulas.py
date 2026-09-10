import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/tinhcongnhungthuongchia.xlsx', data_only=False)
ws = wb['Ngày công']

print("=== CÁC CÔNG THỨC TRONG SHEET NGÀY CÔNG (Header & Rows 2-10) ===")
headers = [ws.cell(1, c).value for c in range(1, 15)]
print("Headers:", headers)

for r in range(2, 12):
    row_vals = [ws.cell(r, c).value for c in range(1, 15)]
    print(f"Row {r:2d}: {row_vals}")

# Check for Hứa Thị Kim Thoa in Ngày công
print("\n=== Chi tiết Hứa Thị Kim Thoa trong Ngày công ===")
for r in range(2, ws.max_row+1):
    name = str(ws.cell(r, 2).value or '')
    if 'thoa' in name.lower():
        row_vals = [ws.cell(r, c).value for c in range(1, 11)]
        print(f"Row {r:3d}: {row_vals}")
