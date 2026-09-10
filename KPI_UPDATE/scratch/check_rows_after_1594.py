import openpyxl

wb = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_nc = wb['Ngày công']

print("=== Checking rows 1590 to 1620 in Ngày công ===")
for r in range(1590, 1625):
    vals = [ws_nc.cell(r, c).value for c in range(1, 9)]
    if any(v is not None for v in vals):
        print(f"R{r}: {vals[:5]}")
    else:
        print(f"R{r}: EMPTY")
