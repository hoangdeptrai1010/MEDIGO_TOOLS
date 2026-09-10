import openpyxl

wb_cc = openpyxl.load_workbook('thang8/DATA/BangChiTietChamCong_thang8.xlsx', data_only=True)
ws_th = wb_cc['Bảng tổng hợp chấm công']
ws_ct = wb_cc['Bảng chi tiết chấm công']

print(f"Bảng tổng hợp chấm công rows: {ws_th.max_row}, cols: {ws_th.max_column}")
print(f"Bảng chi tiết chấm công rows: {ws_ct.max_row}, cols: {ws_ct.max_column}")

# Check columns of Bảng tổng hợp chấm công:
print("\nBảng tổng hợp chấm công headers (Row 2):")
for c in range(1, ws_th.max_column + 1):
    v = ws_th.cell(2, c).value
    if v:
        print(f"Col {c} ({openpyxl.utils.get_column_letter(c)}): {repr(v)}")

# Look at the end columns of Bảng tổng hợp chấm công (after day 31):
print("\nBảng tổng hợp chấm công summary columns (Row 2, cols >= 35):")
for c in range(35, ws_th.max_column + 1):
    v2 = ws_th.cell(2, c).value
    v3 = ws_th.cell(3, c).value
    print(f"Col {c} ({openpyxl.utils.get_column_letter(c)}): r2={repr(v2)}, r3={repr(v3)}")

print("\nSample rows of Bảng tổng hợp chấm công (Rows 4-8, cols 1-7 and last 10 cols):")
for r in range(4, 9):
    left = [ws_th.cell(r, c).value for c in range(1, 8)]
    right = [ws_th.cell(r, c).value for c in range(ws_th.max_column - 8, ws_th.max_column + 1)]
    print(f"R{r}: Left={left} | Right={right}")
