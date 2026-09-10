import openpyxl

wb_orig = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_gc_orig = wb_orig['Giờ công']
ws_nc_orig = wb_orig['Ngày công']
ws_bl_orig = wb_orig['BẢNG LƯƠNG']

print("=== thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx ===")
print("Giờ công max_row:", ws_gc_orig.max_row)
print("Ngày công max_row:", ws_nc_orig.max_row)
print("BẢNG LƯƠNG max_row:", ws_bl_orig.max_row)

print("\n--- Giờ công (Rows 2-10) ---")
for r in range(2, 11):
    print([ws_gc_orig.cell(r, c).value for c in range(1, 12)])

print("\n--- BẢNG LƯƠNG (Rows 3-10, Cols 1-18) ---")
for r in range(3, 11):
    vals = [ws_bl_orig.cell(r, c).value for c in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 18]]
    print(vals)
