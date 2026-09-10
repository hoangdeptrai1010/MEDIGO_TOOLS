import openpyxl

wb_cc = openpyxl.load_workbook('thang8/DATA/BangChiTietChamCong_thang8.xlsx', data_only=True)
wb_bl_aug = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=False)

ws_gc = wb_bl_aug['Giờ công']
ws_nc = wb_bl_aug['Ngày công']

print(f"wb_bl_aug Giờ công max_row: {ws_gc.max_row}")
print(f"wb_bl_aug Ngày công max_row: {ws_nc.max_row}")

# Let's count how many non-empty rows in Giờ công cols A-D
gc_data_rows = 0
for r in range(2, ws_gc.max_row + 1):
    if ws_gc.cell(r, 1).value is not None or ws_gc.cell(r, 2).value is not None:
        gc_data_rows += 1
print(f"Giờ công non-empty data rows: {gc_data_rows}")

# Let's check Ngày công non-empty rows in cols A-E
nc_data_rows = 0
for r in range(2, ws_nc.max_row + 1):
    if ws_nc.cell(r, 1).value is not None or ws_nc.cell(r, 2).value is not None:
        nc_data_rows += 1
print(f"Ngày công non-empty data rows: {nc_data_rows}")

# Let's inspect the source of Giờ công rows:
print("\nFirst 10 rows of Giờ công:")
for r in range(2, 12):
    print(f"R{r}: A={repr(ws_gc.cell(r,1).value)}, B={repr(ws_gc.cell(r,2).value)}, C={repr(ws_gc.cell(r,3).value)}, D={repr(ws_gc.cell(r,4).value)}, E={repr(ws_gc.cell(r,5).value)}")

# Let's inspect the source of Ngày công rows:
print("\nFirst 10 rows of Ngày công:")
for r in range(2, 12):
    print(f"R{r}: A={repr(ws_nc.cell(r,1).value)}, B={repr(ws_nc.cell(r,2).value)}, C={repr(ws_nc.cell(r,3).value)}, D={repr(ws_nc.cell(r,4).value)}, E={repr(ws_nc.cell(r,5).value)}")
