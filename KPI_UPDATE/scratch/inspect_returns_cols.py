import openpyxl

wb = openpyxl.load_workbook('thang8/DATA/DanhSachChiTietTraHang_3182026.xlsx', read_only=True, data_only=True)
ws = wb.active
for i, r in enumerate(ws.iter_rows(values_only=True)):
    if i < 5:
        print(f"r{i+1}: {r}")
    else:
        break
