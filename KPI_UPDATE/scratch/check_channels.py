import openpyxl

wb = openpyxl.load_workbook('thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx', read_only=True, data_only=True)
ws = wb.active
row_iter = ws.iter_rows(values_only=True)
hdr = next(row_iter)
bang_gia_vals = set()
kenh_ban_vals = set()

for i, r in enumerate(row_iter):
    if r[6]: bang_gia_vals.add(str(r[6]).strip())
    if r[8]: kenh_ban_vals.add(str(r[8]).strip())
    if i > 10000:
        break

print("Bảng giá values:", bang_gia_vals)
print("Kênh bán values:", kenh_ban_vals)
