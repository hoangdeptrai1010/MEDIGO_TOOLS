import openpyxl

wb_baug = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_bl = wb_baug['BẢNG LƯƠNG']

# Recalculate using openpyxl or check values
print("Values in BẢNG LƯƠNG of thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx:")
for r in [3, 4, 5, 10, 11, 12, 13, 55, 56, 63, 64, 68, 69]:
    cn = ws_bl.cell(r, 2).value
    name = ws_bl.cell(r, 3).value
    role = ws_bl.cell(r, 4).value
    gn = ws_bl.cell(r, 5).value
    gd = ws_bl.cell(r, 6).value
    g_sang = ws_bl.cell(r, 7).value
    g_dem = ws_bl.cell(r, 8).value
    ngay_cong = ws_bl.cell(r, 18).value
    print(f"R{r} ({cn} - {name} - {role}): G_sang={g_sang}, G_dem={g_dem}, Ngay_cong={ngay_cong}")
