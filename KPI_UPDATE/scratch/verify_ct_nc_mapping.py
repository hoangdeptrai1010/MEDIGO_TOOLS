import openpyxl

wb_cc = openpyxl.load_workbook('thang8/DATA/BangChiTietChamCong_thang8.xlsx', data_only=True)
ws_ct = wb_cc['Bảng chi tiết chấm công']
wb_baug = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_nc = wb_baug['Ngày công']

print("Comparing ws_ct row 4..10 with ws_nc row 2..8:")
for i in range(10):
    r_ct = 4 + i
    r_nc = 2 + i
    ct_b = ws_ct.cell(r_ct, 2).value # Ten NV
    ct_e = ws_ct.cell(r_ct, 5).value # Chi nhanh
    ct_f = ws_ct.cell(r_ct, 6).value # Ngay
    ct_h = ws_ct.cell(r_ct, 8).value # Ca
    ct_u = ws_ct.cell(r_ct, 21).value # Gio thuc te
    
    nc_a = ws_nc.cell(r_nc, 1).value # Chi nhanh
    nc_b = ws_nc.cell(r_nc, 2).value # Ten NV
    nc_c = ws_nc.cell(r_nc, 3).value # Ngay
    nc_d = ws_nc.cell(r_nc, 4).value # Ca
    nc_e = ws_nc.cell(r_nc, 5).value # Gio thuc te
    print(f"CT R{r_ct}: {ct_e} | {ct_b} | {ct_f} | {ct_h} | {ct_u}")
    print(f"NC R{r_nc}: {nc_a} | {nc_b} | {nc_c} | {nc_d} | {nc_e}")
    print("-" * 80)
