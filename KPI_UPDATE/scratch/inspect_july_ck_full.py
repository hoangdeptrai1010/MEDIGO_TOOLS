import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)
ws_ck = wb['Thưởng CK']

print("=== TẤT CẢ GIÁ TRỊ CỘT E (Thưởng CK - Nghĩa điền) TRONG THÁNG 7 TARGET ===")
for r in range(2, ws_ck.max_row + 1):
    cn = ws_ck.cell(r, 1).value
    name = ws_ck.cell(r, 2).value
    c_tot = ws_ck.cell(r, 3).value
    d_hs = ws_ck.cell(r, 4).value
    e_val = ws_ck.cell(r, 5).value
    
    h_name = ws_ck.cell(r, 8).value
    i_tot = ws_ck.cell(r, 9).value
    j_hs = ws_ck.cell(r, 10).value
    k_da = ws_ck.cell(r, 11).value
    l_hshh = ws_ck.cell(r, 12).value
    if name and (e_val or k_da):
        print(f"Row {r:2d}: [{cn}] {name:<24} | Thưởng CK gốc (E) = {str(e_val):>10} | HS (D) = {str(d_hs):>4} | Thực nhận (C) = {str(c_tot):>10} || Dự án (K) = {str(k_da):>10} | HS HH (L) = {str(l_hshh):>4}")
