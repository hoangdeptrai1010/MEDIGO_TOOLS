import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb_val = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
wb_form = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=False)

ws_v = wb_val['Thưởng CK']
ws_f = wb_form['Thưởng CK']

print("=== SHEET THƯỞNG CK (GIÁ TRỊ VÀ CÔNG THỨC MỚI) ===")
print(f"Header: {[ws_v.cell(1, c).value for c in range(1, 13)]}")
for r in range(2, 15):
    h_name = ws_v.cell(r, 8).value
    i_val = ws_v.cell(r, 9).value
    j_val = ws_v.cell(r, 10).value
    k_val = ws_v.cell(r, 11).value
    l_val = ws_v.cell(r, 12).value
    
    b_name = ws_v.cell(r, 2).value
    c_val = ws_v.cell(r, 3).value
    d_val = ws_v.cell(r, 4).value
    e_val = ws_v.cell(r, 5).value
    
    print(f"Row {r:2d}: Trái (B={b_name:<22} | C={c_val:>6} | D={d_val:>4} | E={e_val:>4}) || Phải (H={h_name:<22} | I={i_val:>6} | J={j_val:>4} | K={k_val:>8} | L={l_val:>4})")

