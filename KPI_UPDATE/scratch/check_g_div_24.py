# -*- coding: utf-8 -*-
import openpyxl, sys
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws = wb['BẢNG LƯƠNG']

print(f"{'STT':<3} | {'Chi nhánh':<15} | {'Họ và tên':<24} | {'Chức danh':<8} | {'Giờ ngày (G)':<12} | {'G/24 (exact)':<12} | {'ROUND(G/24,0)':<14} | {'INT(G/24)':<10} | {'Ca đêm (K)':<10}")
print("-" * 125)

for r in range(3, 59):
    stt = ws.cell(r, 1).value
    cn = ws.cell(r, 2).value or ''
    name = ws.cell(r, 3).value or ''
    role = ws.cell(r, 4).value or ''
    g = ws.cell(r, 7).value or 0
    k = ws.cell(r, 11).value or 0
    
    g_num = float(g) if isinstance(g, (int, float)) else 0.0
    k_num = int(k) if isinstance(k, (int, float)) else 0
    
    exact = g_num / 24.0
    r_val = round(exact)
    int_val = int(exact)
    
    print(f"{stt:<3} | {cn:<15} | {name:<24} | {role:<8} | {g_num:<12.2f} | {exact:<12.2f} | {r_val:<14} | {int_val:<10} | {k_num:<10}")
