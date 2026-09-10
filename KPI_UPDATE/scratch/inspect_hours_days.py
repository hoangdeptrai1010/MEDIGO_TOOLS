import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_bl = wb['BẢNG LƯƠNG']

print(f"{'Row':<4} | {'Tên':<24} | {'Giờ ngày':<10} | {'Giờ đêm':<10} | {'Ngày (J)':<8} | {'Đêm (K)':<8} | {'Công (R)':<8}")
print("-" * 85)
for r in range(3, ws_bl.max_row+1):
    name = ws_bl.cell(r, 3).value
    if name:
        gn = ws_bl.cell(r, 7).value or 0
        gd = ws_bl.cell(r, 8).value or 0
        nj = ws_bl.cell(r, 10).value or 0
        nk = ws_bl.cell(r, 11).value or 0
        nr = ws_bl.cell(r, 18).value or 0
        if isinstance(gn, (int, float)) and isinstance(gd, (int, float)):
            print(f"{r:<4} | {name:<24} | {gn:>10.2f} | {gd:>10.2f} | {str(nj):>8} | {str(nk):>8} | {str(nr):>8}")
