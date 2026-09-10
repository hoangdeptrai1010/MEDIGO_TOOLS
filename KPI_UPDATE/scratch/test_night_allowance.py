import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_nc = wb['Ngày công']

# Count valid night shifts (>= 4.0h, name contains "đêm")
night_shifts = {}
for r in range(3, ws_nc.max_row+1):
    name = ws_nc.cell(r, 2).value
    ca = ws_nc.cell(r, 4).value
    gio = ws_nc.cell(r, 6).value
    if not name or not ca or not isinstance(gio, (int, float)):
        continue
    n_str = str(name).strip()
    is_night = 'đêm' in str(ca).lower()
    if is_night and gio >= 4.0:
        night_shifts[n_str] = night_shifts.get(n_str, 0) + 1

ws_bl = wb['BẢNG LƯƠNG']
print(f"{'Row':<4} | {'Tên nhân viên':<24} | {'Chức vụ':<10} | {'Số ca đêm >=4h (K)':<18} | {'Phụ cấp đêm mới (X)':<20}")
print("-" * 85)

tot_pc = 0.0
for r in range(3, ws_bl.max_row+1):
    name = ws_bl.cell(r, 3).value
    role = str(ws_bl.cell(r, 4).value or '').strip()
    if name:
        n_str = str(name).strip()
        k = night_shifts.get(n_str, 0)
        is_dscd = 'DSCD' in role.upper()
        
        if is_dscd:
            pc = round((1500000.0 / 28.0) * k, 0)
        else:
            if k > 20:
                pc = round((1500000.0 / 28.0) * k, 0)
            else:
                pc = 0.0
                
        tot_pc += pc
        flag = "*** ĐẠT THƯỞNG ***" if pc > 0 else ""
        print(f"{r:<4} | {n_str:<24} | {role:<10} | {k:>18} | {pc:>18,.0f} đ {flag}")

print("-" * 85)
print(f"--> TỔNG CỘNG QUỸ PHỤ CẤP CA ĐÊM: {tot_pc:,.0f} đ")
