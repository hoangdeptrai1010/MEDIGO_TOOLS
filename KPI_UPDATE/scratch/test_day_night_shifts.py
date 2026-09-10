import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_nc = wb['Ngày công']

staff_shifts = {}
for r in range(3, ws_nc.max_row+1):
    name = ws_nc.cell(r, 2).value
    date_v = ws_nc.cell(r, 3).value
    ca = ws_nc.cell(r, 4).value
    gio = ws_nc.cell(r, 6).value
    if not name or not date_v or not isinstance(gio, (int, float)) or gio <= 0:
        continue
    n_str = str(name).strip()
    d_str = str(date_v).split()[0]
    if n_str not in staff_shifts:
        staff_shifts[n_str] = {'day_dates': set(), 'night_dates': set(), 'all_dates': set()}
    staff_shifts[n_str]['all_dates'].add(d_str)
    if 'đêm' in str(ca).lower():
        staff_shifts[n_str]['night_dates'].add(d_str)
    else:
        staff_shifts[n_str]['day_dates'].add(d_str)

ws_bl = wb['BẢNG LƯƠNG']
print(f"{'Row':<4} | {'Tên nhân viên':<24} | {'Giờ ngày':<10} | {'Giờ đêm':<10} | {'Ngày ca ngày':<12} | {'Ngày ca đêm':<12} | {'Tổng ngày công (R)':<18}")
print("-" * 105)
for r in range(3, ws_bl.max_row+1):
    name = ws_bl.cell(r, 3).value
    if name and str(name).strip() in staff_shifts:
        n_str = str(name).strip()
        gn = ws_bl.cell(r, 7).value or 0
        gd = ws_bl.cell(r, 8).value or 0
        s = staff_shifts[n_str]
        print(f"{r:<4} | {n_str:<24} | {gn:>10.2f} | {gd:>10.2f} | {len(s['day_dates']):>12} | {len(s['night_dates']):>12} | {min(31, len(s['all_dates'])):>18}")
