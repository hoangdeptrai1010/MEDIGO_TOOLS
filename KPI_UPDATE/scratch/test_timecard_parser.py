import openpyxl, os, sys

sys.stdout.reconfigure(encoding='utf-8')

def parse_timecard_data(timecard_path):
    print(f"--> Parsing raw timecards from: {timecard_path}")
    wb = openpyxl.load_workbook(timecard_path, data_only=True)
    ws = wb.active
    
    # Store shifts per (branch, staff_name)
    data = {}
    curr_staff = ""
    curr_branch = ""
    curr_role = ""
    curr_code = ""
    
    BRANCH_ALIAS = {
        'NT 24H Trường Sa': 'Trường Sa',
        'NT 24h Đỗ Quang Đẩu': 'Đỗ Quang Đẩu',
        'NT 24H Nam Hòa': 'Nam Hòa',
        'NT 24H Minh Châu 1': 'Minh Châu',
        'NT 24H Lê Bình': 'Lê Bình',
        'NT 24H Nguyễn Chí Thanh': 'Nguyễn Chí Thanh',
        'NT 24H Nguyễn Thị Thập': 'Nguyễn Thị Thập',
        'NT 24H Nguyễn Văn Quá': 'Nguyễn Văn Quá',
        'NT 24H Rạch Bùng Binh': 'Rạch Bùng Binh',
        'NT Đường Láng 247': 'Đường Láng',
        'NT Hàng Bông 247': 'Hàng Bông',
    }
    
    for r in range(5, ws.max_row + 1):
        c_name = ws.cell(r, 3).value
        c_branch = ws.cell(r, 6).value
        c_shift = ws.cell(r, 7).value
        c_role = ws.cell(r, 5).value
        c_code = ws.cell(r, 2).value
        
        if c_name and str(c_name).strip():
            curr_staff = str(c_name).strip()
        if c_branch and str(c_branch).strip():
            curr_branch = BRANCH_ALIAS.get(str(c_branch).strip(), str(c_branch).strip())
        if c_role and str(c_role).strip():
            curr_role = str(c_role).strip()
        if c_code and str(c_code).strip():
            curr_code = str(c_code).strip()
            
        if not curr_staff or not c_shift:
            continue
            
        shift_name = str(c_shift).strip()
        
        for day in range(1, 32):
            col_in = 8 + (day - 1) * 2
            col_out = col_in + 1
            v_in = str(ws.cell(r, col_in).value or '').strip()
            v_out = str(ws.cell(r, col_out).value or '').strip()
            
            if v_in and v_out:
                try:
                    # Clean HH:MM
                    p_in = v_in.split(' ')[0].split(':')
                    p_out = v_out.split(' ')[0].split(':')
                    h_in, m_in = int(p_in[0]), int(p_in[1])
                    h_out, m_out = int(p_out[0]), int(p_out[1])
                    
                    t_in = h_in * 60 + m_in
                    t_out = h_out * 60 + m_out
                    
                    # Overnight detection
                    is_overnight = False
                    if 'đêm' in shift_name.lower() or 'qua đêm' in shift_name.lower() or t_out < t_in:
                        is_overnight = True
                        if t_out < t_in:
                            t_out += 24 * 60
                    elif t_out < t_in:
                        t_out += 24 * 60
                        
                    dur_m = t_out - t_in
                    dur_h = dur_m / 60.0
                    
                    # Classify day vs night
                    is_night_shift = ('đêm' in shift_name.lower() or 'qua đêm' in shift_name.lower() or h_in >= 22 or h_in < 6)
                    
                    key = (curr_branch, curr_staff)
                    if key not in data:
                        data[key] = {
                            'branch': curr_branch,
                            'name': curr_staff,
                            'role': curr_role,
                            'code': curr_code,
                            'shifts': [],
                            'day_hrs': 0.0,
                            'night_hrs': 0.0,
                            'tot_hrs': 0.0,
                            'day_shifts_gt4': 0,
                            'night_shifts_gt4': 0,
                            'days_worked': set()
                        }
                    
                    d_obj = data[key]
                    d_obj['shifts'].append({
                        'day': day,
                        'shift': shift_name,
                        'in': v_in,
                        'out': v_out,
                        'hrs': dur_h,
                        'is_night': is_night_shift,
                        'gt4': (dur_h > 4.0)
                    })
                    
                    if is_night_shift:
                        d_obj['night_hrs'] += dur_h
                        if dur_h > 4.0:
                            d_obj['night_shifts_gt4'] += 1
                    else:
                        d_obj['day_hrs'] += dur_h
                        if dur_h > 4.0:
                            d_obj['day_shifts_gt4'] += 1
                            
                    d_obj['tot_hrs'] += dur_h
                    d_obj['days_worked'].add(day)
                except Exception as e:
                    pass
    wb.close()
    return data

res = parse_timecard_data('thang8/DATA/BangChiTietChamCong_thang8.xlsx')
print(f"Total branch-staff entries found: {len(res)}")

# Print summary for 10 entries including Thoa and Tam
for k, v in list(res.items())[:15]:
    print(f"\n{v['branch']} - {v['name']} ({v['role']}):")
    print(f"  - Giờ ca ngày: {v['day_hrs']:.2f}h ({v['day_shifts_gt4']} ca >4h)")
    print(f"  - Giờ ca đêm: {v['night_hrs']:.2f}h ({v['night_shifts_gt4']} ca >4h)")
    print(f"  - Tổng ngày công thực tế: {len(v['days_worked'])} ngày | Tổng giờ: {v['tot_hrs']:.2f}h")
