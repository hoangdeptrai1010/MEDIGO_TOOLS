import pandas as pd
import openpyxl, sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE')

from TOOL_BANGLUONG.payroll_candate import normalize_branch

# 1. Quét tất cả (Chi nhánh, Tên nhân viên) có phát sinh giờ công trong BangChiTietChamCong_thang8.xlsx
cc_file = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/BangChiTietChamCong_thang8.xlsx'
df_cc = pd.read_excel(cc_file)

timecard_staff_branches = {} # (branch, name) -> {'shifts': count, 'hours': total_hours, 'details': []}

for idx, r in df_cc.iterrows():
    name = str(r.iloc[2] or '').strip()
    branch_raw = str(r.iloc[5] or '').strip()
    shift_name = str(r.iloc[6] or '').strip()
    shifts_val = r.iloc[39]
    hours_str = str(r.iloc[41] or '').strip()
    
    if name and branch_raw and name.lower() not in ('nhân viên', 'tổng', 'none', 'nan') and branch_raw.lower() not in ('chi nhánh', 'none', 'nan'):
        branch = normalize_branch(branch_raw)
        
        try:
            s_count = int(shifts_val) if pd.notna(shifts_val) and str(shifts_val) != '-' else 0
        except Exception:
            s_count = 0
            
        if s_count > 0 or (hours_str and hours_str != '-' and 'h' in hours_str):
            key = (branch, name)
            if key not in timecard_staff_branches:
                timecard_staff_branches[key] = {'shifts': 0, 'details': []}
            timecard_staff_branches[key]['shifts'] += s_count
            timecard_staff_branches[key]['details'].append(f"{shift_name} ({hours_str}, {s_count} ca)")

print(f"=== TỔNG CỘNG CÓ {len(timecard_staff_branches)} CẶP (CHI NHÁNH, NHÂN VIÊN) TRONG BẢNG CHẤM CÔNG ===")

# 2. Quét các dòng trong BẢNG LƯƠNG của template / file kết quả
bl_file = 'd:/MEDIGO/KPI_UPDATE/thang8/BANGLUONGTHANG8.xlsx'
wb_bl = openpyxl.load_workbook(bl_file, data_only=True)
ws_bl = wb_bl['BẢNG LƯƠNG']

payroll_keys = set()
for r in range(3, ws_bl.max_row + 1):
    b_val = ws_bl.cell(r, 2).value
    n_val = ws_bl.cell(r, 3).value
    if b_val and n_val:
        b_clean = normalize_branch(b_val)
        n_clean = str(n_val).strip()
        payroll_keys.add((b_clean, n_clean))

print(f"=== TỔNG CỘNG CÓ {len(payroll_keys)} CẶP (CHI NHÁNH, NHÂN VIÊN) TRONG BẢNG LƯƠNG ===\n")

# 3. So sánh tìm các trường hợp BỊ THIẾU DÒNG TRONG BẢNG LƯƠNG
missing_in_payroll = []
for k, data in sorted(timecard_staff_branches.items()):
    if k not in payroll_keys:
        missing_in_payroll.append((k, data))

print(f"🚨 PHÁT HIỆN {len(missing_in_payroll)} TRƯỜNG HỢP CÓ CHẤM CÔNG NHƯNG CHƯA CÓ DÒNG TRONG BẢNG LƯƠNG:")
for (br, nm), data in missing_in_payroll:
    print(f"  ❌ [{br}] {nm}: Tổng {data['shifts']} ca làm việc -> Chi tiết: {data['details']}")

# 4. Kiểm tra sheet Giờ công và Ngày công
print("\n" + "=" * 70)
print("🔍 KIỂM TRA TRONG SHEET 'Giờ công' VÀ 'Ngày công'")
print("=" * 70)

for s_name in ['Giờ công', 'Ngày công']:
    ws = wb_bl[s_name]
    s_keys = set()
    for r in range(2, ws.max_row + 1):
        # Sheet Giờ công: Col A=CN, Col B=Tên
        b = ws.cell(r, 1).value
        n = ws.cell(r, 2).value
        if b and n:
            s_keys.add((normalize_branch(b), str(n).strip()))
    print(f"Sheet [{s_name}]: {len(s_keys)} cặp (Chi nhánh, Nhân viên)")
    
    missing_in_sub = [(k, data) for k, data in timecard_staff_branches.items() if k not in s_keys]
    if missing_in_sub:
        print(f"  🚨 Thiếu {len(missing_in_sub)} cặp trong [{s_name}]:")
        for (br, nm), data in missing_in_sub:
            print(f"     - [{br}] {nm}")
    else:
        print(f"  ✅ Đầy đủ 100% trong [{s_name}]")
