import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_bl = wb['BẢNG LƯƠNG']

print("=== SAMPLE BẢNG LƯƠNG THÁNG 8 2026 ===")
headers = [ws_bl.cell(2, c).value for c in range(1, 45)]

cols_to_show = [
    (1, "STT"), (2, "Chi nhánh"), (3, "Họ tên"), (4, "Chức vụ"),
    (5, "Tổng giờ ngày"), (6, "Tổng giờ đêm"), (18, "Ngày công"),
    (27, "MiniKat DS"), (28, "MiniKat CHT"), (29, "WhatsApp"),
    (30, "Dự án"), (31, "KPI"), (32, "Thưởng CK"), (33, "Cận date"),
    (36, "Tổng trừ"), (41, "Trừ khác"), (42, "Ghi chú trừ")
]

sample_staff = [
    'Đinh Thị Lan Anh', 'Vũ Thanh Hằng', 'Lê Thị Soạn', 'Nguyễn Thị Mai Duyên',
    'Hoàng Thanh Thủy', 'Lê Thị Huyền Trân', 'Nguyễn Trần Ngọc Phương', 'Hồ Thị Thùy Dung'
]

print(f"{'CN':<15} | {'Họ tên':<24} | {'Chức danh':<8} | {'Giờ ngày':<8} | {'Giờ đêm':<8} | {'MiniKat DS':<10} | {'MiniKat CHT':<11} | {'WA':<10} | {'Dự án':<10} | {'KPI':<10} | {'Thưởng CK':<10} | {'Cận date':<9} | {'Trừ':<8} | Ghi chú")
print("-" * 170)

for r in range(3, ws_bl.max_row + 1):
    name = ws_bl.cell(r, 3).value
    if name in sample_staff:
        cn = str(ws_bl.cell(r, 2).value or '')
        cv = str(ws_bl.cell(r, 4).value or '')
        gn = ws_bl.cell(r, 5).value or 0
        gd = ws_bl.cell(r, 6).value or 0
        mk_ds = ws_bl.cell(r, 27).value or 0
        mk_cht = ws_bl.cell(r, 28).value or 0
        wa = ws_bl.cell(r, 29).value or 0
        da = ws_bl.cell(r, 30).value or 0
        kpi = ws_bl.cell(r, 31).value or 0
        ck = ws_bl.cell(r, 32).value or 0
        cd = ws_bl.cell(r, 33).value or 0
        tru = ws_bl.cell(r, 36).value or 0
        gc = str(ws_bl.cell(r, 42).value or '')
        print(f"{cn:<15} | {name:<24} | {cv:<8} | {gn:<8.1f} | {gd:<8.1f} | {mk_ds:<10,.0f} | {mk_cht:<11,.0f} | {wa:<10,.0f} | {da:<10,.0f} | {kpi:<10,.0f} | {ck:<10,.0f} | {cd:<9,.0f} | {tru:<8,.0f} | {gc}")

print("\n=== KIỂM TRA MINI-KAT HN & HCM ===")
for sname in ['MiniKat - HN', 'MiniKat - HCM']:
    ws = wb[sname]
    print(f"\n--- {sname} (Top 3 nhân viên bán nhiều nhất) ---")
    rows = []
    for r in range(2, 50):
        s = ws.cell(r, 1).value
        if s:
            tot = ws.cell(r, 10).value or 0
            rows.append((s, tot, ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 5).value))
    rows.sort(key=lambda x: x[1], reverse=True)
    for row in rows[:3]:
        print(f"  {row[0]}: Thưởng = {row[1]:,.0f} đ (PS: {row[2]}, KAT: {row[3]}, LadyRev: {row[4]:,.0f})")

print("\n=== KIỂM TRA THƯỞNG CK (50K/HỘP) ===")
ws_ck = wb['Thưởng CK']
for r in range(2, 20):
    n = ws_ck.cell(r, 2).value
    rew = ws_ck.cell(r, 3).value
    if rew and rew > 0:
        print(f"  {n}: Thưởng CK = {rew:,.0f} đ (Hệ số = {ws_ck.cell(r, 4).value}, Gốc = {ws_ck.cell(r, 5).value:,.0f})")
