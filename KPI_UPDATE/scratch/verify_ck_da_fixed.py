import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)

print("=== 1. KIỂM TRA SHEET DỰ ÁN (CÁC NHÂN SỰ CÓ THƯỞNG) ===")
ws_d = wb['Dự án']
count_da = 0
for r in range(2, ws_d.max_row + 1):
    cn = ws_d.cell(r, 1).value
    name = ws_d.cell(r, 2).value
    bonus = ws_d.cell(r, 3).value
    if name and bonus and bonus > 0:
        count_da += 1
        print(f"  {count_da:2d}. {cn:<16} | {name:<25} | Thưởng Dự án = {bonus:>10,f} đ")

print(f"\n--> Tổng số nhân sự có thưởng Dự án: {count_da} người (Chuẩn: 18 người)")

print("\n=== 2. KIỂM TRA SHEET THƯỞNG CK ===")
ws_ck = wb['Thưởng CK']
ck_non_zero = []
for r in range(2, ws_ck.max_row + 1):
    name = ws_ck.cell(r, 2).value
    c_val = ws_ck.cell(r, 3).value
    e_val = ws_ck.cell(r, 5).value
    k_val = ws_ck.cell(r, 11).value
    if name and ((c_val and c_val > 0) or (e_val and e_val > 0)):
        ck_non_zero.append((name, c_val, e_val))

print(f"Số nhân sự có thưởng CK riêng (Cột E > 0): {len(ck_non_zero)} (Chuẩn: 0 - không bị trùng lặp)")

print("\n=== 3. KIỂM TRA BẢNG LƯƠNG TỔNG (CỘT AD: DỰ ÁN & CỘT AF: THƯỞNG CK) ===")
ws_bl = wb['BẢNG LƯƠNG']
for r in range(3, ws_bl.max_row + 1):
    name = ws_bl.cell(r, 3).value
    cn = ws_bl.cell(r, 2).value
    ad = ws_bl.cell(r, 30).value or 0
    af = ws_bl.cell(r, 32).value or 0
    if ad > 0 or af > 0:
        print(f"Row {r:2d}: {cn:<16} | {name:<25} | Cột AD (Dự án): {ad:>10,f} đ | Cột AF (Thưởng CK): {af:>10,f} đ")
