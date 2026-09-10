import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)
wb_f = openpyxl.load_workbook('thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)

print("=== 1. SHEET 'Thưởng CK' TRONG thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx ===")
ws = wb['Thưởng CK']
ws_f = wb_f['Thưởng CK']

for r in range(1, ws.max_row + 1):
    b = ws.cell(r, 2).value
    c = ws.cell(r, 3).value
    d = ws.cell(r, 4).value
    e = ws.cell(r, 5).value
    h = ws.cell(r, 8).value
    i = ws.cell(r, 9).value
    j = ws.cell(r, 10).value
    k = ws.cell(r, 11).value
    l = ws.cell(r, 12).value
    if r == 1 or b or h:
        print(f"Row {r:2d}: Trái (B={str(b):<23} | C={str(c):>10} | D={str(d):>6} | E={str(e):>10}) || Phải (H={str(h):<23} | I={str(i):>10} | J={str(j):>6} | K={str(k):>10} | L={str(l):>6})")

print("\n=== 2. SHEET 'Dự án' TRONG thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx ===")
ws_da = wb['Dự án']
for r in range(1, ws_da.max_row + 1):
    b = ws_da.cell(r, 2).value
    c = ws_da.cell(r, 3).value
    h = ws_da.cell(r, 8).value
    i = ws_da.cell(r, 9).value
    j = ws_da.cell(r, 10).value
    k = ws_da.cell(r, 11).value
    l = ws_da.cell(r, 12).value
    if r == 1 or b or h:
        print(f"Row {r:2d}: Trái (B={str(b):<23} | C={str(c):>10}) || Phải (H={str(h):<23} | I={str(i):>10} | J={str(j):>6} | K={str(k):>10} | L={str(l):>6})")

print("\n=== 3. CỘT AD (DỰ ÁN) VÀ CỘT AF (THƯỞNG CK) TRÊN BẢNG LƯƠNG TỔNG THÁNG 7 TARGET ===")
ws_bl = wb['BẢNG LƯƠNG']
for r in range(3, ws_bl.max_row + 1):
    name = ws_bl.cell(r, 3).value
    cn = ws_bl.cell(r, 2).value
    ad = ws_bl.cell(r, 30).value
    af = ws_bl.cell(r, 32).value
    if name:
        print(f"Row {r:2d}: {str(cn):<16} | {str(name):<24} | Cột AD (Dự án): {str(ad):>10} | Cột AF (Thưởng CK): {str(af):>10}")

