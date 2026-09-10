import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING THÁNG 8 ORIGINAL BẢNG LƯƠNG ===")
wb_orig = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=True)
ws_bl = wb_orig['BẢNG LƯƠNG']
for r in range(3, ws_bl.max_row + 1):
    name = ws_bl.cell(r, 3).value
    branch = ws_bl.cell(r, 2).value
    ad = ws_bl.cell(r, 30).value # Cột AD: Dự án
    af = ws_bl.cell(r, 32).value # Cột AF: Thưởng CK
    if (ad is not None and ad != 0) or (af is not None and af != 0):
        print(f"Orig Row {r:2d}: {branch} | {name} | Col AD (Dự án): {ad} | Col AF (Thưởng CK): {af}")

print("\n=== CHECKING THÁNG 7 ORIGINAL BẢNG LƯƠNG ===")
wb_j_orig = openpyxl.load_workbook('thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=True)
ws_j_bl = wb_j_orig['BẢNG LƯƠNG']
for r in range(3, ws_j_bl.max_row + 1):
    name = ws_j_bl.cell(r, 3).value
    branch = ws_j_bl.cell(r, 2).value
    ad = ws_j_bl.cell(r, 30).value # Cột AD: Dự án
    af = ws_j_bl.cell(r, 32).value # Cột AF: Thưởng CK
    if (ad is not None and ad != 0) or (af is not None and af != 0):
        print(f"July Row {r:2d}: {branch} | {name} | Col AD (Dự án): {ad} | Col AF (Thưởng CK): {af}")
