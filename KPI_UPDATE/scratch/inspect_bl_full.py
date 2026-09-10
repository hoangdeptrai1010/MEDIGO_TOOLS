import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb7 = openpyxl.load_workbook('thang7/target/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)
wb8 = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=False)

ws7 = wb7['BẢNG LƯƠNG']
ws8 = wb8['BẢNG LƯƠNG']

print(f"{'Col':<4} | {'Letter':<6} | {'Header T7':<35} | {'Header T8':<35} | {'Match?':<6}")
print("-" * 95)
max_c = max(ws7.max_column, ws8.max_column)
for c in range(1, max_c + 1):
    h7 = str(ws7.cell(2, c).value or '').replace('\n', ' ')
    h8 = str(ws8.cell(2, c).value or '').replace('\n', ' ')
    if not h7 and not h8:
        continue
    c_let = openpyxl.utils.get_column_letter(c)
    m = "YES" if h7 == h8 else "DIFF"
    print(f"{c:<4} | {c_let:<6} | {h7:<35} | {h8:<35} | {m}")
