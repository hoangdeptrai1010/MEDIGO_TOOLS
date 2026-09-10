import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws = wb['BẢNG LƯƠNG']

print("=== TỔNG KẾT CÁC CỘT THƯỞNG TRONG BẢNG LƯƠNG THÁNG 8 ===")
col_sums = {
    'MiniKat DS (Cột AA)': 0,
    'MiniKat CHT (Cột AB)': 0,
    'WhatsApp (Cột AC)': 0,
    'Dự án (Cột AD)': 0,
    'KPI (Cột AE)': 0,
    'Thưởng CK (Cột AF)': 0,
    'Cận date (Cột AG)': 0,
    'Thưởng Maps (Cột AH)': 0,
    'Tổng giảm trừ (Cột AJ)': 0,
}

for r in range(3, 59): # 56 staff from row 3 to 58
    name = ws.cell(r, 3).value
    if not name:
        continue
    col_sums['MiniKat DS (Cột AA)'] += ws.cell(r, 27).value or 0
    col_sums['MiniKat CHT (Cột AB)'] += ws.cell(r, 28).value or 0
    col_sums['WhatsApp (Cột AC)'] += ws.cell(r, 29).value or 0
    col_sums['Dự án (Cột AD)'] += ws.cell(r, 30).value or 0
    col_sums['KPI (Cột AE)'] += ws.cell(r, 31).value or 0
    col_sums['Thưởng CK (Cột AF)'] += ws.cell(r, 32).value or 0
    col_sums['Cận date (Cột AG)'] += ws.cell(r, 33).value or 0
    col_sums['Thưởng Maps (Cột AH)'] += ws.cell(r, 34).value or 0
    col_sums['Tổng giảm trừ (Cột AJ)'] += ws.cell(r, 36).value or 0

for col, val in col_sums.items():
    print(f"  {col:<25}: {val:>15,f} đ")
