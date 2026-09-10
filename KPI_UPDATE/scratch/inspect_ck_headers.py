import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

for fn in ['thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', 'thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', 'thang8/bangluong_thang8_hoanthien.xlsx']:
    print(f"\n==================== {fn} ====================")
    wb = openpyxl.load_workbook(fn, data_only=False)
    for sname in ['Thưởng CK', 'Dự án']:
        ws = wb[sname]
        print(f"\n--- Sheet: {sname} ---")
        for r in range(1, 6):
            row_vals = [f"Col {openpyxl.utils.get_column_letter(c)}: {repr(ws.cell(r, c).value)}" for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None]
            print(f"Row {r}: {', '.join(row_vals)}")
