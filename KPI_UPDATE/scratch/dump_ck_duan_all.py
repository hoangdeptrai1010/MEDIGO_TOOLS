import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

for fname in ['thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', 'thang8/bangluong_thang8_hoanthien.xlsx', 'thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx']:
    print(f"\n=======================================================")
    print(f"FILE: {fname}")
    print(f"=======================================================")
    wb = openpyxl.load_workbook(fname, data_only=False)
    for sname in ['Thưởng CK', 'Dự án']:
        if sname in wb.sheetnames:
            ws = wb[sname]
            print(f"\n--- Sheet: {sname} (max_row={ws.max_row}, max_col={ws.max_column}) ---")
            for r in range(1, ws.max_row + 1):
                row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                if any(v is not None for v in row_vals):
                    print(f"Row {r:2d}: {row_vals}")

