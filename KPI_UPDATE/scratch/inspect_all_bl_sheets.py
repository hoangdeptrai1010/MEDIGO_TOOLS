import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

for f in ['thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', 'thang8/bangluong_thang8_hoanthien.xlsx', 'thang7/BẢNG LƯƠNG THÁNG 7.xlsx']:
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        print(f"\n==================== FILE: {f} ====================")
        print("Sheets:", wb.sheetnames)
        for s in wb.sheetnames:
            ws = wb[s]
            # find first row with values
            for r in range(1, 10):
                row_vals = [ws.cell(r, c).value for c in range(1, min(50, ws.max_column + 1))]
                non_empty = [v for v in row_vals if v is not None]
                if len(non_empty) >= 3:
                    print(f"[{s}] Row {r}: {row_vals[:20]}")
                    break
    except Exception as e:
        print(f"Error {f}: {e}")
