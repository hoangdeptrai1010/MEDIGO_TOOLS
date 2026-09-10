import openpyxl

wb = openpyxl.load_workbook('thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)

for name in wb.sheetnames:
    ws = wb[name]
    print(f"=== SHEET: {name} (max_row={ws.max_row}, max_column={ws.max_column}) ===")
    # Print first 3 rows
    for r in range(1, min(6, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, min(15, ws.max_column + 1))]
        if any(v is not None for v in row_vals):
            print(f"  Row {r}: {row_vals[:10]}")
