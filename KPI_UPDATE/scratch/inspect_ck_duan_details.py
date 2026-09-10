import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

def dump_sheet(wb, sheet_name, max_r=50, max_c=15):
    if sheet_name not in wb.sheetnames:
        print(f"Sheet {sheet_name} NOT FOUND")
        return
    ws = wb[sheet_name]
    print(f"--- Sheet: {sheet_name} (max_row={ws.max_row}, max_col={ws.max_column}) ---")
    for r in range(1, min(max_r, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, min(max_c, ws.max_column + 1))]
        if any(v is not None for v in row_vals):
            print(f"Row {r:2d}: {row_vals}")

for f in ['thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', 'thang8/bangluong_thang8_hoanthien.xlsx', 'thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx']:
    print(f"\n==================== FILE: {f} ====================")
    wb = openpyxl.load_workbook(f, data_only=False)
    dump_sheet(wb, 'Thưởng CK', max_r=30, max_c=15)
    dump_sheet(wb, 'Dự án', max_r=30, max_c=15)
    # Check formulas in BẢNG LƯƠNG for Thưởng CK and Dự án columns
    if 'BẢNG LƯƠNG' in wb.sheetnames:
        ws_bl = wb['BẢNG LƯƠNG']
        # find header row
        header_row = 2
        headers = [ws_bl.cell(header_row, c).value for c in range(1, ws_bl.max_column + 1)]
        print("\nBẢNG LƯƠNG Headers:")
        for idx, h in enumerate(headers, 1):
            if h:
                print(f"  Col {idx} ({openpyxl.utils.get_column_letter(idx)}): {repr(h)}")
        # Print first 5 data rows for key columns
        print("\nBẢNG LƯƠNG Sample Rows (formulas):")
        for r in range(3, 8):
            vals = {openpyxl.utils.get_column_letter(c): ws_bl.cell(r, c).value for c in range(1, min(40, ws_bl.max_column + 1)) if ws_bl.cell(r, c).value is not None}
            print(f"Row {r}: {ws_bl.cell(r, 2).value} ({ws_bl.cell(r, 3).value}): {vals}")
