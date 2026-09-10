import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

for p in [
    'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/CK-HN.xlsx',
    'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/CK-HCM.xlsx'
]:
    print(f"\n==================== {p} ====================")
    wb = openpyxl.load_workbook(p, data_only=True)
    for s in wb.sheetnames:
        ws = wb[s]
        print(f"\n--- Sheet: {s} (rows={ws.max_row}, cols={ws.max_column}) ---")
        for r in range(1, 10):
            vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if any(v is not None for v in vals):
                print(f"Row {r}: {vals[:10]}")
