import openpyxl

wb = openpyxl.load_workbook('thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx', data_only=False)

for sname in ['MiniKat - HN', 'MiniKat - HCM']:
    ws = wb[sname]
    print(f"=== {sname} ===")
    for r in range(1, 15):
        vals = [ws.cell(r, c).value for c in range(1, 20)]
        if any(x is not None for x in vals):
            print(f"r{r}: {vals[:11]}")
            if len(vals) > 12 and any(x is not None for x in vals[12:]):
                print(f"   extra: {vals[12:]}")
