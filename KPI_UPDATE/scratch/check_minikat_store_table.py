import openpyxl

wb = openpyxl.load_workbook('thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx', data_only=False)
for sn in ['MiniKat - HN', 'MiniKat - HCM']:
    ws = wb[sn]
    print(f"=== {sn} STORE TABLE (Cols 12-20) ===")
    for r in range(1, 15):
        vals = [f"Col{c}: {ws.cell(r, c).value}" for c in range(12, 20) if ws.cell(r, c).value is not None]
        if vals:
            print(f"  r{r}: {vals}")
