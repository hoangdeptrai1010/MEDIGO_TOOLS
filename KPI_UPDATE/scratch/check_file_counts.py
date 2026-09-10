import openpyxl, os, sys

sys.stdout.reconfigure(encoding='utf-8')
files = [
    'thang8/tinhcongnhungthuongchia.xlsx',
    'thang8/BANGLUONGTHANG8.xlsx',
    'thang8/BANGLUONGTHANG8_HOANG_.xlsx',
    'thang8/bangluong_thang8_hoanthien.xlsx',
    'thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx'
]

for fp in files:
    if os.path.exists(fp):
        wb = openpyxl.load_workbook(fp, data_only=False)
        ws = wb['BẢNG LƯƠNG']
        # Count staff rows between row 3 and row 75
        staff_cnt = 0
        for r in range(3, 75):
            name = ws.cell(r, 3).value
            if name and str(name).strip() and str(ws.cell(r, 4).value).strip() != 'Tổng':
                staff_cnt += 1
        print(f"File '{fp}': max_row={ws.max_row}, staff rows count={staff_cnt}")
