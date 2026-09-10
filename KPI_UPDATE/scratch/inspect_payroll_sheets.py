import openpyxl, os

paths = [
    'thang7/BẢNG LƯƠNG THÁNG 7 2026.xlsx',
    'thang7/bangluong_thang7_hoanthien.xlsx',
    'thang8/BẢNG LƯƠNG THÁNG 8 2026.xlsx',
    'thang8/bangluong_thang8_hoanthien.xlsx',
]

for p in paths:
    if os.path.exists(p):
        try:
            wb = openpyxl.load_workbook(p, read_only=True)
            print(f"=== {p} ===")
            print("Sheets:", wb.sheetnames)
        except Exception as e:
            print(f"Error {p}: {e}")
    else:
        print(f"Not found: {p}")
