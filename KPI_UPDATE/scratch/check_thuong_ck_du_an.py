import openpyxl, glob

print("=== CHECKING SHEETS ===")
for p in glob.glob('thang8/*.xlsx') + glob.glob('*.xlsx'):
    try:
        wb = openpyxl.load_workbook(p, read_only=True)
        print(f"{p}: {wb.sheetnames}")
    except Exception as e:
        pass
