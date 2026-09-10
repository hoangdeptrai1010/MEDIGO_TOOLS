import pathlib, sys, openpyxl

sys.stdout.reconfigure(encoding='utf-8')

p = pathlib.Path('thang8/Pharmacy_retail_Store_KPIs_August 2026')
for f in p.rglob('*.xlsx'):
    if any(k in f.name.lower() for k in ['kat', 'ladycare', 'party']):
        print(f"\n=======================================================")
        print(f"FILE: {f.name}")
        print(f"PATH: {f}")
        print(f"=======================================================")
        wb = openpyxl.load_workbook(f, data_only=True)
        print("Sheets:", wb.sheetnames)
        for sname in wb.sheetnames:
            ws = wb[sname]
            print(f"\n--- Sheet: {sname} (max_row={ws.max_row}, max_col={ws.max_column}) ---")
            for r in range(1, min(15, ws.max_row + 1)):
                vals = [ws.cell(r, c).value for c in range(1, min(15, ws.max_column + 1))]
                if any(vals):
                    print(f"Row {r:2d}: {vals}")
