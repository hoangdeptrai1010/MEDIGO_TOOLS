import sys, os, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING FILES IN thang7/Pharmacy_Retail_store_KPIs_July_2026 ===")
for root, dirs, files in os.walk('thang7/Pharmacy_Retail_store_KPIs_July_2026'):
    for f in files:
        p = os.path.join(root, f)
        print(f"\n--- FILE: {p} ---")
        if f.endswith('.xlsx'):
            wb = openpyxl.load_workbook(p, data_only=True)
            print("Sheets:", wb.sheetnames)
            for s in wb.sheetnames:
                ws = wb[s]
                print(f"  Sheet: {s} (rows={ws.max_row}, cols={ws.max_column})")
                for r in range(1, min(6, ws.max_row + 1)):
                    row_vals = [ws.cell(r, c).value for c in range(1, min(15, ws.max_column + 1))]
                    if any(v is not None for v in row_vals):
                        print(f"    Row {r}: {row_vals}")

