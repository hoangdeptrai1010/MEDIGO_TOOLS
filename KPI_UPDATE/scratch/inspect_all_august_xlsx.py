import os, openpyxl, zipfile, xml.etree.ElementTree as ET

def inspect_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    res = []
    res.append(f"Sheets: {wb.sheetnames}")
    for sn in wb.sheetnames:
        ws = wb[sn]
        res.append(f"  [{sn}] ({ws.max_row}x{ws.max_column}):")
        for r in range(1, min(10, ws.max_row + 1)):
            row_v = [ws.cell(r, c).value for c in range(1, min(10, ws.max_column + 1))]
            if any(x is not None for x in row_v):
                res.append(f"    r{r}: {row_v}")
    return '\n'.join(res)

root_dir = 'thang8/Pharmacy_retail_Store_KPIs_August 2026'
for r, d, files in os.walk(root_dir):
    for f in sorted(files):
        if f.endswith('.xlsx'):
            full = os.path.join(r, f)
            print("="*70)
            print("FILE:", f)
            print("="*70)
            print(inspect_xlsx(full))
