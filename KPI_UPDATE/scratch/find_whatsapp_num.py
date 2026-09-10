import os, openpyxl

target = 91605700
for root, dirs, files in os.walk('.'):
    for f in files:
        if f.endswith('.xlsx') and not f.startswith('~$') and not 'test' in f:
            p = os.path.join(root, f)
            try:
                wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                for s in wb.sheetnames:
                    ws = wb[s]
                    for r in ws.iter_rows(values_only=True):
                        if any(x == target or x == 91605700.0 for x in r if isinstance(x, (int, float))):
                            print(f"Found in {p} -> Sheet {s}")
                            break
            except:
                pass
