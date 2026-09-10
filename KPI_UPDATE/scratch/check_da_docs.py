import sys, openpyxl, glob, os
sys.stdout.reconfigure(encoding='utf-8')

print("=== CHECKING JULY DỰ ÁN T7 ===")
wb_kpi7 = openpyxl.load_workbook('baocaokpi_thang7_hoanthien.xlsx', data_only=True)
if 'Dự án T7' in wb_kpi7.sheetnames:
    ws = wb_kpi7['Dự án T7']
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 14)]
        if any(v is not None for v in vals):
            total_da = ws.cell(r, 13).value
            if r <= 3 or (isinstance(total_da, (int, float)) and total_da > 0):
                print(f"July DA T7 Row {r}: {vals[0]} | {vals[1]} | {vals[2]} | {vals[9]} | {vals[10]} | {vals[11]} | {vals[12]}")

print("\n=== SEARCHING ALL MD / DOCX / TEXT FILES FOR 'Thưởng CK' or 'Dự án' ===")
for root, dirs, files in os.walk('.'):
    if '.git' in root or '.cache' in root:
        continue
    for f in files:
        if f.endswith(('.md', '.txt', '.json')):
            p = os.path.join(root, f)
            try:
                with open(p, 'r', encoding='utf-8', errors='ignore') as fl:
                    content = fl.read()
                    if 'Thưởng CK' in content or 'Dự án' in content or 'thuong ck' in content.lower():
                        print(f"Found in {p}")
            except:
                pass

