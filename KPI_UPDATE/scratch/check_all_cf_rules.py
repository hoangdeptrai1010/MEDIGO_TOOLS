import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/tinhcongnhungthuongchia.xlsx', data_only=False)
ws = wb['BẢNG LƯƠNG']

for cf in ws.conditional_formatting:
    for rule in cf.rules:
        # check if it affects rows 77-87 or col H
        print(f"Range: {cf.sqref} | Type: {rule.type} | Formula: {rule.formula} | Operator: {rule.operator} | text: {getattr(rule, 'text', None)}")

wb.close()
