import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=False)
ws = wb['BẢNG LƯƠNG']

print("=== CONDITIONAL FORMATTING IN BẢNG LƯƠNG ===")
for cf in ws.conditional_formatting:
    print(f"Range: {cf.sqref}, Rules: {len(cf.rules)}")
    for rule in cf.rules:
        print(f"  Rule type: {rule.type}, formula: {rule.formula}, dxf: {rule.dxf}")

print("\n=== ROWS 76 TO 89 ===")
for r in range(76, 89):
    row_vals = [f"Col {c} ({ws.cell(r, c).coordinate}): {ws.cell(r, c).value}" for c in range(4, 11) if ws.cell(r, c).value is not None]
    print(f"Row {r}: {', '.join(row_vals)}")

wb.close()
