import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('thang8/tinhcongnhungthuongchia.xlsx', data_only=True)
ws = wb['BẢNG LƯƠNG']

for r in range(76, 89):
    for c in range(5, 10):
        cell = ws.cell(r, c)
        fill = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor else None
        print(f"Cell {cell.coordinate}: val='{cell.value}', fill={fill}")

wb.close()
