import openpyxl, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook('d:/MEDIGO/KPI_UPDATE/thang8/output/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws = wb['whatsapp']

print("=== INSPECTING COLUMNS F TO Y IN SHEET whatsapp ===")
for r in range(1, ws.max_row + 1):
    vals = [ws.cell(r, c).value for c in range(5, ws.max_column + 1)]
    if any(vals):
        print(f"Row {r:3d}: {vals}")

wb.close()
