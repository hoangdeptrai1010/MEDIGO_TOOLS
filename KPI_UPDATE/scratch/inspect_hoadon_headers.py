import openpyxl, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook('d:/MEDIGO/KPI_UPDATE/thang7/DATAKIOT/hoadon.xlsx', read_only=True)
ws = wb.active
row1 = next(ws.iter_rows(values_only=True))
print("=== HEADERS IN hoadon.xlsx ===")
for i, h in enumerate(row1):
    print(f"Col {i:2d}: {h}")
wb.close()
