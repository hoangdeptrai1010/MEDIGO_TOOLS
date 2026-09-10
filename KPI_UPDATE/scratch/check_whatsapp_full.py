import openpyxl

wb = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=False)
ws = wb['whatsapp']
print("=== WHATSAPP in thang8/bangluong_thang8_hoanthien.xlsx ===")
for r in range(1, 20):
    vals = [ws.cell(r, c).value for c in range(1, 6)]
    if any(x is not None for x in vals):
        print(f"r{r}: {vals}")
