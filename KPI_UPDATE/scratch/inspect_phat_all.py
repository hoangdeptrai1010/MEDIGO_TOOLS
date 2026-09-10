import openpyxl

wb = openpyxl.load_workbook('thang8/DATA/TongSoPhatViPhamVaThuongNgay_thang8.xlsx', read_only=True, data_only=True)
ws = wb.active
count = 0
for i, r in enumerate(ws.iter_rows(values_only=True)):
    if any(x is not None for x in r):
        count += 1
        if count <= 15:
            print(f"r{i+1}: {r}")
print(f"Total non-empty rows: {count}")
