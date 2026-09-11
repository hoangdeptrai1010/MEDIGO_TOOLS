import openpyxl, sys, io
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

hd_path = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb = openpyxl.load_workbook(hd_path, read_only=True)
ws = wb.active

rows_iter = ws.iter_rows(values_only=True)
header = next(rows_iter)

channels = Counter()
price_books = Counter()
customer_types = Counter()
notes = Counter()

for row in rows_iter:
    if row[8]: channels[str(row[8])] += 1
    if row[6]: price_books[str(row[6])] += 1
    if row[20]: notes[str(row[20])] += 1

wb.close()

print("=== CHANNELS (KÊNH BÁN) ===")
for k, v in channels.most_common(20):
    print(f"  {k}: {v}")

print("\n=== PRICE BOOKS (BẢNG GIÁ) ===")
for k, v in price_books.most_common(20):
    print(f"  {k}: {v}")

print("\n=== GHI CHÚ HÀNG HÓA ===")
for k, v in notes.most_common(30):
    print(f"  {k}: {v}")
