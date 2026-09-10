import openpyxl

# Check SKUs in invoices
wb = openpyxl.load_workbook('thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx', read_only=True, data_only=True)
ws = wb.active

target_skus = {
    'SP017162': 'Party Smart',
    'SP2725289': 'KAT Calcium', 'SP2725285': 'KAT Hair plus', 'SP2725291': 'KAT Liver pro',
    'SP2725287': 'KAT Super iq', 'SP2725353': 'KAT Maca 2', 'SP2725351': 'KAT Maca 6',
    'SP2723124': 'LadyCare Aloe', 'SP2723125': 'LadyCare Classic 190', 'SP2723127': 'LadyCare Classic 90',
    'SP2723883': 'AVC Foslin', 'SP2722826': 'AVC Trung thao', 'SP2722817': 'AVC Thanh nhiet', 'SP2725666': 'AVC Ceregold'
}

found_counts = {k: 0 for k in target_skus}
found_sellers = set()

# Header
row_iter = ws.iter_rows(values_only=True)
header = next(row_iter)
sku_col = None
qty_col = None
seller_col = None
note_col = None

for i, h in enumerate(header):
    h_str = str(h).strip().lower()
    if 'mã hàng' in h_str:
        sku_col = i
    elif 'số lượng' in h_str:
        qty_col = i
    elif 'người bán' in h_str:
        seller_col = i
    elif 'ghi chú' in h_str:
        note_col = i

print(f"sku_col={sku_col}, qty_col={qty_col}, seller_col={seller_col}, note_col={note_col}")

whatsapp_bills = []
for r in row_iter:
    if not r or len(r) <= max(sku_col, qty_col, seller_col):
        continue
    sku = str(r[sku_col]).strip() if r[sku_col] else ''
    if sku in target_skus:
        found_counts[sku] += float(r[qty_col] or 0)
        found_sellers.add(str(r[seller_col]).strip())
    
    if note_col and note_col < len(r) and r[note_col]:
        note_str = str(r[note_col]).lower()
        if 'whatsapp' in note_str:
            whatsapp_bills.append(r)

print("Found SKU sales in August:", {target_skus[k]: found_counts[k] for k in target_skus if found_counts[k] > 0})
print("Total WhatsApp bills found in notes:", len(whatsapp_bills))
if whatsapp_bills:
    print("Sample whatsapp bill:", [(whatsapp_bills[0][i], header[i]) for i in [0, 1, 2, seller_col, note_col] if i is not None])
