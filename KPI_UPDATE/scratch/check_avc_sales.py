import openpyxl, sys
sys.stdout.reconfigure(encoding='utf-8')

wb_inv = openpyxl.load_workbook('thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx', read_only=True)
ws_inv = wb_inv.active
avc_skus = {'SP2723883', 'SP2722826', 'SP2722817', 'SP2725666'}
avc_sales_by_seller = {}
for r, row in enumerate(ws_inv.iter_rows(values_only=True)):
    if r == 0: continue
    sku = str(row[14]).strip() if row[14] else ''
    if sku in avc_skus:
        branch = str(row[0]).strip() if row[0] else ''
        seller = str(row[7]).strip() if row[7] else ''
        qty = float(row[21] or 0)
        tt = float(row[26] or 0)
        if seller not in avc_sales_by_seller:
            avc_sales_by_seller[seller] = {'branch': branch, 'qty': 0.0, 'rev': 0.0}
        avc_sales_by_seller[seller]['qty'] += qty
        avc_sales_by_seller[seller]['rev'] += tt

print('=== AVC Sales in August Invoices ===')
for s, d in sorted(avc_sales_by_seller.items(), key=lambda x: x[1]['rev'], reverse=True):
    b = d['branch']
    q = d['qty']
    rev = d['rev']
    print(f"{b:<25} | {s:<25}: Qty={q:>4.0f} | Revenue={rev:>12,.0f}")
