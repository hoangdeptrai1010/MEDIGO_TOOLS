import openpyxl, sys, io
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

hd_path = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb = openpyxl.load_workbook(hd_path, read_only=True)
ws = wb.active

rows_iter = ws.iter_rows(values_only=True)
header = next(rows_iter)

target_sellers = {
    'Ngô Thị Thanh Thắm': 67046300,
    'Phạm Thị Nghĩa Hương': 8501100,
    'Trần Thiên Phát': 9377800,
    'Hoàng Thanh Thủy': 92439700,
    'Lê Thị Huyền Trân': 58824400,
    'Đinh Thị Khánh Ly': 19489000
}

# Group all invoices by seller and look at notes, customer codes, channels, etc.
seller_invoices = defaultdict(lambda: defaultdict(lambda: {'can_tra': 0, 'channel': '', 'time': '', 'branch': '', 'customer': '', 'items': []}))

for row in rows_iter:
    seller = str(row[7] or '').strip()
    if seller in target_sellers:
        ma_hd = str(row[1] or '').strip()
        can_tra = float(row[12] or 0)
        channel = str(row[8] or '').strip()
        time = str(row[2] or '')
        branch = str(row[0] or '')
        cust = str(row[5] or '')
        sku = str(row[14] or '')
        ten = str(row[16] or '')
        sl = float(row[21] or 1)
        tt = float(row[26] or 0)
        note = str(row[20] or '')
        
        inv = seller_invoices[seller][ma_hd]
        inv['can_tra'] = can_tra
        inv['channel'] = channel
        inv['time'] = time
        inv['branch'] = branch
        inv['customer'] = cust
        inv['items'].append({'sku': sku, 'ten': ten, 'sl': sl, 'tt': tt, 'note': note})

wb.close()

for seller, target_rev in target_sellers.items():
    invs = seller_invoices[seller]
    print(f"\n=== SELLER: {seller} (Target WhatsApp Rev: {target_rev:,.0f} đ, Total Invoices: {len(invs)}) ===")
    total_seller_rev = sum(inv['can_tra'] for inv in invs.values())
    print(f"  Total Invoices Revenue in Month 8: {total_seller_rev:,.0f} đ")
    
    # Check if any combination of invoices matches target_rev or if there are specific customer codes
    cust_counter = Counter(inv['customer'] for inv in invs.values())
    print(f"  Top Customers: {cust_counter.most_common(5)}")
