import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')

# 1. Load SKU commission catalogs from CK-HN and CK-HCM
ck_rate_hn = {}
ck_rate_hcm = {}

# HN
wb_hn = openpyxl.load_workbook('thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/CK-HN.xlsx', data_only=True)
for sname in wb_hn.sheetnames:
    ws = wb_hn[sname]
    hdr = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    sku_col = None
    rate_col = None
    for idx, h in enumerate(hdr, 1):
        if h and 'mã hàng' in str(h).lower():
            sku_col = idx
        if h and ('chiết khấu' in str(h).lower() or 'hoa hồng' in str(h).lower()):
            rate_col = idx
    if sku_col and rate_col:
        for r in range(2, ws.max_row + 1):
            sku = ws.cell(r, sku_col).value
            rate = ws.cell(r, rate_col).value
            if sku and rate and isinstance(rate, (int, float)):
                ck_rate_hn[str(sku).strip()] = rate

# HCM
wb_hcm = openpyxl.load_workbook('thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/CK-HCM.xlsx', data_only=True)
for sname in wb_hcm.sheetnames:
    ws = wb_hcm[sname]
    hdr = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    sku_col = None
    rate_col = None
    for idx, h in enumerate(hdr, 1):
        if h and 'mã hàng' in str(h).lower():
            sku_col = idx
        if h and ('chiết khấu' in str(h).lower() or 'hoa hồng' in str(h).lower()):
            rate_col = idx
    if sku_col and rate_col:
        for r in range(2, ws.max_row + 1):
            sku = ws.cell(r, sku_col).value
            rate = ws.cell(r, rate_col).value
            if sku and rate and isinstance(rate, (int, float)):
                ck_rate_hcm[str(sku).strip()] = rate

print(f"Loaded {len(ck_rate_hn)} HN SKU rates, {len(ck_rate_hcm)} HCM SKU rates.")

# 2. Scan DanhSachChiTietHoaDon_3182026.xlsx
inv_file = 'thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb_inv = openpyxl.load_workbook(inv_file, read_only=True, data_only=True)
ws_inv = wb_inv.active
iter_inv = ws_inv.iter_rows(values_only=True)
hdr = next(iter_inv)
h_map = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}

c_cn = h_map.get('chi nhánh', 0)
c_seller = h_map.get('người bán', 6)
c_sku = h_map.get('mã hàng', 28)
c_qty = h_map.get('số lượng', 35)

seller_comm = {}
HN_BRANCHES = ['hàng bông', 'đường láng', 'nt hàng bông', 'nt đường láng 247', 'nt đường láng']

for r in iter_inv:
    if not r or len(r) <= max(c_cn, c_seller, c_sku, c_qty):
        continue
    cn = str(r[c_cn]).strip().lower() if r[c_cn] else ''
    seller = str(r[c_seller]).strip() if r[c_seller] else ''
    sku = str(r[c_sku]).strip() if r[c_sku] else ''
    qty = r[c_qty] if isinstance(r[c_qty], (int, float)) else 0
    if not seller or qty <= 0:
        continue
    
    is_hn = any(b in cn for b in ['hàng bông', 'đường láng', 'láng'])
    rate = ck_rate_hn.get(sku, 0) if is_hn else ck_rate_hcm.get(sku, 0)
    if rate > 0:
        seller_comm[seller] = seller_comm.get(seller, 0) + (qty * rate)

print("\n=== TÍNH HOA HỒNG CHIẾT KHẤU THEO MÃ HÀNG (NẾU CÓ CHÍNH SÁCH HOA HỒNG SẢN PHẨM) ===")
for s, val in sorted(seller_comm.items(), key=lambda x: x[1], reverse=True)[:20]:
    print(f"  {s:<25}: {val:>12,f} đ")

