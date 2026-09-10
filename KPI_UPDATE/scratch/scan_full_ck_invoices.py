import sys, os, openpyxl
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

print("=" * 80)
print("=== BÁO CÁO TỔNG QUAN CHI TIẾT QUÉT TOÀN BỘ HÓA ĐƠN THÁNG 8 VỀ CK ===")
print("=" * 80)

# 1. Nạp danh mục CK, Combo, Hoa hồng
hn_ck_catalog = {} # sku -> {'name': ..., 'rate': ..., 'type': ...}
hcm_ck_catalog = {}

# Hà Nội
p_hn = 'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/CK-HN.xlsx'
wb_hn = openpyxl.load_workbook(p_hn, data_only=True)
ws = wb_hn['danhmucchietkhau-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    name = ws.cell(r, 3).value
    rate = ws.cell(r, 6).value
    if sku:
        hn_ck_catalog[str(sku).strip()] = {'name': str(name).strip() if name else '', 'rate': rate or 0, 'type': 'CK'}

ws = wb_hn['danhmuccombolieu-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 4).value
    name = ws.cell(r, 5).value
    rate = ws.cell(r, 10).value
    if sku:
        hn_ck_catalog[str(sku).strip()] = {'name': str(name).strip() if name else '', 'rate': rate or 0, 'type': 'Combo'}

ws = wb_hn['danhmuchoahong']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    name = ws.cell(r, 3).value
    rate = ws.cell(r, 6).value
    if sku:
        hn_ck_catalog[str(sku).strip()] = {'name': str(name).strip() if name else '', 'rate': rate or 0, 'type': 'NY3'}

# HCM
p_hcm = 'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/CK-HCM.xlsx'
wb_hcm = openpyxl.load_workbook(p_hcm, data_only=True)
ws = wb_hcm['danhmucck-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    name = ws.cell(r, 3).value
    rate = ws.cell(r, 6).value
    if sku:
        hcm_ck_catalog[str(sku).strip()] = {'name': str(name).strip() if name else '', 'rate': rate or 0, 'type': 'CK'}

ws = wb_hcm['danhmuccombolieu-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 4).value
    name = ws.cell(r, 5).value
    rate = ws.cell(r, 10).value
    if sku:
        hcm_ck_catalog[str(sku).strip()] = {'name': str(name).strip() if name else '', 'rate': rate or 0, 'type': 'Combo'}

ws = wb_hcm['danhmuchoahong']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    name = ws.cell(r, 3).value
    rate = ws.cell(r, 6).value
    if sku:
        hcm_ck_catalog[str(sku).strip()] = {'name': str(name).strip() if name else '', 'rate': rate or 0, 'type': 'NY3'}

print(f"-> Đã nạp danh mục: HN có {len(hn_ck_catalog)} mã, HCM có {len(hcm_ck_catalog)} mã.")

# 2. Quét hóa đơn trả hàng (nếu có để trừ)
ret_file = 'thang8/DATA/DanhSachChiTietTraHang_3182026.xlsx'
ret_dict = defaultdict(float) # (seller, sku) -> qty
if os.path.exists(ret_file):
    wb_ret = openpyxl.load_workbook(ret_file, read_only=True, data_only=True)
    ws_ret = wb_ret.active
    it = ws_ret.iter_rows(values_only=True)
    hdr = next(it)
    h_m = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}
    c_cn, c_seller, c_sku, c_qty = h_m.get('chi nhánh', 0), h_m.get('người bán', 6), h_m.get('mã hàng', 28), h_m.get('số lượng', 35)
    for r in it:
        if r and len(r) > max(c_cn, c_seller, c_sku, c_qty):
            s = str(r[c_seller]).strip() if r[c_seller] else ''
            sku = str(r[c_sku]).strip() if r[c_sku] else ''
            q = float(r[c_qty]) if r[c_qty] else 0
            if s and sku:
                ret_dict[(s, sku)] += q

# 3. Quét hóa đơn chi tiết
inv_file = 'thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb_inv = openpyxl.load_workbook(inv_file, read_only=True, data_only=True)
ws_inv = wb_inv.active
it = ws_inv.iter_rows(values_only=True)
hdr = next(it)
h_m = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}

c_cn = h_m.get('chi nhánh', 0)
c_code = h_m.get('mã hóa đơn', 1)
c_date = h_m.get('thời gian', 2)
c_seller = h_m.get('người bán', 6)
c_sku = h_m.get('mã hàng', 28)
c_ten = h_m.get('tên hàng', 30)
c_qty = h_m.get('số lượng', 35)
c_price = h_m.get('giá bán', 36)
c_tt = h_m.get('thành tiền', 37)

seller_data = defaultdict(lambda: {
    'branch': set(),
    'rev_ck': 0.0, 'rev_combo': 0.0, 'rev_ny3': 0.0,
    'comm_ck': 0.0, 'comm_combo': 0.0, 'comm_ny3': 0.0,
    'qty_ck': 0.0, 'qty_combo': 0.0, 'qty_ny3': 0.0,
    'bills_ck': set(),
})

for r in it:
    if not r or len(r) <= max(c_cn, c_seller, c_sku, c_qty):
        continue
    cn_raw = str(r[c_cn]).strip() if r[c_cn] else ''
    seller = str(r[c_seller]).strip() if r[c_seller] else ''
    sku = str(r[c_sku]).strip() if r[c_sku] else ''
    qty = float(r[c_qty]) if r[c_qty] and isinstance(r[c_qty], (int, float)) else 0.0
    price = float(r[c_price]) if r[c_price] and isinstance(r[c_price], (int, float)) else 0.0
    tt = float(r[c_tt]) if r[c_tt] and isinstance(r[c_tt], (int, float)) else qty * price
    bill = str(r[c_code]).strip() if r[c_code] else ''

    if not seller or qty <= 0:
        continue

    is_hn = any(b in cn_raw.lower() for b in ['hàng bông', 'đường láng', 'láng'])
    cat = hn_ck_catalog if is_hn else hcm_ck_catalog

    if sku in cat:
        info = cat[sku]
        stype = info['type']
        rate = info['rate'] if isinstance(info['rate'], (int, float)) else 0.0

        seller_data[seller]['branch'].add(cn_raw)
        
        if stype == 'CK':
            seller_data[seller]['rev_ck'] += tt
            seller_data[seller]['qty_ck'] += qty
            seller_data[seller]['comm_ck'] += (qty * rate)
            seller_data[seller]['bills_ck'].add(bill)
        elif stype == 'Combo':
            seller_data[seller]['rev_combo'] += tt
            seller_data[seller]['qty_combo'] += qty
            seller_data[seller]['comm_combo'] += (qty * rate)
        elif stype == 'NY3':
            seller_data[seller]['rev_ny3'] += tt
            seller_data[seller]['qty_ny3'] += qty
            seller_data[seller]['comm_ny3'] += (qty * rate)

print("\n" + "=" * 135)
print(f"{'STT':<4} | {'Họ và tên':<24} | {'Chi nhánh':<16} | {'Doanh số CK':<13} | {'DS Combo':<11} | {'Hoa hồng CK':<12} | {'HH Combo':<10} | {'HH NY3':<10} | {'TỔNG HOA HỒNG (CK)':<18}")
print("-" * 135)

stt = 0
sorted_sellers = sorted(seller_data.items(), key=lambda x: (x[1]['comm_ck'] + x[1]['comm_combo'] + x[1]['comm_ny3']), reverse=True)

for seller, d in sorted_sellers:
    stt += 1
    cn_str = ", ".join(list(d['branch'])[:1])
    tot_comm = d['comm_ck'] + d['comm_combo'] + d['comm_ny3']
    print(f"{stt:<4} | {seller:<24} | {cn_str:<16} | {d['rev_ck']:>13,.0f} | {d['rev_combo']:>11,.0f} | {d['comm_ck']:>12,.0f} | {d['comm_combo']:>10,.0f} | {d['comm_ny3']:>10,.0f} | {tot_comm:>18,.0f} đ")

