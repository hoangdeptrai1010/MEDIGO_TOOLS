import sys, os, openpyxl
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

# 1. Load SKU catalogs
hn_ck = {}
hcm_ck = {}

# HN
wb_hn = openpyxl.load_workbook('thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/CK-HN.xlsx', data_only=True)
ws = wb_hn['danhmucchietkhau-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    rate = ws.cell(r, 6).value
    if sku and rate and isinstance(rate, (int, float)):
        hn_ck[str(sku).strip()] = ('CK', float(rate))

ws = wb_hn['danhmuccombolieu-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 4).value
    rate = ws.cell(r, 10).value
    if sku and rate and isinstance(rate, (int, float)):
        hn_ck[str(sku).strip()] = ('Combo', float(rate))

ws = wb_hn['danhmuchoahong']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    rate = ws.cell(r, 6).value
    if sku and rate and isinstance(rate, (int, float)):
        hn_ck[str(sku).strip()] = ('NY3', float(rate))

# HCM
wb_hcm = openpyxl.load_workbook('thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/CK-HCM.xlsx', data_only=True)
ws = wb_hcm['danhmucck-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    rate = ws.cell(r, 6).value
    if sku and rate and isinstance(rate, (int, float)):
        hcm_ck[str(sku).strip()] = ('CK', float(rate))

ws = wb_hcm['danhmuccombolieu-duan']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 4).value
    rate = ws.cell(r, 10).value
    if sku and rate and isinstance(rate, (int, float)):
        hcm_ck[str(sku).strip()] = ('Combo', float(rate))

ws = wb_hcm['danhmuchoahong']
for r in range(2, ws.max_row + 1):
    sku = ws.cell(r, 2).value
    rate = ws.cell(r, 6).value
    if sku and rate and isinstance(rate, (int, float)):
        hcm_ck[str(sku).strip()] = ('NY3', float(rate))

# 2. Quét hóa đơn chi tiết
inv_file = 'thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb_inv = openpyxl.load_workbook(inv_file, read_only=True, data_only=True)
ws_inv = wb_inv.active
it = ws_inv.iter_rows(values_only=True)
hdr = next(it)
h_m = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}

c_cn = h_m.get('chi nhánh', 0)
c_seller = h_m.get('người bán', 6)
c_sku = h_m.get('mã hàng', 28)
c_qty = h_m.get('số lượng', 35)

# Staff commission: seller -> {'CK': 0, 'Combo': 0, 'NY3': 0}
comm_by_seller = defaultdict(lambda: {'CK': 0.0, 'Combo': 0.0, 'NY3': 0.0})

for r in it:
    if not r or len(r) <= max(c_cn, c_seller, c_sku, c_qty):
        continue
    cn = str(r[c_cn]).strip().lower() if r[c_cn] else ''
    seller = str(r[c_seller]).strip() if r[c_seller] else ''
    sku = str(r[c_sku]).strip() if r[c_sku] else ''
    qty = float(r[c_qty]) if r[c_qty] and isinstance(r[c_qty], (int, float)) else 0.0

    if not seller or qty <= 0:
        continue

    is_hn = any(b in cn for b in ['hàng bông', 'đường láng', 'láng'])
    cat = hn_ck if is_hn else hcm_ck

    if sku in cat:
        stype, rate = cat[sku]
        comm_by_seller[seller][stype] += (qty * rate)

# 3. Đọc danh sách 56 nhân sự từ BẢNG LƯƠNG
wb_bl = openpyxl.load_workbook('thang8/bangluong_thang8_hoanthien.xlsx', data_only=True)
ws_bl = wb_bl['BẢNG LƯƠNG']

print("\n=== DANH SÁCH THƯỞNG CHIẾT KHẤU (CK) CỦA TOÀN BỘ 56 NHÂN SỰ BẢNG LƯƠNG THÁNG 8 ===")
print(f"{'STT':<4} | {'Chi nhánh':<15} | {'Họ và tên':<24} | {'Chức vụ':<8} | {'HH Chiết Khấu':<14} | {'HH Combo':<10} | {'HH NY3':<10} | {'TỔNG THƯỞNG CK':<15}")
print("-" * 115)

total_ck_all = 0.0
stt = 0
for r in range(3, ws_bl.max_row + 1):
    name = ws_bl.cell(r, 3).value
    cn = ws_bl.cell(r, 2).value
    cv = ws_bl.cell(r, 4).value
    if not name or not str(name).strip():
        continue
    stt += 1
    n_str = str(name).strip()
    c_info = comm_by_seller.get(n_str, {'CK': 0.0, 'Combo': 0.0, 'NY3': 0.0})
    ck_val = c_info['CK']
    combo_val = c_info['Combo']
    ny3_val = c_info['NY3']
    tot_ck = ck_val + combo_val + ny3_val
    total_ck_all += tot_ck
    print(f"{stt:<4} | {str(cn):<15} | {n_str:<24} | {str(cv):<8} | {ck_val:>14,.0f} | {combo_val:>10,.0f} | {ny3_val:>10,.0f} | {tot_ck:>15,.0f} đ")

print("-" * 115)
print(f"--> TỔNG CỘNG QUỸ THƯỞNG CK TOÀN HỆ THỐNG: {total_ck_all:>15,.0f} đ")
