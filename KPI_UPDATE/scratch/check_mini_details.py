import pandas as pd
import openpyxl

# Let's inspect the invoices of Phan Cong Vu Tai, Trinh Thi Phuong, Cu Thi Tuong Vy for KAT and Party Smart
file_hd = 'thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb_hd = openpyxl.load_workbook(file_hd, data_only=True)
ws_hd = wb_hd.active

headers = [str(ws_hd.cell(1, c).value or '').strip() for c in range(1, ws_hd.max_column+1)]
print('Invoice headers:', headers[:15])

# Find col indices
col_seller = headers.index('Người bán') + 1 if 'Người bán' in headers else 35
col_sku = headers.index('Mã hàng') + 1 if 'Mã hàng' in headers else 4
col_name = headers.index('Tên hàng') + 1 if 'Tên hàng' in headers else 5
col_qty = headers.index('Số lượng') + 1 if 'Số lượng' in headers else 7
col_branch = headers.index('Chi nhánh') + 1 if 'Chi nhánh' in headers else 36

kat_skus = ['SP2725289', 'SP2725285', 'SP2725291', 'SP2725287', 'SP2725353', 'SP2725351']
ps_skus = ['SP017162']

print(f"Col Seller={col_seller}, Col SKU={col_sku}, Col Qty={col_qty}, Col Branch={col_branch}")

# Summarize KAT sales by seller
kat_by_seller = {}
kat_by_branch = {}
ps_by_seller = {}
ps_by_branch = {}

for r in range(2, ws_hd.max_row+1):
    sku = str(ws_hd.cell(r, col_sku).value or '').strip()
    seller = str(ws_hd.cell(r, col_seller).value or '').strip()
    branch = str(ws_hd.cell(r, col_branch).value or '').strip()
    qty = ws_hd.cell(r, col_qty).value or 0
    if not isinstance(qty, (int, float)):
        continue
        
    if sku in kat_skus:
        if seller not in kat_by_seller: kat_by_seller[seller] = {'total_qty': 0, 'skus': set()}
        kat_by_seller[seller]['total_qty'] += qty
        kat_by_seller[seller]['skus'].add(sku)
        
        if branch not in kat_by_branch: kat_by_branch[branch] = {'total_qty': 0, 'skus': set()}
        kat_by_branch[branch]['total_qty'] += qty
        kat_by_branch[branch]['skus'].add(sku)

    if sku in ps_skus:
        ps_by_seller[seller] = ps_by_seller.get(seller, 0) + qty
        ps_by_branch[branch] = ps_by_branch.get(branch, 0) + qty

print("\n=== TOP DƯỢC SĨ BÁN KAT ===")
for s, d in sorted(kat_by_seller.items(), key=lambda x: x[1]['total_qty'], reverse=True)[:10]:
    print(f"  {s:25s} | SL={d['total_qty']:>3.0f} hộp | Số mã khác nhau={len(d['skus'])}/6 mã: {list(d['skus'])}")

print("\n=== TOP NHÀ THUỐC BÁN KAT (ĐỘ PHỦ 3/6 MÃ) ===")
for b, d in sorted(kat_by_branch.items(), key=lambda x: x[1]['total_qty'], reverse=True):
    qualify = "ĐẠT ĐỘ PHỦ (>=3 mã)" if len(d['skus']) >= 3 else "KHÔNG ĐẠT ĐỘ PHỦ (<3 mã)"
    print(f"  {b:25s} | Tổng SL={d['total_qty']:>3.0f} hộp | Số mã={len(d['skus'])}/6 mã | {qualify}")

print("\n=== TOP DƯỢC SĨ BÁN PARTY SMART ===")
for s, qty in sorted(ps_by_seller.items(), key=lambda x: x[1], reverse=True)[:10]:
    print(f"  {s:25s} | SL={qty:>3.0f} hộp")

print("\n=== TOP NHÀ THUỐC BÁN PARTY SMART ===")
for b, qty in sorted(ps_by_branch.items(), key=lambda x: x[1], reverse=True)[:10]:
    print(f"  {b:25s} | SL={qty:>3.0f} hộp")
