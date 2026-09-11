import openpyxl, sys, io
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

hd_path = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb = openpyxl.load_workbook(hd_path, read_only=True)
ws = wb.active

rows_iter = ws.iter_rows(values_only=True)
header = next(rows_iter)
print("=== HEADERS ===")
for i, h in enumerate(header):
    print(f"Col {i:2d}: {h}")

whatsapp_invoices = defaultdict(list)
all_notes_seen = set()

for r_idx, row in enumerate(rows_iter, start=2):
    # Search all text columns for 'whatsapp'
    row_text = " ".join([str(c or '') for c in row]).lower()
    if 'whatsapp' in row_text or 'whats app' in row_text or 'what app' in row_text or 'whatapp' in row_text:
        ma_hd = row[1]
        cn = row[0]
        time = row[2]
        seller = row[7]
        kenh = row[8]
        can_tra = row[12]
        sku = row[14]
        ten_hang = row[16]
        sl = row[21]
        dg = row[22]
        tt = row[26]
        ghi_chu = row[20]
        
        whatsapp_invoices[ma_hd].append({
            'row_idx': r_idx,
            'branch': cn,
            'ma_hd': ma_hd,
            'time': time,
            'seller': seller,
            'channel': kenh,
            'can_tra': can_tra,
            'sku': sku,
            'ten_hang': ten_hang,
            'sl': sl,
            'don_gia': dg,
            'thanh_tien': tt,
            'ghi_chu': ghi_chu,
            'full_row': row
        })

wb.close()

print(f"\n=== TÌM THẤY {len(whatsapp_invoices)} HÓA ĐƠN WHATSAPP ===")

seller_totals = defaultdict(float)
branch_totals = defaultdict(float)

for ma_hd, items in whatsapp_invoices.items():
    first = items[0]
    # Sum distinct invoice revenue
    inv_rev = float(first['can_tra'] or 0)
    seller_totals[first['seller']] += inv_rev
    branch_totals[first['branch']] += inv_rev

print("\n--- TỔNG DOANH THU THEO CHI NHÁNH ---")
for b, val in sorted(branch_totals.items(), key=lambda x: x[1], reverse=True):
    print(f"  {b:25s}: {val:>15,.0f} đ")

print("\n--- TỔNG DOANH THU THEO DƯỢC SĨ (NGƯỜI BÁN) ---")
for s, val in sorted(seller_totals.items(), key=lambda x: x[1], reverse=True):
    print(f"  {s:25s}: {val:>15,.0f} đ")

print("\n--- MẪU 5 ĐƠN HÀNG ĐẦU TIÊN ---")
for idx, (ma_hd, items) in enumerate(list(whatsapp_invoices.items())[:5], 1):
    first = items[0]
    print(f"{idx}. Mã HĐ: {ma_hd} | Chi nhánh: {first['branch']} | DS: {first['seller']} | Thời gian: {first['time']} | Tiền: {first['can_tra']:,.0f} đ")
    for it in items:
        print(f"     - SP: {it['sku']} - {it['ten_hang']} (SL: {it['sl']}, TT: {it['thanh_tien']:,.0f} đ) | Ghi chú: {it['ghi_chu']}")
