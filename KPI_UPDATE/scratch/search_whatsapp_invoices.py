import openpyxl, sys, os
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

def scan_whatsapp_orders(file_path):
    print("=" * 65)
    print(f"BẮT ĐẦU QUÉT ĐƠN WHATSAPP TRONG FILE:\n{file_path}")
    print("=" * 65)
    
    if not os.path.exists(file_path):
        print(f"Lỗi: Không tìm thấy file: {file_path}")
        return

    wb = openpyxl.load_workbook(file_path, read_only=True)
    ws = wb.active
    
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    
    print(f"File có tổng cộng {len(hdr)} cột.")
    
    # Xác định các cột quan trọng
    col_map = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}
    
    branch_idx = col_map.get('chi nhánh', 0)
    bill_idx = col_map.get('mã hóa đơn', 1)
    seller_idx = col_map.get('người bán', 20 if len(hdr) > 20 else 7)
    time_idx = col_map.get('thời gian', 6 if len(hdr) > 6 else 2)
    need_pay_idx = col_map.get('khách cần trả', 43 if len(hdr) > 43 else 12)
    
    # Tìm tất cả các cột ghi chú
    note_indices = [i for i, h in enumerate(hdr) if h and 'ghi chú' in str(h).strip().lower()]
    print("Các cột Ghi chú được quét:")
    for i in note_indices:
        print(f"  - Cột {i}: {hdr[i]}")
        
    if not note_indices:
        print("Cảnh báo: Không tìm thấy cột ghi chú nào, sẽ quét toàn bộ cell.")
        note_indices = range(len(hdr))
        

    keywords = ['whatsapp', 'whats app', 'watsapp']
    
    wa_bills = {}
    matched_notes_stat = defaultdict(int)
    row_count = 0
    
    for r in it:
        row_count += 1
        matched = False
        found_note = ''
        
        for idx in note_indices:
            if idx < len(r) and r[idx] is not None:
                val_str = str(r[idx]).lower()
                # Kiểm tra có chứa chữ whatsapp
                if any(k in val_str for k in keywords):
                    matched = True
                    found_note = str(r[idx]).strip()
                    break
                    
        if matched:
            bill_code = r[bill_idx]
            if bill_code not in wa_bills:
                rev = r[need_pay_idx] if need_pay_idx < len(r) and r[need_pay_idx] is not None else 0.0
                wa_bills[bill_code] = {
                    'branch': r[branch_idx],
                    'bill': bill_code,
                    'time': r[time_idx],
                    'seller': r[seller_idx],
                    'note': found_note,
                    'rev': float(rev) if rev else 0.0
                }
            matched_notes_stat[found_note] += 1

    print(f"\nĐã quét qua {row_count:,} dòng dữ liệu.")
    print("=" * 65)
    print(f"TỔNG SỐ ĐƠN HÀNG CÓ GHI CHÚ WHATSAPP: {len(wa_bills)} ĐƠN")
    print("=" * 65)

    if not wa_bills:
        print("--> Không tìm thấy đơn nào có ghi chú 'whatsapp'.")
        if len(hdr) <= 30:
            print("LƯU Ý: File này chỉ có 30 cột (thiếu cột 'Ghi chú' hóa đơn). Cần xuất mẫu đầy đủ từ KiotViet.")
        return

    # Thống kê theo Chi nhánh
    branch_stats = defaultdict(lambda: {'bills': 0, 'rev': 0.0})
    seller_stats = defaultdict(lambda: {'bills': 0, 'rev': 0.0, 'branch': ''})

    for b, d in wa_bills.items():
        branch_stats[d['branch']]['bills'] += 1
        branch_stats[d['branch']]['rev'] += d['rev']
        seller_stats[d['seller']]['bills'] += 1
        seller_stats[d['seller']]['rev'] += d['rev']
        seller_stats[d['seller']]['branch'] = d['branch']

    print("\n--- 1. TỔNG HỢP THEO CHI NHÁNH ---")
    for br, st in sorted(branch_stats.items()):
        # Xác định bậc KPI theo quy định docx
        rev = st['rev']
        if rev >= 280_000_000:
            rate = 0.06
            tier = "Mức 3 (>=280M: 6%)"
        elif rev >= 200_000_000:
            rate = 0.04
            tier = "Mức 2 (>=200M: 4%)"
        elif rev >= 150_000_000:
            rate = 0.03
            tier = "Mức 1 (>=150M: 3%)"
        else:
            rate = 0.015
            tier = "Không đạt target (<150M: 1.5%)"
            
        print(f"• {br}: {st['bills']} đơn | Doanh thu: {rev:,.0f} đ -> Đạt {tier}")

    print("\n--- 2. TỔNG HỢP THEO TỪNG DƯỢC SĨ ---")
    for se, st in sorted(seller_stats.items(), key=lambda x: x[1]['rev'], reverse=True):
        print(f"• {se} ({st['branch']}): {st['bills']} đơn | Doanh thu: {st['rev']:,.0f} đ")

    print("\n--- 3. MỘT SỐ MẪU GHI CHÚ WHATSAPP ĐƯỢC GHI NHẬN ---")
    for note, count in sorted(matched_notes_stat.items(), key=lambda x: x[1], reverse=True)[:15]:
        print(f"  [{count:2d} lần] : {repr(note)}")

if __name__ == '__main__':
    # File KiotViet đầy đủ 72 cột
    path_kv = r'C:\Users\10102\Downloads\DanhSachChiTietHoaDon_KV03092026-140832-762.xlsx'
    scan_whatsapp_orders(path_kv)
