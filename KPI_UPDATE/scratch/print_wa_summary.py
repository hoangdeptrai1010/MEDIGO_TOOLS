import openpyxl, sys
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

def scan_file_for_whatsapp(file_path):
    print(f"=== Đang quét file: {file_path} ===")
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
    except Exception as e:
        print(f"Lỗi đọc file: {e}")
        return

    ws = wb.active
    
    # Tìm vị trí các cột
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    col_map = {}
    for idx, h in enumerate(header_row):
        if h:
            col_map[str(h).strip().lower()] = idx
            
    print(f"Tổng số cột: {len(header_row)}")
    branch_col = col_map.get('chi nhánh', 0)
    bill_col = col_map.get('mã hóa đơn', 1)
    seller_col = col_map.get('người bán', 20 if len(header_row) > 20 else 7)
    need_pay_col = col_map.get('khách cần trả', 43 if len(header_row) > 43 else 12)
    amount_col = col_map.get('thành tiền', 68 if len(header_row) > 68 else 26)
    
    # Tìm tất cả các cột có chữ 'ghi chú'
    note_cols = [idx for idx, h in enumerate(header_row) if h and 'ghi chú' in str(h).strip().lower()]
    print(f"Các cột Ghi chú được tìm thấy: {[(idx, header_row[idx]) for idx in note_cols]}")
    
    keywords = ['whatsapp', 'whats app', 'watsapp']
    
    wa_bills = {}
    all_matched_notes = defaultdict(int)
    row_count = 0
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_count += 1
        matched = False
        matched_note = ''
        
        # Chỉ quét trong các cột ghi chú hoặc toàn bộ cell nếu ít cột
        for col_idx in (note_cols if note_cols else range(len(row))):
            val = row[col_idx]
            if val is not None:
                val_lower = str(val).lower()
                # Yêu cầu: CHỈ CHẤP NHẬN NẾU CÓ CHỮ 'whatsapp'
                if any(k in val_lower for k in keywords):
                    matched = True
                    matched_note = str(val).strip()
                    break
                    
        if matched:
            b_code = row[bill_col]
            if b_code not in wa_bills:
                rev = row[need_pay_col] if need_pay_col < len(row) and row[need_pay_col] is not None else 0.0
                wa_bills[b_code] = {
                    'branch': row[branch_col],
                    'bill': b_code,
                    'seller': row[seller_col],
                    'note': matched_note,
                    'rev': float(rev) if rev else 0.0
                }
            all_matched_notes[matched_note] += 1

    print(f"Đã quét tổng cộng {row_count} dòng.")
    print(f"--> TỔNG SỐ HÓA ĐƠN CÓ CHỮ 'WHATSAPP': {len(wa_bills)} bill")
    
    if not wa_bills:
        print("KHÔNG tìm thấy đơn nào có ghi chú chứa chữ 'whatsapp'!")
        return

    # Thống kê theo chi nhánh
    by_branch = defaultdict(lambda: {'bills': 0, 'rev': 0.0})
    by_seller = defaultdict(lambda: {'bills': 0, 'rev': 0.0, 'branch': ''})

    for b, d in wa_bills.items():
        by_branch[d['branch']]['bills'] += 1
        by_branch[d['branch']]['rev'] += d['rev']
        by_seller[d['seller']]['bills'] += 1
        by_seller[d['seller']]['rev'] += d['rev']
        by_seller[d['seller']]['branch'] = d['branch']

    print("\n" + "="*50)
    print("=== THỐNG KÊ THEO CHI NHÁNH ===")
    for br, data in sorted(by_branch.items()):
        print(f"{br}: {data['bills']} đơn | Doanh thu: {data['rev']:,.0f} đ")

    print("\n" + "="*50)
    print("=== THỐNG KÊ THEO DƯỢC SĨ ===")
    for se, data in sorted(by_seller.items(), key=lambda x: x[1]['rev'], reverse=True):
        print(f"{se} ({data['branch']}): {data['bills']} đơn | Doanh thu: {data['rev']:,.0f} đ")

    print("\n" + "="*50)
    print(f"=== CÁC DẠNG GHI CHÚ WHATSAPP THỰC TẾ (Tổng {len(all_matched_notes)} kiểu) ===")
    for note, count in sorted(all_matched_notes.items(), key=lambda x: x[1], reverse=True)[:25]:
        print(f"{count:3d} lần : {repr(note)}")

if __name__ == '__main__':
    # Quét file xuất chuẩn 72 cột
    f_72col = r'C:\Users\10102\Downloads\DanhSachChiTietHoaDon_KV03092026-140832-762.xlsx'
    scan_file_for_whatsapp(f_72col)
