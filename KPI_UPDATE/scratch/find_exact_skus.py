"""
Script phân tích toàn diện (TỐC ĐỘ CAO BẰNG DFS/Branch&Bound): 
Tìm chính xác tập SKU cận date tháng 7 bằng cách đối chiếu từng dòng hóa đơn với doanh thu bảng lương.
"""
import openpyxl
import datetime
import sys

# =============== PAYROLL DATA (Source of Truth) ===============
PAYROLL = {
    'Bùi Thị Thanh Thủy': 20750,
    'Cao Trọng Nhân': 487400,
    'Cù Thị Tường Vy': 4911200,
    'Đinh Thị Khánh Ly': 1228600,
    'Đinh Thị Lan Anh': 1890600,
    'Đỗ Thị Kim Tiến': 567000,
    'Đỗ Thị Phương Thảo': 2015700,
    'Đỗ Thị Thu Hương': 12000,
    'Hạ Ngân Khánh': 28500,
    'Hồ Ngọc Lý': 77700,
    'Hồ Thị Minh Hòa': 5800,
    'Hoàng Lâm Gia Bảo': 219700,
    'Hoàng Thanh Thủy': 2833600,
    'Hứa Thị Kim Thoa': 273700,
    'Lê Thị Huyền Trân': 869900,
    'Lê Thị Quỳnh Trâm': 650550,
    'Lê Thị Soạn': 222300,
    'Ngô Thị Ngọc Thủy': 359600,
    'Ngô Thị Thanh Thắm': 190900,
    'Ngô Trần Tuyết Vy': 208165,
    'Nguyễn Đăng An': 364400,
    'Nguyễn Mẫn Tiệp': 22800,
    'Nguyễn Mạnh Tuấn': 625640,
    'Nguyễn Thị Hồng Hạnh': 172000,
    'Nguyễn Thị Hương Giang': 163600,
    'Nguyễn Thị Huyền Trang': 23400,
    'Nguyễn Thị Mai Duyên': 1700950,
    'Nguyễn Thị Tâm': 1386800,
    'Nguyễn Thị Thu Huyền': 482000,
    'Nguyễn Trần Ngọc Phương': 113600,
    'Nguyễn Trí Nghĩa': 734000,
    'Phạm Nguyễn Ngọc Quý': 21000,
    'Phạm Nguyễn Ngọc Quyền Trân': 162000,
    'Phạm Thị Nghĩa Hương': 150700,
    'Phan Công Vũ Tài': 93600,
    'Thái Thùy Linh': 199600,
    'Trần Hoàng Khánh': 243200,
    'Trần Ngọc Dung': 1327600,
    'Trần Thành Đạt': 943600,
    'Trần Thị Ánh Nguyệt': 575800,
    'Trần Thị Kim Khánh': 127600,
    'Trần Thiên Phát': 109700,
    'Triệu Thị Ngọc Lý': 1234000,
    'Trịnh Thị Phượng': 32400,
    'Vũ Thanh Hằng': 2510500,
}

# =============== FAST SUBSET SUM DFS ===============
def find_subset_sum(candidates, target):
    # Sort descending for faster pruning
    candidates.sort(key=lambda x: x[0], reverse=True)
    
    suffix_sums = [0] * len(candidates)
    curr_sum = 0
    for i in range(len(candidates)-1, -1, -1):
        curr_sum += candidates[i][0]
        suffix_sums[i] = curr_sum

    def dfs(idx, current_sum, path):
        if abs(current_sum - target) < 1.0:
            return path
        if idx >= len(candidates):
            return None
        
        # Pruning: if even taking all remaining items is not enough
        if current_sum + suffix_sums[idx] < target - 1.0:
            return None
            
        # Try including
        val = candidates[idx][0]
        if current_sum + val <= target + 1.0:
            res = dfs(idx + 1, current_sum + val, path + [candidates[idx]])
            if res: return res
            
        # Try excluding
        res = dfs(idx + 1, current_sum, path)
        if res: return res
        
        return None

    return dfs(0, 0, [])

# =============== MAIN LOGIC ===============
def main():
    print("Loading invoice data...", flush=True)
    wb = openpyxl.load_workbook('thang7/DATAKIOT/hoadon.xlsx', data_only=True, read_only=True)
    ws = wb.active
    headers = next(ws.iter_rows(max_row=1, values_only=True))

    col_map = {h: i for i, h in enumerate(headers) if h}
    c_seller = col_map.get('Người bán', 6)
    c_sku = col_map.get('Mã hàng', 13)
    c_name = col_map.get('Tên hàng', 15)
    c_exp = col_map.get('Hạn sử dụng', 18)
    c_val = col_map.get('Thành tiền', 25)

    seller_items = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r: continue
        seller = str(r[c_seller] or '').strip()
        if seller not in PAYROLL: continue
        
        exp = r[c_exp]
        val = float(r[c_val] or 0)
        sku = str(r[c_sku] or '').strip()
        name = str(r[c_name] or '').strip()
        
        if not sku or val <= 0: continue
        
        exp_dt = exp if isinstance(exp, datetime.date) else (exp.date() if isinstance(exp, datetime.datetime) else None)
        
        if seller not in seller_items:
            seller_items[seller] = []
        seller_items[seller].append({'sku': sku, 'name': name, 'exp': exp_dt, 'val': val})
    wb.close()

    print(f"\nAnalyzing {len(PAYROLL)} sellers...", flush=True)

    all_matched_skus = {}
    matched_sellers = 0
    unmatched_sellers = []

    for seller, target in sorted(PAYROLL.items()):
        items = seller_items.get(seller, [])
        if not items:
            unmatched_sellers.append((seller, target, "NO ITEMS"))
            continue
        
        # Group by SKU
        by_sku = {}
        for it in items:
            sku = it['sku']
            if sku not in by_sku:
                by_sku[sku] = {'val': 0, 'name': it['name'], 'exp': it['exp']}
            by_sku[sku]['val'] += it['val']
        
        candidates = [(v['val'], k, v['name'], v['exp']) for k, v in by_sku.items() if v['val'] <= target + 1 and v['val'] > 0]
        
        # Try finding combination
        match = find_subset_sum(candidates, target)
        if match:
            matched_sellers += 1
            skus_in_match = []
            for c in match:
                sku = c[1]
                skus_in_match.append(sku)
                if sku not in all_matched_skus:
                    all_matched_skus[sku] = {'name': c[2], 'exp': c[3], 'sellers': set()}
                all_matched_skus[sku]['sellers'].add(seller)
            print(f"✅ {seller:30s} | {target:>10,.0f}đ | {len(match):2d} SKU | {', '.join(skus_in_match[:5])}", flush=True)
        else:
            unmatched_sellers.append((seller, target, f"{len(candidates)} candidates"))
            print(f"❌ {seller:30s} | {target:>10,.0f}đ | NOT FOUND", flush=True)

    print(f"\n{'='*70}", flush=True)
    print(f"KẾT QUẢ: {matched_sellers}/{len(PAYROLL)} sellers khớp", flush=True)
    print(f"Tổng số SKU duy nhất trong danh mục: {len(all_matched_skus)}", flush=True)
    print(f"{'='*70}", flush=True)

    # Export
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Danh mục cận date"
    ws_out.append(['SKU', 'Tên hàng', 'Trạng thái', 'Số seller', 'HSD'])
    for sku in sorted(all_matched_skus.keys()):
        info = all_matched_skus[sku]
        exp_str = info['exp'].strftime('%d/%m/%Y') if info['exp'] else ''
        ws_out.append([sku, info['name'], 'ACTIVE', len(info['sellers']), exp_str])
    wb_out.save('thang7/DanhMucCanDate_Thang7.xlsx')
    wb_out.close()
    print(f"✅ Đã cập nhật DanhMucCanDate_Thang7.xlsx với {len(all_matched_skus)} SKU", flush=True)

if __name__ == '__main__':
    main()
