"""
MODULE 2: TÍNH THƯỞNG DỰ ÁN MINIKAT (HÀ NỘI & HỒ CHÍ MINH)
- Quét hóa đơn & trừ trả hàng các nhóm SKUs:
  + Party Smart (SP017162, SP2723878, SP2723879)
  + KAT (SP2725289, SP2725285, SP2725291, SP2725287, SP2725353, SP2725351)
  + LadyCare (SP2723124, SP2723125, SP2723127)
- Tính thưởng cho từng dược sĩ (10,000 đ - 20,000 đ/sản phẩm)
- Tính thưởng cho đại diện CHT theo chi nhánh
- Đổ dữ liệu và gắn công thức vào 2 sheet: 'MiniKat - HN' và 'MiniKat - HCM'
"""

import openpyxl
from collections import defaultdict
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

SKUS_PARTY_SMART = {'SP017162', 'SP2723878', 'SP2723879'}
SKUS_KAT = {'SP2725289', 'SP2725285', 'SP2725291', 'SP2725287', 'SP2725353', 'SP2725351'}
SKUS_LADYCARE = {'SP2723124', 'SP2723125', 'SP2723127'}
HN_BRANCHES = {'Hàng Bông', 'Đường Láng'}

def to_num(v, default=0.0):
    if v is None: return default
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(',', '')
    try: return float(s)
    except: return default

def scan_minikat_sales(inv_file, ret_file=None):
    minikat_counts = defaultdict(lambda: {'ps': 0, 'kat': 0, 'lc': 0})
    if not inv_file:
        return minikat_counts

    # 1. Đọc hóa đơn bán
    try:
        wb_inv = openpyxl.load_workbook(inv_file, data_only=True, read_only=True)
        ws_inv = wb_inv.active
        hdr = next(ws_inv.iter_rows(min_row=1, max_row=1, values_only=True))
        col_map = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h}

        seller_idx = col_map.get('người bán', 20 if len(hdr) > 20 else 7)
        sku_idx = col_map.get('mã hàng', 1)
        qty_idx = col_map.get('số lượng', 13 if len(hdr) > 13 else 4)

        for r in ws_inv.iter_rows(min_row=2, values_only=True):
            seller = str(r[seller_idx]).strip() if seller_idx < len(r) and r[seller_idx] else ''
            sku = str(r[sku_idx]).strip() if sku_idx < len(r) and r[sku_idx] else ''
            qty = to_num(r[qty_idx]) if qty_idx < len(r) else 0.0

            if seller and sku and qty > 0:
                if sku in SKUS_PARTY_SMART:
                    minikat_counts[seller]['ps'] += qty
                elif sku in SKUS_KAT:
                    minikat_counts[seller]['kat'] += qty
                elif sku in SKUS_LADYCARE:
                    minikat_counts[seller]['lc'] += qty
        wb_inv.close()
    except Exception as e:
        print(f"Warning scanning MiniKat invoices: {e}")

    # 2. Trừ trả hàng
    if ret_file:
        try:
            wb_ret = openpyxl.load_workbook(ret_file, data_only=True, read_only=True)
            ws_ret = wb_ret.active
            hdr_r = next(ws_ret.iter_rows(min_row=1, max_row=1, values_only=True))
            col_map_r = {str(h).strip().lower(): i for i, h in enumerate(hdr_r) if h}

            seller_idx_r = col_map_r.get('người nhận trả', col_map_r.get('người bán', 10))
            sku_idx_r = col_map_r.get('mã hàng', 1)
            qty_idx_r = col_map_r.get('số lượng', 5)

            for r in ws_ret.iter_rows(min_row=2, values_only=True):
                seller = str(r[seller_idx_r]).strip() if seller_idx_r < len(r) and r[seller_idx_r] else ''
                sku = str(r[sku_idx_r]).strip() if sku_idx_r < len(r) and r[sku_idx_r] else ''
                qty = to_num(r[qty_idx_r]) if qty_idx_r < len(r) else 0.0

                if seller and sku and qty > 0:
                    if sku in SKUS_PARTY_SMART:
                        minikat_counts[seller]['ps'] = max(0, minikat_counts[seller]['ps'] - qty)
                    elif sku in SKUS_KAT:
                        minikat_counts[seller]['kat'] = max(0, minikat_counts[seller]['kat'] - qty)
                    elif sku in SKUS_LADYCARE:
                        minikat_counts[seller]['lc'] = max(0, minikat_counts[seller]['lc'] - qty)
            wb_ret.close()
        except Exception as e:
            print(f"Warning scanning MiniKat returns: {e}")

    return minikat_counts

def process_minikat_sheets(wb_out, inv_file=None, ret_file=None):
    mk_data = scan_minikat_sales(inv_file, ret_file)
    print(f"--> [Module MiniKAT] Đã tính toán doanh số MiniKAT cho {len(mk_data)} nhân sự.")

    # 1. Đổ dữ liệu vào sheet 'MiniKat - HN'
    if 'MiniKat - HN' in wb_out.sheetnames:
        ws_mkhn = wb_out['MiniKat - HN']
        for r in range(2, ws_mkhn.max_row + 1):
            nm = ws_mkhn.cell(r, 1).value
            if nm and str(nm).strip() and str(nm).strip() != 'Tổng':
                name_str = str(nm).strip()
                if name_str in mk_data:
                    d = mk_data[name_str]
                    ws_mkhn.cell(r, 4, d['ps']).number_format = '#,##0'
                    ws_mkhn.cell(r, 7, d['kat']).number_format = '#,##0'
                    ws_mkhn.cell(r, 10, d['lc']).number_format = '#,##0'
                ws_mkhn.cell(r, 6, f"=D{r}*E{r}").number_format = '#,##0'
                ws_mkhn.cell(r, 9, f"=G{r}*H{r}").number_format = '#,##0'
                ws_mkhn.cell(r, 11, f"=I{r}*J{r}").number_format = '#,##0'
                ws_mkhn.cell(r, 12, f"=SUM(F{r}, I{r}, K{r})").number_format = '#,##0'

    # 2. Đổ dữ liệu vào sheet 'MiniKat - HCM'
    if 'MiniKat - HCM' in wb_out.sheetnames:
        ws_mkhcm = wb_out['MiniKat - HCM']
        for r in range(2, ws_mkhcm.max_row + 1):
            nm = ws_mkhcm.cell(r, 1).value
            if nm and str(nm).strip() and str(nm).strip() != 'Tổng':
                name_str = str(nm).strip()
                if name_str in mk_data:
                    d = mk_data[name_str]
                    ws_mkhcm.cell(r, 4, d['ps']).number_format = '#,##0'
                    ws_mkhcm.cell(r, 8, d['lc']).number_format = '#,##0'
                ws_mkhcm.cell(r, 6, f"=D{r}*E{r}").number_format = '#,##0'
                ws_mkhcm.cell(r, 9, f"=G{r}*H{r}").number_format = '#,##0'
                ws_mkhcm.cell(r, 10, f"=SUM(F{r}, I{r})").number_format = '#,##0'

    print("✅ Đã cập nhật xong dữ liệu MiniKAT Hà Nội & Hồ Chí Minh!")
