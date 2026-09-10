import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import sys
import datetime

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if not BASE_DIR:
    BASE_DIR = '.'

# 1. Load SKU dictionaries
# CK-HN
wb_hn = openpyxl.load_workbook(os.path.join(BASE_DIR, 'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/CK-HN.xlsx'), data_only=True)
ws_ck_hn = wb_hn['danhmucchietkhau-duan']
ck_hn_skus = {}
for r in range(2, ws_ck_hn.max_row + 1):
    sku = str(ws_ck_hn.cell(r, 2).value or '').strip()
    name = str(ws_ck_hn.cell(r, 3).value or '').strip()
    dvt = str(ws_ck_hn.cell(r, 4).value or '').strip()
    price = ws_ck_hn.cell(r, 5).value
    bonus = ws_ck_hn.cell(r, 6).value
    if sku:
        ck_hn_skus[sku] = {'name': name, 'dvt': dvt, 'price': price, 'bonus': bonus}

# CK-HCM
wb_hcm = openpyxl.load_workbook(os.path.join(BASE_DIR, 'thang8/Pharmacy_retail_Store_KPIs_August 2026/Hồ Chí Minh/Dự án/Chương trình/CK-HCM.xlsx'), data_only=True)
ws_ck_hcm = wb_hcm['danhmucck-duan']
ck_hcm_skus = {}
for r in range(2, ws_ck_hcm.max_row + 1):
    sku = str(ws_ck_hcm.cell(r, 2).value or '').strip()
    name = str(ws_ck_hcm.cell(r, 3).value or '').strip()
    dvt = str(ws_ck_hcm.cell(r, 4).value or '').strip()
    price = ws_ck_hcm.cell(r, 5).value
    bonus = ws_ck_hcm.cell(r, 6).value
    if sku:
        ck_hcm_skus[sku] = {'name': name, 'dvt': dvt, 'price': price, 'bonus': bonus}

# Plan general CK
wb_plan = openpyxl.load_workbook(os.path.join(BASE_DIR, 'plans/KeHoachKPI_2026-08.xlsx'), data_only=True)
ws_dm = wb_plan['Danh mục dự án']
plan_ck_skus = {}
for r in range(2, ws_dm.max_row + 1):
    sku = str(ws_dm.cell(r, 1).value or '').strip()
    grp = str(ws_dm.cell(r, 2).value or '').strip()
    name = str(ws_dm.cell(r, 3).value or '').strip()
    if sku and grp == 'CK':
        plan_ck_skus[sku] = {'name': name, 'grp': grp}

all_target_skus = set(ck_hn_skus.keys()) | set(ck_hcm_skus.keys()) | set(plan_ck_skus.keys())
print(f"Tổng số SKU CK cần trích xuất: {len(all_target_skus)}")

# 2. Read Invoices
hd_path = os.path.join(BASE_DIR, 'thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx')
wb_hd = openpyxl.load_workbook(hd_path, read_only=True)
ws_hd = wb_hd.active

rows_extracted = []
product_summary = {} # sku -> dict
staff_summary = {}   # (branch, seller) -> dict
branch_summary = {}  # branch -> dict

for r in ws_hd.iter_rows(values_only=True):
    sku = str(r[14] or '').strip()
    if sku in all_target_skus:
        cn = str(r[0] or '').strip()
        ma_hd = str(r[1] or '').strip()
        dt_val = r[2]
        if isinstance(dt_val, datetime.datetime):
            dt_str = dt_val.strftime('%d/%m/%Y %H:%M:%S')
        elif isinstance(dt_val, str):
            dt_str = dt_val[:19]
        else:
            dt_str = str(dt_val or '')

        kh = str(r[5] or '').strip()
        seller = str(r[7] or '').strip()
        channel = str(r[8] or '').strip()
        tong_tien_hd = float(r[9] or 0)
        giam_gia_hd = float(r[10] or 0)
        
        item_name = str(r[16] or '').strip()
        dvt = str(r[17] or '').strip()
        qty = float(r[21] or 0)
        don_gia = float(r[22] or 0)
        giam_gia_dong = float(r[24] or 0)
        gia_ban = float(r[25] or 0)
        thanh_tien = float(r[26] or 0)

        # Net revenue after proportional invoice discount allocation
        alloc_disc = (giam_gia_hd * thanh_tien / tong_tien_hd) if tong_tien_hd > 0 else 0.0
        net_revenue = thanh_tien - alloc_disc

        # Determine classification & bonus rate
        is_hn = sku in ck_hn_skus
        is_hcm = sku in ck_hcm_skus
        
        if is_hn and is_hcm:
            cat_label = "CK-HN & CK-HCM"
            bonus_rate = ck_hn_skus[sku]['bonus'] or ck_hcm_skus[sku]['bonus'] or 0
        elif is_hn:
            cat_label = "CK-HN"
            bonus_rate = ck_hn_skus[sku]['bonus'] or 0
        elif is_hcm:
            cat_label = "CK-HCM"
            bonus_rate = ck_hcm_skus[sku]['bonus'] or 0
        else:
            cat_label = "CK chung"
            bonus_rate = 0

        bonus_total = qty * float(bonus_rate or 0)

        row_dict = {
            'ma_hd': ma_hd,
            'thoi_gian': dt_str,
            'chi_nhanh': cn,
            'nguoi_ban': seller,
            'kenh_ban': channel,
            'ma_kh': kh,
            'ma_hang': sku,
            'ten_hang': item_name,
            'dvt': dvt,
            'so_luong': qty,
            'don_gia': don_gia,
            'giam_gia_dong': giam_gia_dong,
            'thanh_tien': thanh_tien,
            'doanh_thu_net': net_revenue,
            'phan_loai': cat_label,
            'muc_thuong_dv': float(bonus_rate or 0),
            'tong_tien_thuong': bonus_total
        }
        rows_extracted.append(row_dict)

        # Aggregate by Product
        if sku not in product_summary:
            product_summary[sku] = {
                'sku': sku,
                'name': item_name,
                'dvt': dvt,
                'cat': cat_label,
                'bonus_rate': float(bonus_rate or 0),
                'total_qty': 0.0,
                'total_rev': 0.0,
                'total_net_rev': 0.0,
                'total_bonus': 0.0,
                'bill_count': 0
            }
        product_summary[sku]['total_qty'] += qty
        product_summary[sku]['total_rev'] += thanh_tien
        product_summary[sku]['total_net_rev'] += net_revenue
        product_summary[sku]['total_bonus'] += bonus_total
        product_summary[sku]['bill_count'] += 1

        # Aggregate by Staff
        staff_key = (cn, seller)
        if staff_key not in staff_summary:
            staff_summary[staff_key] = {
                'branch': cn,
                'seller': seller,
                'total_qty': 0.0,
                'total_rev': 0.0,
                'total_net_rev': 0.0,
                'total_bonus': 0.0,
                'bill_count': 0
            }
        staff_summary[staff_key]['total_qty'] += qty
        staff_summary[staff_key]['total_rev'] += thanh_tien
        staff_summary[staff_key]['total_net_rev'] += net_revenue
        staff_summary[staff_key]['total_bonus'] += bonus_total
        staff_summary[staff_key]['bill_count'] += 1

        # Aggregate by Branch
        if cn not in branch_summary:
            branch_summary[cn] = {
                'branch': cn,
                'total_qty': 0.0,
                'total_rev': 0.0,
                'total_net_rev': 0.0,
                'total_bonus': 0.0,
                'bill_count': 0
            }
        branch_summary[cn]['total_qty'] += qty
        branch_summary[cn]['total_rev'] += thanh_tien
        branch_summary[cn]['total_net_rev'] += net_revenue
        branch_summary[cn]['total_bonus'] += bonus_total
        branch_summary[cn]['bill_count'] += 1

print(f"Tổng số dòng hóa đơn CK đã trích xuất: {len(rows_extracted):,}")

# 3. Create Excel Workbook
wb_out = openpyxl.Workbook()

# Setup Styles
font_title = Font(name='Calibri', size=14, bold=True, color='1E3A8A')
font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
font_bold = Font(name='Calibri', size=11, bold=True)
font_regular = Font(name='Calibri', size=11)
font_small = Font(name='Calibri', size=10, italic=True, color='64748B')

fill_header = PatternFill(start_color='1E40AF', end_color='1E40AF', fill_type='solid')
fill_header_green = PatternFill(start_color='047857', end_color='047857', fill_type='solid')
fill_header_purple = PatternFill(start_color='6D28D9', end_color='6D28D9', fill_type='solid')
fill_total = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
fill_zebra = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

border_thin = Side(border_style='thin', color='CBD5E1')
border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
border_total = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='double', color='000000'))

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

# --- SHEET 1: Chi tiết Hóa đơn CK ---
ws1 = wb_out.active
ws1.title = "Chi tiết Hóa đơn CK"
ws1.views.sheetView[0].showGridLines = True

ws1.cell(1, 1, "DANH SÁCH CHI TIẾT HÓA ĐƠN BÁN CÁC SẢN PHẨM CHIẾT KHẤU (CK / CK-HN / CK-HCM) - THÁNG 8/2026").font = font_title
ws1.cell(2, 1, f"Thời gian xuất: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | Tổng số dòng hóa đơn: {len(rows_extracted):,} | Tổng số lượng: {sum(r['so_luong'] for r in rows_extracted):,.1f} | Doanh thu: {sum(r['thanh_tien'] for r in rows_extracted):,.0f} VNĐ").font = font_small

headers1 = [
    "STT", "Mã Hóa Đơn", "Thời Gian", "Chi Nhánh", "Người Bán", "Kênh Bán", "Mã KH",
    "Mã Hàng (SKU)", "Tên Sản Phẩm", "ĐVT", "Số Lượng", "Đơn Giá (VNĐ)", "Giảm Giá Dòng (VNĐ)",
    "Thành Tiền (VNĐ)", "Doanh Thu Net (VNĐ)", "Phân Loại", "Thưởng Đơn Vị (VNĐ)", "Tổng Thưởng CK (VNĐ)"
]

ws1.row_dimensions[4].height = 28
for c_idx, h_text in enumerate(headers1, start=1):
    cell = ws1.cell(4, c_idx, h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = border_box

for r_idx, r_data in enumerate(rows_extracted, start=5):
    row_vals = [
        r_idx - 4,
        r_data['ma_hd'],
        r_data['thoi_gian'],
        r_data['chi_nhanh'],
        r_data['nguoi_ban'],
        r_data['kenh_ban'],
        r_data['ma_kh'],
        r_data['ma_hang'],
        r_data['ten_hang'],
        r_data['dvt'],
        r_data['so_luong'],
        r_data['don_gia'],
        r_data['giam_gia_dong'],
        r_data['thanh_tien'],
        r_data['doanh_thu_net'],
        r_data['phan_loai'],
        r_data['muc_thuong_dv'],
        r_data['tong_tien_thuong']
    ]
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws1.cell(r_idx, c_idx, val)
        cell.font = font_regular
        cell.border = border_box
        if is_even:
            cell.fill = fill_zebra
        if c_idx in [1, 2, 6, 8, 10, 16]:
            cell.alignment = align_center
        elif c_idx in [3, 4, 5, 7, 9]:
            cell.alignment = align_left
        else:
            cell.alignment = align_right
            if c_idx == 11:
                cell.number_format = '#,##0.0' if not float(val).is_integer() else '#,##0'
            elif c_idx in [12, 13, 14, 15, 17, 18]:
                cell.number_format = '#,##0'

# Total row Sheet 1
tot_r1 = len(rows_extracted) + 5
ws1.cell(tot_r1, 1, "TỔNG CỘNG").font = font_bold
ws1.cell(tot_r1, 1).alignment = align_center
ws1.merge_cells(start_row=tot_r1, start_column=1, end_row=tot_r1, end_column=10)
for c in range(1, 11):
    ws1.cell(tot_r1, c).fill = fill_total
    ws1.cell(tot_r1, c).border = border_total

ws1.cell(tot_r1, 11, f"=SUM(K5:K{tot_r1-1})").font = font_bold
ws1.cell(tot_r1, 11).number_format = '#,##0.0'
ws1.cell(tot_r1, 11).fill = fill_total
ws1.cell(tot_r1, 11).border = border_total
ws1.cell(tot_r1, 11).alignment = align_right

ws1.cell(tot_r1, 13, f"=SUM(M5:M{tot_r1-1})").font = font_bold
ws1.cell(tot_r1, 13).number_format = '#,##0'
ws1.cell(tot_r1, 13).fill = fill_total
ws1.cell(tot_r1, 13).border = border_total
ws1.cell(tot_r1, 13).alignment = align_right

ws1.cell(tot_r1, 14, f"=SUM(N5:N{tot_r1-1})").font = font_bold
ws1.cell(tot_r1, 14).number_format = '#,##0'
ws1.cell(tot_r1, 14).fill = fill_total
ws1.cell(tot_r1, 14).border = border_total
ws1.cell(tot_r1, 14).alignment = align_right

ws1.cell(tot_r1, 15, f"=SUM(O5:O{tot_r1-1})").font = font_bold
ws1.cell(tot_r1, 15).number_format = '#,##0'
ws1.cell(tot_r1, 15).fill = fill_total
ws1.cell(tot_r1, 15).border = border_total
ws1.cell(tot_r1, 15).alignment = align_right

ws1.cell(tot_r1, 16, "").fill = fill_total
ws1.cell(tot_r1, 16).border = border_total
ws1.cell(tot_r1, 17, "").fill = fill_total
ws1.cell(tot_r1, 17).border = border_total

ws1.cell(tot_r1, 18, f"=SUM(R5:R{tot_r1-1})").font = font_bold
ws1.cell(tot_r1, 18).number_format = '#,##0'
ws1.cell(tot_r1, 18).fill = fill_total
ws1.cell(tot_r1, 18).border = border_total
ws1.cell(tot_r1, 18).alignment = align_right

# --- SHEET 2: Tổng hợp theo Sản phẩm ---
ws2 = wb_out.create_sheet(title="Tổng hợp theo Sản phẩm")
ws2.views.sheetView[0].showGridLines = True
ws2.cell(1, 1, "TỔNG HỢP DOANH SỐ & SỐ LƯỢNG BÁN THEO TỪNG SẢN PHẨM CK - THÁNG 8/2026").font = font_title
headers2 = ["STT", "Mã Hàng (SKU)", "Tên Sản Phẩm", "ĐVT", "Phân Loại", "Mức Thưởng ĐV (VNĐ)", "Số Lượng Bán", "Số Hóa Đơn", "Doanh Thu Thô (VNĐ)", "Doanh Thu Net (VNĐ)", "Tổng Thưởng Chiết Khấu (VNĐ)"]

ws2.row_dimensions[3].height = 26
for c_idx, h_text in enumerate(headers2, start=1):
    cell = ws2.cell(3, c_idx, h_text)
    cell.font = font_header
    cell.fill = fill_header_green
    cell.alignment = align_center
    cell.border = border_box

prod_list = sorted(product_summary.values(), key=lambda x: x['total_rev'], reverse=True)
for r_idx, p in enumerate(prod_list, start=4):
    row_vals = [
        r_idx - 3, p['sku'], p['name'], p['dvt'], p['cat'], p['bonus_rate'],
        p['total_qty'], p['bill_count'], p['total_rev'], p['total_net_rev'], p['total_bonus']
    ]
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws2.cell(r_idx, c_idx, val)
        cell.font = font_regular
        cell.border = border_box
        if is_even: cell.fill = fill_zebra
        if c_idx in [1, 2, 4, 5]: cell.alignment = align_center
        elif c_idx == 3: cell.alignment = align_left
        else:
            cell.alignment = align_right
            if c_idx == 7: cell.number_format = '#,##0.0' if not float(val).is_integer() else '#,##0'
            elif c_idx == 8: cell.number_format = '#,##0'
            elif c_idx in [6, 9, 10, 11]: cell.number_format = '#,##0'

# Total row Sheet 2
tot_r2 = len(prod_list) + 4
ws2.cell(tot_r2, 1, "TỔNG CỘNG").font = font_bold
ws2.cell(tot_r2, 1).alignment = align_center
ws2.merge_cells(start_row=tot_r2, start_column=1, end_row=tot_r2, end_column=6)
for c in range(1, 7):
    ws2.cell(tot_r2, c).fill = fill_total; ws2.cell(tot_r2, c).border = border_total

for c_idx, fmla_col in [(7, 'G'), (8, 'H'), (9, 'I'), (10, 'J'), (11, 'K')]:
    cell = ws2.cell(tot_r2, c_idx, f"=SUM({fmla_col}4:{fmla_col}{tot_r2-1})")
    cell.font = font_bold
    cell.fill = fill_total
    cell.border = border_total
    cell.alignment = align_right
    cell.number_format = '#,##0.0' if c_idx == 7 else '#,##0'

# --- SHEET 3: Tổng hợp theo Dược sĩ & Nhà thuốc ---
ws3 = wb_out.create_sheet(title="Tổng hợp theo Dược sĩ")
ws3.views.sheetView[0].showGridLines = True
ws3.cell(1, 1, "TỔNG HỢP DOANH SỐ & SỐ LƯỢNG SẢN PHẨM CK THEO DƯỢC SĨ - THÁNG 8/2026").font = font_title
headers3 = ["STT", "Chi Nhánh", "Họ và Tên Dược Sĩ", "Tổng SL Sản Phẩm CK", "Số Lần Bán (Bill)", "Doanh Thu Thô CK (VNĐ)", "Doanh Thu Net CK (VNĐ)", "Tổng Thưởng Chiết Khấu (VNĐ)"]

ws3.row_dimensions[3].height = 26
for c_idx, h_text in enumerate(headers3, start=1):
    cell = ws3.cell(3, c_idx, h_text)
    cell.font = font_header
    cell.fill = fill_header_purple
    cell.alignment = align_center
    cell.border = border_box

staff_list = sorted(staff_summary.values(), key=lambda x: x['total_rev'], reverse=True)
for r_idx, s in enumerate(staff_list, start=4):
    row_vals = [
        r_idx - 3, s['branch'], s['seller'], s['total_qty'], s['bill_count'],
        s['total_rev'], s['total_net_rev'], s['total_bonus']
    ]
    is_even = (r_idx % 2 == 0)
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws3.cell(r_idx, c_idx, val)
        cell.font = font_regular
        cell.border = border_box
        if is_even: cell.fill = fill_zebra
        if c_idx == 1: cell.alignment = align_center
        elif c_idx in [2, 3]: cell.alignment = align_left
        else:
            cell.alignment = align_right
            if c_idx == 4: cell.number_format = '#,##0.0' if not float(val).is_integer() else '#,##0'
            elif c_idx in [5, 6, 7, 8]: cell.number_format = '#,##0'

# Total row Sheet 3
tot_r3 = len(staff_list) + 4
ws3.cell(tot_r3, 1, "TỔNG CỘNG").font = font_bold
ws3.cell(tot_r3, 1).alignment = align_center
ws3.merge_cells(start_row=tot_r3, start_column=1, end_row=tot_r3, end_column=3)
for c in range(1, 4):
    ws3.cell(tot_r3, c).fill = fill_total; ws3.cell(tot_r3, c).border = border_total

for c_idx, fmla_col in [(4, 'D'), (5, 'E'), (6, 'F'), (7, 'G'), (8, 'H')]:
    cell = ws3.cell(tot_r3, c_idx, f"=SUM({fmla_col}4:{fmla_col}{tot_r3-1})")
    cell.font = font_bold
    cell.fill = fill_total
    cell.border = border_total
    cell.alignment = align_right
    cell.number_format = '#,##0.0' if c_idx == 4 else '#,##0'

# Auto-adjust column widths
for ws in [ws1, ws2, ws3]:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if cell.row in [1, 2]: continue
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

out_file_thang8 = os.path.join(BASE_DIR, 'thang8/DanhSach_HoaDon_SanPham_CK_Thang8_2026.xlsx')
out_file_root = os.path.join(BASE_DIR, 'DanhSach_HoaDon_SanPham_CK_Thang8_2026.xlsx')

wb_out.save(out_file_thang8)
wb_out.save(out_file_root)

print(f"\n✅ Đã xuất thành công file Excel chi tiết hóa đơn CK ra:")
print(f"   1. {out_file_thang8}")
print(f"   2. {out_file_root}")
print(f"📊 TỔNG KẾT:")
print(f" - Tổng số dòng hóa đơn CK: {len(rows_extracted):,} dòng")
print(f" - Tổng số lượng sản phẩm CK bán: {sum(r['so_luong'] for r in rows_extracted):,.1f}")
print(f" - Tổng doanh thu thô CK: {sum(r['thanh_tien'] for r in rows_extracted):,.0f} VNĐ")
print(f" - Tổng doanh thu Net CK (sau trừ giảm giá hóa đơn): {sum(r['doanh_thu_net'] for r in rows_extracted):,.0f} VNĐ")
print(f" - Tổng tiền thưởng chiết khấu theo đơn vị: {sum(r['tong_tien_thuong'] for r in rows_extracted):,.0f} VNĐ")
