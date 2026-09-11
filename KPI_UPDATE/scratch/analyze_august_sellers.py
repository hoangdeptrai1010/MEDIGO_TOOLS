from python_calamine import CalamineWorkbook
import sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

f = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
wb = CalamineWorkbook.from_path(f)
sheet = wb.get_sheet_by_name(wb.sheet_names[0])
rows = sheet.to_python()

sellers = {
    'Ngô Thị Thanh Thắm': 'Đỗ Quang Đẩu',
    'Hoàng Thanh Thủy': 'Đỗ Quang Đẩu',
    'Lê Thị Huyền Trân': 'Đỗ Quang Đẩu',
    'Trần Thiên Phát': 'Đỗ Quang Đẩu',
    'Phạm Thị Nghĩa Hương': 'Đỗ Quang Đẩu',
    'Đinh Thị Khánh Ly': 'Hàng Bông'
}

inv_by_seller = {s: {} for s in sellers}

for r in rows[1:]:
    s = str(r[7] or '').strip()
    if s in sellers:
        code = str(r[1] or '').strip()
        if code not in inv_by_seller[s]:
            inv_by_seller[s][code] = {
                'branch': r[0],
                'code': code,
                'time': str(r[2])[:19] if r[2] else '',
                'cust_id': r[5],
                'channel': r[8],
                'price_list': r[6],
                'need_pay': float(r[12]) if r[12] is not None else 0.0,
                'notes': set()
            }
        note = str(r[20] or '').strip()
        if note and note != 'None':
            inv_by_seller[s][code]['notes'].add(note)

for s, invs in inv_by_seller.items():
    tot_rev = sum(inv['need_pay'] for inv in invs.values())
    print(f"Seller: {s:25} | Invoices: {len(invs):5} | Total Rev: {tot_rev:15,.0f} đ")
    for code, inv in list(invs.items())[:3]:
        t = inv['time']
        ch = inv['channel']
        pl = inv['price_list']
        np = inv['need_pay']
        print(f"   {code} | {t} | Channel: {ch} | PriceList: {pl} | Rev: {np:,.0f} đ")
