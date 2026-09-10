import openpyxl, sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE')

from TOOL_BANGLUONG.payroll_candate import scan_candate_sales

inv_file = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
ret_file = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietTraHang_3182026.xlsx'

candate_res = scan_candate_sales(inv_file, ret_file)
print(f'Test scan_candate_sales completed cleanly: {len(candate_res)} staff.')

lb_rev = candate_res.get(('Lê Bình', 'Hoàng Lâm Gia Bảo'), 0)
ntt_rev = candate_res.get(('Nguyễn Thị Thập', 'Hoàng Lâm Gia Bảo'), 0)

print(f'Hoàng Lâm Gia Bảo (Lê Bình): {lb_rev:,.0f} đ -> Thưởng 5% = {lb_rev*0.05:,.0f} đ')
print(f'Hoàng Lâm Gia Bảo (Nguyễn Thị Thập): {ntt_rev:,.0f} đ -> Thưởng 5% = {ntt_rev*0.05:,.0f} đ')
print(f'Tổng thưởng cận date của Bảo = {(lb_rev + ntt_rev)*0.05:,.0f} đ')
