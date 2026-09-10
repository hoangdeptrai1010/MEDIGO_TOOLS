import os, sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE')

from TOOL_KPI.kpi_engine import execute_kpi_engine
from TOOL_KPI.build_monthly_plan_packages import build_plan_from_manager_proposals

# 1. Test build plan from proposals in TOOL_KPISHEET
hcm_proposal = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/KPI CNT HCM Tháng 09.xlsx'
hn_proposal = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/e_xuat_KPI_Quy_3.26.xlsx'
plan_res = build_plan_from_manager_proposals(hcm_proposal, hn_proposal, month_num=9)
print('Plan built:', plan_res)

# 2. Run engine on August data with August plan
hoadon_file = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietHoaDon_3182026.xlsx'
trahang_file = 'd:/MEDIGO/KPI_UPDATE/thang8/DATA/DanhSachChiTietTraHang_3182026.xlsx'
template_file = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPI/goc/NHÀ THUỐC THÁNG 8 2026.xlsx'
plan_file = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPI/plans/KeHoachKPI_2026-08.xlsx'
output_file = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPI/baocaokpi_thang8_test.xlsx'

stats = execute_kpi_engine(
    hoadon_file=hoadon_file,
    trahang_file=trahang_file,
    template_file=template_file,
    plan_file=plan_file,
    output_file=output_file,
    report_date_str='2026-08-31',
    target_month=8,
    target_year=2026
)

print('\n=== KPI Engine Run Result ===')
print('Total Rev:', f"{stats.get('total_rev', 0):,.0f} VNĐ")
print('Total Invoices:', stats.get('total_invoices'))
print('Active Staff:', stats.get('active_staff'))
print('Hot Bills Count:', stats.get('hot_bills_count'))
print('Output File Size:', os.path.getsize(output_file), 'bytes')
