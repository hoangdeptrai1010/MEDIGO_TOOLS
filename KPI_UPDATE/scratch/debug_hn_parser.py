import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET')
from builder_engine import parse_hn_proposal

hn_file = 'd:/MEDIGO/KPI_UPDATE/thang9/data/kpi_thang9/e_xuat_KPI_Quy_3.26.xlsx'
staff, stores = parse_hn_proposal(hn_file, month_num=9)
for s in staff:
    print(f"{s['name']}: kpi_thang = {s['kpi_thang']:,.0f}, rev_ngay = {s['kpi_rev_ngay']:,.0f}")
