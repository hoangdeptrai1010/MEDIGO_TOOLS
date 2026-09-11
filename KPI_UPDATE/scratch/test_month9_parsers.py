import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET')
from builder_engine import parse_hn_proposal, parse_hcm_proposal

hn_file = 'd:/MEDIGO/KPI_UPDATE/thang9/data/kpi_thang9/e_xuat_KPI_Quy_3.26.xlsx'
hcm_file = 'd:/MEDIGO/KPI_UPDATE/thang9/data/kpi_thang9/KPI CNT HCM Tháng 09.xlsx'

hn_staff, hn_stores = parse_hn_proposal(hn_file, 9)
hcm_staff, hcm_stores = parse_hcm_proposal(hcm_file, 9)

print('=== HN STORES ===')
for k, v in hn_stores.items():
    print(f'  {k}: {v:,.0f}')

print('\n=== HN STAFF ===')
for s in hn_staff:
    print(f"  {s['store']} - {s['name']} ({s['role']}): monthly={s['kpi_thang']:,.0f}, daily={s.get('kpi_rev_ngay', 0):,.0f}, tb_bill={s['tb_bill']}")

print('\n=== HCM STORES ===')
for k, v in hcm_stores.items():
    print(f'  {k}: {v:,.0f}')

print('\n=== HCM STAFF COUNT ===', len(hcm_staff))
