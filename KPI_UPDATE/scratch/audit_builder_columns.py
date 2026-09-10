import openpyxl
import sys
import io
import os

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, 'TOOL_KPISHEET')
sys.path.insert(0, '.')

from builder_engine import parse_hcm_proposal, parse_hn_proposal

print("=== PARSING MONTH 9 ===")
staff_hcm, stores_hcm = parse_hcm_proposal('TOOL_KPISHEET/KPI CNT HCM Tháng 09.xlsx', month_num=9)
print(f"HCM Staff parsed ({len(staff_hcm)}):")
for s in staff_hcm[:10]:
    print(" ", s)
print("HCM Store Targets:", stores_hcm)

staff_hn, stores_hn = parse_hn_proposal('TOOL_KPISHEET/e_xuat_KPI_Quy_3.26.xlsx', month_num=9)
print(f"\nHN Staff parsed ({len(staff_hn)}):")
for s in staff_hn[:10]:
    print(" ", s)
print("HN Store Targets:", stores_hn)

print("\n=== PARSING MONTH 8 ===")
staff_hcm8, stores_hcm8 = parse_hcm_proposal('TOOL_KPISHEET/KPI CNT HCM Tháng 09.xlsx', month_num=8)
print(f"HCM Staff parsed Month 8 ({len(staff_hcm8)}):")
for s in staff_hcm8[:10]:
    print(" ", s)
print("HCM Store Targets Month 8:", stores_hcm8)

staff_hn8, stores_hn8 = parse_hn_proposal('TOOL_KPISHEET/e_xuat_KPI_Quy_3.26.xlsx', month_num=8)
print(f"\nHN Staff parsed Month 8 ({len(staff_hn8)}):")
for s in staff_hn8[:10]:
    print(" ", s)
print("HN Store Targets Month 8:", stores_hn8)
