import openpyxl
import sys
import io
import os
import json

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("=====================================================================")
print("PART A: AUDIT PROPOSAL INPUT vs OUTPUT IN TOOL_KPISHEET (THÁNG 9)")
print("=====================================================================")
hcm_p = 'TOOL_KPISHEET/KPI CNT HCM Tháng 09.xlsx'
hn_p = 'TOOL_KPISHEET/e_xuat_KPI_Quy_3.26.xlsx'
out_p = 'TOOL_KPISHEET/output/NHÀ THUỐC THÁNG 9 2026.xlsx'

wb_hcm = openpyxl.load_workbook(hcm_p, data_only=True)
wb_hn = openpyxl.load_workbook(hn_p, data_only=True)
wb_out = openpyxl.load_workbook(out_p, data_only=True)

# 1. Check HCM Proposal Staff & Targets
print("\n--- 1.1 Checking HCM Proposal Sheets ---")
for sname in wb_hcm.sheetnames:
    ws = wb_hcm[sname]
    print(f"HCM Sheet '{sname}': max_row={ws.max_row}, max_col={ws.max_column}")

# Let's inspect rows in HCM proposal sheet 'Target T9' or similar
ws_hcm = wb_hcm.active
print(f"Active HCM sheet '{ws_hcm.title}' first 10 rows:")
for r in range(1, min(15, ws_hcm.max_row + 1)):
    vals = [ws_hcm.cell(r, c).value for c in range(1, min(12, ws_hcm.max_column + 1))]
    print(f"  Row {r:2d}: {vals}")

# 2. Check HN Proposal Sheets
print("\n--- 1.2 Checking HN Proposal Sheets ---")
for sname in wb_hn.sheetnames:
    ws = wb_hn[sname]
    print(f"HN Sheet '{sname}': max_row={ws.max_row}, max_col={ws.max_column}")
    for r in range(1, min(10, ws.max_row + 1)):
        vals = [ws.cell(r, c).value for c in range(1, min(12, ws.max_column + 1))]
        if any(v is not None for v in vals):
            print(f"  Row {r:2d}: {vals}")

# 3. Check what was written in Output 'kpi dược sĩ' and 'kpi nhà thuốc'
print("\n--- 1.3 Output Sheet 'kpi dược sĩ' in THÁNG 9 ---")
ws_ds = wb_out['kpi dược sĩ']
for r in range(1, min(20, ws_ds.max_row + 1)):
    vals = [ws_ds.cell(r, c).value for c in range(1, min(25, ws_ds.max_column + 1))]
    print(f"  DS Row {r:2d}: {vals[:10]}")

print("\n--- 1.4 Output Sheet 'kpi nhà thuốc' in THÁNG 9 ---")
ws_nt = wb_out['kpi nhà thuốc']
for r in range(1, ws_nt.max_row + 1):
    vals = [ws_nt.cell(r, c).value for c in range(1, min(25, ws_nt.max_column + 1))]
    print(f"  NT Row {r:2d}: {vals[:10]}")
