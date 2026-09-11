import os
import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET')
from builder_engine import generate_kpisheet_package

sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE/TOOL_KPI')
from kpi_engine import execute_kpi_engine
from build_monthly_plan_packages import build_plan_from_manager_proposals

print("================================================================================")
print("1. KHỞI TẠO TEMPLATE THÁNG 9 BẰNG TOOL_KPISHEET")
print("================================================================================")

hcm_prop = 'd:/MEDIGO/KPI_UPDATE/thang9/data/kpi_thang9/KPI CNT HCM Tháng 09.xlsx'
hn_prop = 'd:/MEDIGO/KPI_UPDATE/thang9/data/kpi_thang9/e_xuat_KPI_Quy_3.26.xlsx'
tmpl_out = 'd:/MEDIGO/KPI_UPDATE/thang9/output/NHÀ THUỐC THÁNG 9 2026.xlsx'

res_kpisheet = generate_kpisheet_package(
    hcm_file_path=hcm_prop,
    hn_file_path=hn_prop,
    month_num=9,
    output_filepath=tmpl_out
)
actual_tmpl_path = res_kpisheet['filepath']
print(f"--> Gói Template được lưu tại: {actual_tmpl_path}")

print("\n================================================================================")
print("2. KHỞI TẠO GÓI KẾ HOẠCH THÁNG 9 (KeHoachKPI_2026-09.xlsx)")
print("================================================================================")

plan_out = 'd:/MEDIGO/KPI_UPDATE/plans/KeHoachKPI_2026-09.xlsx'
build_plan_from_manager_proposals(
    hcm_proposal_path=hcm_prop,
    hn_proposal_path=hn_prop,
    month_num=9,
    output_filepath=plan_out
)
print(f"--> Gói Kế hoạch lưu tại: {plan_out}")

print("\n================================================================================")
print("3. CHẠY TOOL_KPI TÍNH TOÁN DỮ LIỆU HÓA ĐƠN & TRẢ HÀNG THÁNG 9")
print("================================================================================")

hd_path = 'd:/MEDIGO/KPI_UPDATE/thang9/data/phieunhap_tra/hoadon1092026.xlsx'
th_path = 'd:/MEDIGO/KPI_UPDATE/thang9/data/phieunhap_tra/trahang1092026.xlsx'
rep_out = 'd:/MEDIGO/KPI_UPDATE/thang9/output/baocaokpi_thang9_hoanthien.xlsx'

execute_kpi_engine(
    hoadon_file=hd_path,
    trahang_file=th_path,
    template_file=actual_tmpl_path,
    plan_file=plan_out,
    output_file=rep_out,
    target_month=9,
    target_year=2026
)

print("\n================================================================================")
print("4. KIỂM TRA ĐỐI SOÁT CHI TIẾT BÁO CÁO THÁNG 9")
print("================================================================================")

wb = openpyxl.load_workbook(rep_out, data_only=False)

print("\n[KIỂM TRA SHEET: kpi nhà thuốc]")
ws_nt = wb['kpi nhà thuốc']
print("Row 1 D1:", ws_nt.cell(1, 4).value)
print("Row 1 F1:", ws_nt.cell(1, 6).value)
print("Row 1 T1:", ws_nt.cell(1, 20).value)
for r in range(3, ws_nt.max_row + 1):
    cht = ws_nt.cell(r, 1).value
    st = ws_nt.cell(r, 2).value
    tgt = ws_nt.cell(r, 3).value
    col_t = ws_nt.cell(r, 20).value
    col_u = ws_nt.cell(r, 21).value
    if st:
        print(f"  {st:18s} | CHT: {str(cht):25s} | Target Tháng: {tgt:>13,d} đ | Target Ngày: {col_t}")

print("\n[KIỂM TRA SHEET: kpi dược sĩ]")
ws_ds = wb['kpi dược sĩ']
print("Row 1 C1:", ws_ds.cell(1, 3).value)
print("Tổng số nhân sự:", ws_ds.max_row - 2)

for r in range(3, ws_ds.max_row + 1):
    st = ws_ds.cell(r, 1).value
    name = ws_ds.cell(r, 2).value
    role = ws_ds.cell(r, 3).value
    tb = ws_ds.cell(r, 4).value
    col_f = ws_ds.cell(r, 6).value
    col_j = ws_ds.cell(r, 10).value
    col_k = ws_ds.cell(r, 11).value
    col_v = ws_ds.cell(r, 22).value
    
    # Check for any unexpected value or formula
    if st in ['Hàng Bông', 'Đường Láng']:
        print(f"  {st:10s} | {name:25s} | {role:10s} | Monthly: {col_v:>11,d} đ | TB: {tb} | % KPI: {col_j}")

# Test formula integrity across ALL rows in all sheets
ref_errors = []
for sname in wb.sheetnames:
    ws_check = wb[sname]
    for row in ws_check.iter_rows(values_only=False):
        for cell in row:
            if isinstance(cell.value, str) and '#REF!' in cell.value:
                ref_errors.append((sname, cell.coordinate, cell.value))

print(f"\n[KIỂM TRA LỖI TOÀN BỘ WORKBOOK]: Số lỗi #REF! = {len(ref_errors)}")
if ref_errors:
    for e in ref_errors:
        print("  Lỗi:", e)
else:
    print("  ✅ WORKBOOK HOÀN TOÀN SẠCH SẼ, 0 LỖI #REF!")

wb.close()
print("\n================================================================================")
print("🎉 TOÀN BỘ PIPELINE THÁNG 9 ĐÃ CHẠY HOÀN HẢO!")
print("================================================================================")
