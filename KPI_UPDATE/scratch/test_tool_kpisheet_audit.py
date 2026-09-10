import openpyxl
import os
import sys
import json
import urllib.request
import urllib.parse

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET')

from builder_engine import generate_kpisheet_package, parse_hcm_proposal, parse_hn_proposal

def test_kpisheet():
    print("=" * 70)
    print("🧪 [TEST 1] Kiểm tra hàm bóc tách dữ liệu QL HCM & QL HN")
    print("=" * 70)
    
    hcm_file = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/KPI CNT HCM Tháng 09.xlsx'
    hn_file = 'd:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/e_xuat_KPI_Quy_3.26.xlsx'
    
    hcm_staff, hcm_stores = parse_hcm_proposal(hcm_file, month_num=9)
    print(f"✅ HCM Staff count: {len(hcm_staff)}")
    print(f"✅ HCM Stores count: {len(hcm_stores)}")
    print("   Stores detail:", hcm_stores)
    
    hn_staff, hn_stores = parse_hn_proposal(hn_file, month_num=9)
    print(f"✅ HN Staff count: {len(hn_staff)}")
    print(f"✅ HN Stores count: {len(hn_stores)}")
    print("   Stores detail:", hn_stores)

    print("\n" + "=" * 70)
    print("🧪 [TEST 2] Kiểm tra sinh gói kế hoạch Tháng 9 hoàn chỉnh")
    print("=" * 70)
    
    proj_folder = 'd:/MEDIGO/KPI_UPDATE/thang8/Pharmacy_retail_Store_KPIs_August 2026'
    res = generate_kpisheet_package(
        hcm_file_path=hcm_file,
        hn_file_path=hn_file,
        month_num=9,
        project_folder=proj_folder
    )
    print("✅ Build Result Summary:", json.dumps(res, ensure_ascii=False, indent=2))
    
    out_path = res['filepath']
    assert os.path.exists(out_path), f"File {out_path} không tồn tại!"
    
    print("\n" + "=" * 70)
    print("🧪 [TEST 3] Kiểm tra chi tiết 6 Sheet trong file Excel kết quả")
    print("=" * 70)
    
    wb = openpyxl.load_workbook(out_path, data_only=True)
    expected_sheets = ['Nhân sự', 'KPI nhà thuốc', 'KPI nhân viên', 'Danh mục dự án', 'Quy tắc thưởng', 'Chương trình đặc biệt']
    print(f"Các Sheet thực tế trong file: {wb.sheetnames}")
    for es in expected_sheets:
        assert es in wb.sheetnames, f"Thiếu sheet: {es}"
        ws = wb[es]
        print(f"  - Sheet [{es}]: {ws.max_row} dòng, {ws.max_column} cột")
        # In 3 dòng đầu
        for r in range(1, min(4, ws.max_row + 1)):
            vals = [ws.cell(r, c).value for c in range(1, min(10, ws.max_column + 1))]
            print(f"      Row {r}: {vals}")

    # Check staff integrity
    ws_staff = wb['Nhân sự']
    staff_rows = ws_staff.max_row - 1
    print(f"\n✅ Tổng số nhân sự trong Sheet Nhân sự: {staff_rows}")
    assert staff_rows >= 50, f"Số lượng nhân sự {staff_rows} ít hơn kỳ vọng (51)!"
    
    # Check stores integrity
    ws_stores = wb['KPI nhà thuốc']
    store_rows = ws_stores.max_row - 1
    print(f"✅ Tổng số nhà thuốc trong Sheet KPI nhà thuốc: {store_rows}")
    assert store_rows == 11, f"Số lượng nhà thuốc {store_rows} != 11!"

    print("\n" + "=" * 70)
    print("🎉 TẤT CẢ CÁC BƯỚC KIỂM TRA ĐỀU CHÍNH XÁC 100%!")
    print("=" * 70)

if __name__ == '__main__':
    test_kpisheet()
