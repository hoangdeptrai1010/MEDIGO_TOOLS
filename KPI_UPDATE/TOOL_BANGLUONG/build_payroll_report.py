"""
HỆ THỐNG TỔNG HỢP BẢNG LƯƠNG MEDIGO (ENTERPRISE PAYROLL ORCHESTRATOR)
Tập hợp tất cả các module tính toán con:
  1. payroll_timecards.py       -> Bóc tách máy chấm công, tính Giờ công & Ngày công
  2. payroll_minikat.py         -> Tính thưởng MiniKAT từng cửa hàng (HN & HCM)
  3. payroll_candate.py         -> Tính thưởng Hàng cận date (5% HSD <= 6 tháng)
  4. payroll_project_reward.py  -> Tính thưởng Dự án nhóm hàng & Hot Bill
  5. payroll_whatsapp.py        -> Tính thưởng hoa hồng WhatsApp (1.5% - 6.0%)
  6. payroll_kpi.py             -> Tính thưởng KPI Doanh số
  7. payroll_ck_points.py       -> Tính thưởng Hàng điểm CK & Cắt liều (50k HN)
  8. payroll_deductions.py      -> Tính Giảm trừ KPI & Phạt vi phạm
  9. payroll_maps.py            -> Tính thưởng Đánh giá Google Maps
  10. payroll_main_summary.py   -> Hoàn thiện Sheet trung tâm 'BẢNG LƯƠNG'
"""

import os
import sys
import io
import argparse
import subprocess
import openpyxl

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if _CURRENT_DIR not in sys.path:
    sys.path.insert(0, _CURRENT_DIR)

# Import các module con
from payroll_timecards import process_timecard_sheets
from payroll_minikat import process_minikat_sheets
from payroll_candate import process_candate_sheet
from payroll_project_reward import process_project_reward_sheet
from payroll_whatsapp import process_whatsapp_sheet
from payroll_kpi import process_kpi_sheet
from payroll_ck_points import process_ck_points_sheet
from payroll_deductions import process_deductions_sheet
from payroll_maps import process_maps_sheet
from payroll_main_summary import process_main_payroll_sheet

def standardize_formula(f_str):
    if not isinstance(f_str, str): return f_str
    res = f_str.replace('_xludf.', '_xlfn.').replace('_xlfn.SUMIFS(', 'SUMIFS(').replace('_xlfn.COUNTIFS(', 'COUNTIFS(').replace('_xlfn.AVERAGEIFS(', 'AVERAGEIFS(')
    for std_fn in ['IF', 'SUM', 'INDEX', 'MATCH', 'VLOOKUP', 'COUNTIFS', 'SUMIFS', 'AVERAGEIFS', 'MIN', 'MAX', 'AND', 'OR', 'ISNUMBER', 'SEARCH', 'ROUND', 'TRIM']:
        res = res.replace(f'_xlfn.{std_fn}(', f'{std_fn}(')
    return res

def generate_payroll_report_perfect(kpi_file, template_file, output_file, month=8, inv_file=None, ret_file=None, timecard_file=None):
    print("=" * 80)
    print(f"=== BẮT ĐẦU TỰ ĐỘNG HÓA BẢNG LƯƠNG THÁNG {month} ===")
    print("=" * 80)
    print(f"1. File Template Mẫu: {template_file}")
    print(f"2. File Chấm Công:   {timecard_file}")
    print(f"3. File Báo Cáo KPI: {kpi_file}")
    print(f"4. File Hóa Đơn:     {inv_file}")
    print(f"5. File Trả Hàng:    {ret_file}")
    print(f"6. File Output:      {output_file}")
    print("-" * 80)

    # 1. Nạp workbook template
    if not template_file or not os.path.exists(template_file):
        raise FileNotFoundError(f"Không tìm thấy file template bảng lương: {template_file}")

    wb_out = openpyxl.load_workbook(template_file)

    # 2. Chạy Module 1: Chấm công (Giờ công & Ngày công)
    if timecard_file and os.path.exists(timecard_file):
        process_timecard_sheets(wb_out, timecard_file, month=month)

    # 3. Chạy Module 2: MiniKAT
    process_minikat_sheets(wb_out, inv_file, ret_file)

    # 4. Chạy Module 3: Hàng cận date
    process_candate_sheet(wb_out, inv_file, ret_file)

    # 5. Chạy Module 4: Thưởng dự án & Hot Bill
    process_project_reward_sheet(wb_out, kpi_file)

    # 6. Chạy Module 5: WhatsApp
    process_whatsapp_sheet(wb_out, inv_file, month=month)

    # 7. Chạy Module 6: KPI Doanh số
    process_kpi_sheet(wb_out, kpi_file)

    # 8. Chạy Module 7: Thưởng CK Hàng điểm
    process_ck_points_sheet(wb_out)

    # 9. Chạy Module 8: Giảm trừ KPI
    process_deductions_sheet(wb_out)

    # 10. Chạy Module 9: Maps
    process_maps_sheet(wb_out)

    # 11. Chạy Module 10: Sheet trung tâm 'BẢNG LƯƠNG'
    process_main_payroll_sheet(wb_out)

    # 12. Chuẩn hóa công thức OpenXML & Unhide rows
    print("--> Đang chuẩn hóa công thức OpenXML và mở ẩn toàn bộ dòng cho 16 sheets...")
    for ws_c in wb_out.worksheets:
        for dim in ws_c.row_dimensions.values():
            if dim.hidden: dim.hidden = False
        if ws_c.auto_filter: ws_c.auto_filter.filterColumn = []

        for cell in ws_c._cells.values():
            val = cell.value
            if isinstance(val, str) and val.startswith('='):
                cell.value = standardize_formula(val)
            elif hasattr(val, 'text') and isinstance(val.text, str):
                val.text = standardize_formula(val.text)

    try:
        wb_out.calculation.fullCalcOnLoad = True
        wb_out.calculation.calcMode = 'auto'
    except:
        pass

    wb_out.save(output_file)
    print(f"--> BẢNG LƯƠNG THÁNG {month} ĐÃ XUẤT THÀNH CÔNG RA: {output_file}")

    # 13. Auto Recalculate via PowerShell Excel COM
    base_dir = os.path.dirname(os.path.abspath(__file__))
    recalc_script = os.path.join(base_dir, 'recalc_workbook.ps1')
    if os.path.exists(recalc_script):
        try:
            print(f"--> Đang tự động gọi Excel COM để tính toán và lưu sẵn 100% giá trị số thực tế...")
            res = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', recalc_script, '-FilePath', output_file],
                                 capture_output=True, text=True, timeout=90)
            if res.returncode == 0:
                print("✅ [Auto-Recalc] Đã tính toán và lưu toàn bộ giá trị số vào file thành công!")
        except Exception as e:
            print(f"Warning during auto-recalc: {e}")

    print("=" * 80)
    print(f"=== HOÀN THÀNH TỰ ĐỘNG HÓA BẢNG LƯƠNG THÁNG {month}! ===")
    return output_file

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Medigo Modular Payroll Orchestrator")
    parser.add_argument('--month', type=int, default=8)
    parser.add_argument('--kpi', type=str, default=None)
    parser.add_argument('--template', type=str, default=None)
    parser.add_argument('--output', type=str, default=None)
    parser.add_argument('--invoice', type=str, default=None)
    parser.add_argument('--return_file', type=str, default=None)
    parser.add_argument('--timecard', type=str, default=None)
    args = parser.parse_args()

    m = args.month
    base_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(base_dir)
    m_folder = os.path.join(parent_dir, f'thang{m}')

    kpi_file = args.kpi or os.path.join(parent_dir, f'baocaokpi_thang{m}_hoanthien.xlsx')
    tmpl_file = args.template or os.path.join(m_folder, 'tinhcongnhungthuongchia.xlsx')
    out_file = args.output or os.path.join(m_folder, f'BANGLUONGTHANG{m}.xlsx')
    inv_file = args.invoice or os.path.join(m_folder, 'DATA', 'DanhSachChiTietHoaDon_3182026.xlsx')
    ret_file = args.return_file or os.path.join(m_folder, 'DATA', 'DanhSachChiTietTraHang_3182026.xlsx')
    timecard_file = args.timecard or os.path.join(m_folder, 'DATA', f'BangChiTietChamCong_thang{m}.xlsx')

    generate_payroll_report_perfect(kpi_file, tmpl_file, out_file, m, inv_file, ret_file, timecard_file)
