import os
import sys
import io
import argparse
import subprocess
from kpi_engine import execute_kpi_engine

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANS_DIR = os.path.join(BASE_DIR, 'plans')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Medigo KPI Calculation Engine (Data-Driven Wrapper)")
    parser.add_argument("--month", type=int, default=8, help="Số tháng cần xử lý (ví dụ: 7 hoặc 8)")
    parser.add_argument("--plan", type=str, default=None, help="Đường dẫn file kế hoạch hoặc mã kỳ (ví dụ: 2026-08)")
    parser.add_argument("--hoadon", type=str, default=None, help="Đường dẫn file hóa đơn")
    parser.add_argument("--trahang", type=str, default=None, help="Đường dẫn file trả hàng")
    parser.add_argument("--template", type=str, default=None, help="Đường dẫn file template mẫu")
    parser.add_argument("--output", type=str, default=None, help="Đường dẫn file output kết quả")
    parser.add_argument("--date", type=str, default=None, help="Ngày báo cáo (YYYY-MM-DD)")

    args = parser.parse_args()

    m = args.month
    period_str = f"2026-{m:02d}"
    plan_file = args.plan or os.path.join(PLANS_DIR, f"KeHoachKPI_{period_str}.xlsx")
    
    m_folder = os.path.join(BASE_DIR, f'thang{m}')
    
    hd_file = args.hoadon
    if not hd_file:
        candidate_paths = [
            os.path.join(m_folder, 'DATA', 'DanhSachChiTietHoaDon_3182026.xlsx'),
            os.path.join(m_folder, '288', 'DanhSachChiTietHoaDon_KV02092026-113450-301.xlsx'),
            os.path.join(m_folder, 'DanhSachChiTietHoaDon_3182026.xlsx'),
            os.path.join(m_folder, 'DATAKIOT', 'hoadon.xlsx'),
            os.path.join(m_folder, 'hoadon.xlsx')
        ]
        for cp in candidate_paths:
            if os.path.exists(cp):
                hd_file = cp
                break

    th_file = args.trahang
    if not th_file:
        candidate_th = [
            os.path.join(m_folder, 'DATA', 'DanhSachChiTietTraHang_3182026.xlsx'),
            os.path.join(m_folder, '288', 'trahang288.xlsx'),
            os.path.join(m_folder, 'DanhSachChiTietTraHang_3182026.xlsx'),
            os.path.join(m_folder, 'DATAKIOT', 'DanhSachChiTietTraHang.xlsx'),
            os.path.join(m_folder, 'DanhSachChiTietTraHang.xlsx')
        ]
        for cth in candidate_th:
            if os.path.exists(cth):
                th_file = cth
                break

    tmpl_file = args.template
    if not tmpl_file:
        tmpl_candidate = os.path.join(BASE_DIR, f'NHÀ THUỐC THÁNG {m} 2026.xlsx')
        if not os.path.exists(tmpl_candidate):
            tmpl_candidate = os.path.join(BASE_DIR, 'goc', f'NHÀ THUỐC THÁNG {m} 2026.xlsx')
        if os.path.exists(tmpl_candidate):
            tmpl_file = tmpl_candidate
        else:
            tmpl_file = os.path.join(BASE_DIR, 'NHÀ THUỐC THÁNG 8 2026.xlsx')

    out_file = args.output
    if not out_file:
        out_file = os.path.join(BASE_DIR, f'baocaokpi_thang{m}_hoanthien.xlsx')

    print(f"=== BẮT ĐẦU CHẠY DATA-DRIVEN KPI ENGINE (KỲ {period_str}) ===")
    execute_kpi_engine(hd_file, th_file, tmpl_file, plan_file, out_file, args.date)
    
    # Auto Recalculate via Excel COM to store formula values
    recalc_script = os.path.join(BASE_DIR, 'recalc_workbook.ps1')
    if os.path.exists(recalc_script):
        try:
            print(f"--> Đang tự động gọi Excel COM để tính toán và lưu sẵn 100% giá trị số KPI...")
            subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', recalc_script, '-FilePath', out_file],
                           capture_output=True, text=True, timeout=90)
            print("✅ [Auto-Recalc KPI] Đã tính toán và lưu toàn bộ giá trị số vào file KPI thành công!")
        except Exception as e:
            print(f"Warning during auto-recalc KPI: {e}")

    print("=== HOÀN THÀNH BÁO CÁO! ===")
