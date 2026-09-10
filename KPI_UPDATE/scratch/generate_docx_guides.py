import os
import sys
import io
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_header(doc, title, subtitle):
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_title.add_run(title)
    run_t.font.name = 'Arial'
    run_t.font.size = Pt(20)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(31, 78, 121) # Navy

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_s = p_sub.add_run(subtitle)
    run_s.font.name = 'Arial'
    run_s.font.size = Pt(11)
    run_s.font.italic = True
    run_s.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

def add_callout(doc, text, title="LƯU Ý QUAN TRỌNG", box_type="info"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    
    fill_hex = "DDEBF7" if box_type == "info" else ("E2EFDA" if box_type == "success" else "FFF2CC")
    
    set_cell_background(cell, fill_hex)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    r_title = p.add_run(f"📌 {title}\n")
    r_title.bold = True
    r_title.font.name = 'Arial'
    r_title.font.size = Pt(10.5)
    r_title.font.color.rgb = RGBColor(31, 78, 121) if box_type == "info" else RGBColor(180, 83, 9)
    
    r_txt = p.add_run(text)
    r_txt.font.name = 'Arial'
    r_txt.font.size = Pt(10)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def style_heading_1(doc, text):
    h = doc.add_heading(level=1)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    r = h.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(31, 78, 121)
    return h

def style_heading_2(doc, text):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(4)
    r = h.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(51, 65, 85)
    return h

def create_table(doc, headers, rows_data, col_widths=None):
    tbl = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Header row
    hdr_cells = tbl.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1F4E79")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(10)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            
    # Data rows
    for r_idx, row in enumerate(rows_data):
        row_cells = tbl.rows[r_idx + 1].cells
        bg_hex = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row):
            row_cells[c_idx].text = str(val)
            set_cell_background(row_cells[c_idx], bg_hex)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=120, right=120)
            p = row_cells[c_idx].paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.name = 'Arial'
                r.font.size = Pt(9.5)

    if col_widths:
        for row in tbl.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)
                
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    return tbl

# ==============================================================================
# 1. TOOL KPISHEET DOCUMENTATION
# ==============================================================================
def generate_kpisheet_doc():
    doc = docx.Document()
    
    add_header(
        doc,
        "HƯỚNG DẪN SỬ DỤNG TOOL_KPISHEET",
        "Hệ Thống Tự Động Bóc Tách Đề Xuất Quản Lý & Khởi Tạo Gói Kế Hoạch KPI Tháng (Port 8767)"
    )
    
    style_heading_1(doc, "1. Tổng Quan Về TOOL_KPISHEET")
    doc.add_paragraph(
        "TOOL_KPISHEET là công cụ chuyên biệt để giải quyết bài toán khởi tạo gói kế hoạch đầu tháng. "
        "Công cụ tự động bóc tách danh sách nhân sự, chức danh, trung bình bill mục tiêu và doanh thu khoán "
        "từ 2 file đề xuất định dạng tự do của Quản lý Hồ Chí Minh và Quản lý Hà Nội, sau đó dựng thành file kế hoạch chuẩn "
        "(NHÀ THUỐC THÁNG X 2026.xlsx) với 4 sheet có đầy đủ công thức Excel tự động và màu sắc phân cấp chuyên nghiệp."
    )
    
    style_heading_1(doc, "2. Cấu Trúc Thư Mục & Các File Trong TOOL_KPISHEET")
    headers = ["Tên File / Thư Mục", "Vai Trò / Chức Năng", "Chi Tiết Kỹ Thuật"]
    rows = [
        ["app.py", "Máy chủ Web API (Cổng 8767)", "Hỗ trợ HTTP Multipart Upload, tải file trực tiếp với chuẩn RFC 5987 Unicode UTF-8."],
        ["builder_engine.py", "Lõi Engine Bóc Tách & Dựng Excel", "Smart Header Parser, tự động nhận diện tháng, sinh công thức IF, VLOOKUP, CEILING, SUMIFS."],
        ["kpi_styling.py", "Module Tô Màu & Định Dạng Medigo", "Palette màu Navy, Soft Blue, Emerald Green, Soft Yellow, Orange, Border chuẩn."],
        ["index.html", "Giao diện Web Glassmorphism", "Giao diện tối hiện đại, hỗ trợ Drag-and-Drop, tự động nhận diện tháng từ tên file, xem trước bảng dữ liệu."],
        ["CHAY_TOOL_KPISHEET.bat", "File chạy 1-click cho người dùng", "Tự động khởi động máy chủ web và tự động mở trình duyệt web tại http://localhost:8767."],
        ["uploads/", "Thư mục lưu file tạm người dùng nạp", "Lưu trữ các file đề xuất người dùng tải lên từ giao diện web."],
        ["output/", "Thư mục chứa file kết quả đầu ra", "Lưu file gói kế hoạch được tạo ra: NHÀ THUỐC THÁNG X 2026.xlsx."]
    ]
    create_table(doc, headers, rows, [1.8, 2.2, 2.5])
    
    style_heading_1(doc, "3. Hướng Dẫn Cài Đặt (Cho Người Mới)")
    doc.add_paragraph("Để chạy công cụ, máy tính chỉ cần có môi trường Python (Khuyến nghị Python 3.10+ hoặc uv package manager):")
    p_cmd = doc.add_paragraph()
    p_cmd.paragraph_format.left_indent = Inches(0.3)
    r_cmd = p_cmd.add_run("pip install openpyxl\n# Hoặc nếu sử dụng uv:\nuv pip install openpyxl")
    r_cmd.font.name = 'Consolas'
    r_cmd.font.size = Pt(9.5)
    r_cmd.font.color.rgb = RGBColor(37, 99, 235)
    
    style_heading_1(doc, "4. Hướng Dẫn Sử Dụng Cho Người Dùng Không Chuyên (Non-Code)")
    doc.add_paragraph("Người dùng chỉ cần thực hiện 3 bước đơn giản:")
    doc.add_paragraph("👉 Bước 1: Nhấp đúp chuột vào file CHAY_TOOL_KPISHEET.bat trong thư mục TOOL_KPISHEET.")
    doc.add_paragraph("👉 Bước 2: Trình duyệt web sẽ tự động mở trang http://localhost:8767. Bạn kéo thả file Đề xuất HCM (ví dụ: KPI CNT HCM Tháng 09.xlsx) và Đề xuất HN (ví dụ: e_xuat_KPI_Quy_3.26.xlsx) vào 2 ô kéo thả.")
    doc.add_paragraph("👉 Bước 3: Hệ thống sẽ tự nhận diện đúng Tháng 9. Bạn bấm nút 'TỰ ĐỘNG BÓC TÁCH & TẠO GÓI KẾ HOẠCH KPI THÁNG'. Sau 1-2 giây, bấm nút 'Tải Về Gói Kế Hoạch KPI (Excel)' để nhận file.")
    
    style_heading_1(doc, "5. Hướng Dẫn Cho Lập Trình Viên (Developer / Code)")
    doc.add_paragraph("Lập trình viên có thể tích hợp hoặc gọi trực tiếp từ dòng lệnh / mã nguồn Python:")
    
    p_code = doc.add_paragraph()
    p_code.paragraph_format.left_indent = Inches(0.3)
    r_code = p_code.add_run(
        "from builder_engine import generate_kpisheet_package\n\n"
        "# Khởi tạo kế hoạch Tháng 9 từ file đề xuất\n"
        "result = generate_kpisheet_package(\n"
        "    hcm_file_path='KPI CNT HCM Tháng 09.xlsx',\n"
        "    hn_file_path='e_xuat_KPI_Quy_3.26.xlsx',\n"
        "    month_num=9,\n"
        "    project_folder='../thang9/Pharmacy_retail_Store_KPIs_September 2026'\n"
        ")\n"
        "print('Tạo thành công:', result['filename'], 'Tổng nhân sự:', result['staff_count'])"
    )
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(9)
    r_code.font.color.rgb = RGBColor(30, 41, 59)
    
    add_callout(
        doc,
        "1. Cơ chế Auto-Detect Month: Hệ thống luôn tự động phát hiện tháng từ tên file và tên Sheet nội bộ. Đưa file Tháng 9 vào sẽ luôn xuất đúng Tháng 9, không bao giờ bị lệch sang Tháng 8.\n"
        "2. Đồng bộ tự động: File sau khi tạo thành công sẽ tự động sao chép sang thư mục plans/, goc/, và TOOL_KPI/plans/ để các tool sau sử dụng ngay mà không cần copy thủ công.\n"
        "3. Xử lý quyền truy cập file: Nếu file Excel đang được mở bởi ứng dụng khác, engine tự động lưu sang đuôi _new.xlsx tránh lỗi xung đột hệ thống.",
        "MỘT SỐ LƯU Ý ĐẶC BIỆT KHI SỬ DỤNG TOOL_KPISHEET",
        "warning"
    )
    
    out_path = os.path.join(BASE_DIR, 'TOOL_KPISHEET', 'HUONG_DAN_SU_DUNG_TOOL_KPISHEET.docx')
    doc.save(out_path)
    print(f"--> Đã tạo: {out_path}")

# ==============================================================================
# 2. TOOL KPI DOCUMENTATION
# ==============================================================================
def generate_kpi_doc():
    doc = docx.Document()
    
    add_header(
        doc,
        "HƯỚNG DẪN SỬ DỤNG TOOL_KPI",
        "Hệ Thống Tính Toán Báo Cáo KPI Bán Hàng & Thưởng Dược Sĩ Tự Động (Port 8766)"
    )
    
    style_heading_1(doc, "1. Tổng Quan Về TOOL_KPI")
    doc.add_paragraph(
        "TOOL_KPI là hệ thống tính toán cốt lõi tiếp nhận dữ liệu Hóa đơn bán hàng KiotViet và Danh sách trả hàng "
        "để tính toán chi tiết doanh thu thực tế (Offline, Online, Tổng), số giao dịch, trung bình bill, xét đạt chỉ tiêu 3 mức (80%, 90%, 100%), "
        "tính thưởng chương trình nhóm hàng dự án (CK, Combo Liều, NY3, Mini KAT, Hot Bill HN), và xuất báo cáo hoàn thiện baocaokpi_thangX_hoanthien.xlsx."
    )
    
    style_heading_1(doc, "2. Cấu Trúc Thư Mục & Các File Trong TOOL_KPI")
    headers = ["Tên File / Thư Mục", "Vai Trò / Chức Năng", "Chi Tiết Kỹ Thuật"]
    rows = [
        ["app.py", "Máy chủ Web API (Cổng 8766)", "Xử lý nạp hóa đơn, trả hàng, kế hoạch qua Web UI và gửi file tải về."],
        ["kpi_engine.py", "Lõi Tính Toán KPI (Core Engine)", "Phân loại kênh Online/Offline, tính thưởng dự án theo slide/quy chế, gộp doanh số xoay ca, đối soát tài chính."],
        ["build_kpi_report.py", "Script chạy CLI độc lập", "Cho phép chạy dòng lệnh nhanh: python build_kpi_report.py --month 9 --invoice ..."],
        ["build_monthly_plan_packages.py", "Trình dựng kế hoạch dự án", "Bóc tách quy chế từ file PowerPoint PPTX và Excel nhóm hàng."],
        ["index.html", "Giao diện Web Trực Quan", "Dashboard nạp file, hiển thị biểu đồ và bảng tiến độ."],
        ["CHAY_TOOL_KPI.bat", "File chạy 1-click cho người dùng", "Khởi chạy máy chủ và tự động mở trình duyệt tại http://localhost:8766."],
        ["plans/", "Thư mục chứa các gói kế hoạch KPI", "Chứa KeHoachKPI_2026-08.xlsx, NHÀ THUỐC THÁNG 9 2026.xlsx."],
        ["goc/", "Thư mục chứa template gốc các tháng", "Chứa NHÀ THUỐC THÁNG 8 2026.xlsx, NHÀ THUỐC THÁNG 9 2026.xlsx."]
    ]
    create_table(doc, headers, rows, [1.8, 2.2, 2.5])
    
    style_heading_1(doc, "3. Hướng Dẫn Sử Dụng Cho Người Dùng Không Chuyên (Non-Code)")
    doc.add_paragraph("👉 Bước 1: Nhấp đúp chuột vào file CHAY_TOOL_KPI.bat.")
    doc.add_paragraph("👉 Bước 2: Trình duyệt mở trang http://localhost:8766. Bạn kéo thả file Hóa đơn bán hàng (DanhSachChiTietHoaDon_...xlsx) và file Trả hàng (nếu có).")
    doc.add_paragraph("👉 Bước 3: Hệ thống tự động phát hiện tháng bán hàng từ hóa đơn. Bấm nút 'XỬ LÝ VÀ XUẤT BÁO CÁO KPI' và tải file báo cáo hoàn thiện về.")

    style_heading_1(doc, "4. Hướng Dẫn Cho Lập Trình Viên (Developer / Code)")
    doc.add_paragraph("Lập trình viên có thể thực hiện chạy kiểm thử hoặc tự động hóa dòng lệnh như sau:")
    p_code = doc.add_paragraph()
    p_code.paragraph_format.left_indent = Inches(0.3)
    r_code = p_code.add_run(
        "# Chạy qua dòng lệnh:\n"
        "python build_kpi_report.py --month 9 \\\n"
        "  --invoice \"thang9/DATA/DanhSachChiTietHoaDon_...xlsx\" \\\n"
        "  --return_file \"thang9/DATA/DanhSachChiTietTraHang_...xlsx\"\n\n"
        "# Hoặc import vào script Python:\n"
        "from kpi_engine import execute_kpi_engine\n"
        "stats = execute_kpi_engine(\n"
        "    hoadon_file='DanhSachChiTietHoaDon.xlsx',\n"
        "    trahang_file='DanhSachChiTietTraHang.xlsx',\n"
        "    template_file='goc/NHÀ THUỐC THÁNG 9 2026.xlsx',\n"
        "    plan_file='plans/NHÀ THUỐC THÁNG 9 2026.xlsx',\n"
        "    output_file='baocaokpi_thang9_hoanthien.xlsx'\n"
        ")"
    )
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(9)
    r_code.font.color.rgb = RGBColor(30, 41, 59)
    
    add_callout(
        doc,
        "1. Xử lý Nhân sự Xoay ca: 100% doanh số nhóm hàng dự án của nhân sự làm việc tại nhiều chi nhánh sẽ được gộp tự động về chi nhánh chính qua bảng STAFF_MAIN_BRANCH_MAP.\n"
        "2. Đối soát tài chính: Engine tự động kiểm tra phương trình: Doanh thu Hóa đơn - Hàng trả lại == Tổng Doanh thu tính KPI (Lệch 0 đồng).\n"
        "3. Tự động nhận diện Data Tháng: Khi nạp hóa đơn Tháng 9 nhưng dùng chương trình Tháng 8, engine sẽ tự mở dải ngày Tháng 9 và sinh Sheet Dự án T9 chuẩn xác.",
        "CÁC LƯU Ý KỸ THUẬT QUAN TRỌNG TRONG TOOL_KPI",
        "warning"
    )
    
    out_path = os.path.join(BASE_DIR, 'TOOL_KPI', 'HUONG_DAN_SU_DUNG_TOOL_KPI.docx')
    doc.save(out_path)
    print(f"--> Đã tạo: {out_path}")

# ==============================================================================
# 3. TOOL BANGLUONG DOCUMENTATION
# ==============================================================================
def generate_bangluong_doc():
    doc = docx.Document()
    
    add_header(
        doc,
        "HƯỚNG DẪN SỬ DỤNG TOOL_BANGLUONG",
        "Hệ Thống Tính Toán Bảng Lương, Giờ Công & Thưởng Tổng Hợp Tự Động (Port 8768)"
    )
    
    style_heading_1(doc, "1. Tổng Quan Về TOOL_BANGLUONG")
    doc.add_paragraph(
        "TOOL_BANGLUONG là phân hệ tính lương tự động toàn diện được thiết kế theo kiến trúc module hóa 12 thành phần. "
        "Hệ thống kết hợp dữ liệu Chấm công KiotViet, Báo cáo KPI bán hàng hoàn thiện và Dữ liệu đánh giá Google Maps để xuất "
        "bảng lương tổng hợp BẢNG LƯƠNG THÁNG X 2026.xlsx chính xác tuyệt đối từng dòng nhân sự."
    )
    
    style_heading_1(doc, "2. Cấu Trúc Thư Mục & 12 Module Lương")
    headers = ["Module / File Python", "Sheet Đảm Nhận", "Chức Năng & Nghiệp Vụ Tính Toán"]
    rows = [
        ["payroll_core_timecard.py", "Giờ công, Ngày công", "Bóc tách chấm công KiotViet, xử lý xoay ca, phát hiện ca gãy, tính công chuẩn."],
        ["payroll_kpi_adapter.py", "KPI Dược Sĩ, CHT, QLKV", "Đổ tiền thưởng KPI từ file báo cáo KPI đã được phê duyệt."],
        ["payroll_project_reward.py", "Dự án, Hot Bill HN", "Đổ thưởng nhóm hàng chiết khấu, combo và thưởng nóng hóa đơn Hà Nội."],
        ["payroll_night_shift.py", "Trực đêm", "Tính phụ cấp trực đêm (50.000 đ/đêm) cho các ca làm qua 22:00."],
        ["payroll_maps.py", "Map", "Đổ thưởng chương trình đánh giá sao Google Maps."],
        ["payroll_allowances.py", "Phụ cấp", "Tính tiền ăn trưa (30k/ngày công), điện thoại, phụ cấp trách nhiệm."],
        ["payroll_summary_builder.py", "Tổng hợp lương", "Lắp ráp toàn bộ các cột thu nhập, giảm trừ, BHXH, tạm ứng và thực lĩnh."],
        ["build_payroll_report.py", "Script tổng hợp CLI", "Điều phối toàn bộ quy trình tính lương từ dòng lệnh."],
        ["app.py & index.html", "Giao diện Web (Port 8768)", "Giao diện kéo thả file chấm công & KPI, hiển thị bảng lương trực quan."],
        ["CHAY_TOOL_BANGLUONG.bat", "File chạy 1-click", "Khởi chạy máy chủ lương và mở trình duyệt tại http://localhost:8768."]
    ]
    create_table(doc, headers, rows, [1.8, 1.8, 2.9])

    style_heading_1(doc, "3. Hướng Dẫn Sử Dụng Cho Người Dùng Không Chuyên (Non-Code)")
    doc.add_paragraph("👉 Bước 1: Nhấp đúp chuột vào file CHAY_TOOL_BANGLUONG.bat.")
    doc.add_paragraph("👉 Bước 2: Kéo thả file Chấm công (BangChiTietChamCong_thangX.xlsx) và file Báo cáo KPI (baocaokpi_thangX_hoanthien.xlsx) vào màn hình.")
    doc.add_paragraph("👉 Bước 3: Bấm nút 'XUẤT BẢNG LƯƠNG TỔNG HỢP' và tải file BẢNG LƯƠNG THÁNG X 2026.xlsx về sử dụng.")

    style_heading_1(doc, "4. Hướng Dẫn Cho Lập Trình Viên (Developer / Code)")
    p_code = doc.add_paragraph()
    p_code.paragraph_format.left_indent = Inches(0.3)
    r_code = p_code.add_run(
        "# Chạy dòng lệnh tính bảng lương Tháng 9:\n"
        "python build_payroll_report.py --month 9 \\\n"
        "  --timecard \"thang9/DATA/BangChiTietChamCong_thang9.xlsx\" \\\n"
        "  --kpi \"baocaokpi_thang9_hoanthien.xlsx\" \\\n"
        "  --output \"BẢNG LƯƠNG THÁNG 9 2026.xlsx\""
    )
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(9)
    r_code.font.color.rgb = RGBColor(30, 41, 59)

    add_callout(
        doc,
        "1. Khắc phục lỗi thiếu dòng nhân sự xoay ca: Hệ thống sử dụng cơ chế gom dòng đa chi nhánh dựa trên 74 cặp (Chi nhánh, Nhân viên) thực tế trong file KiotViet, đảm bảo không bỏ sót giờ công nào khi nhân viên tăng cường sang chi nhánh khác.\n"
        "2. Kiểm tra công thức Excel: Các sheet Giờ công, Ngày công, Tổng hợp lương đều giữ nguyên công thức Excel sống (VLOOKUP, SUMIFS, IF) giúp kế toán kiểm tra và đối chiếu minh bạch.",
        "LƯU Ý ĐỐI SOÁT CHẤM CÔNG & LƯƠNG",
        "warning"
    )

    out_path = os.path.join(BASE_DIR, 'TOOL_BANGLUONG', 'HUONG_DAN_SU_DUNG_TOOL_BANGLUONG.docx')
    doc.save(out_path)
    print(f"--> Đã tạo: {out_path}")

# ==============================================================================
# 4. MASTER COMBINED DOCUMENTATION (ROOT PORTAL)
# ==============================================================================
def generate_master_doc():
    doc = docx.Document()
    
    add_header(
        doc,
        "CẨM NANG HƯỚNG DẪN TOÀN DIỆN BỘ CÔNG CỤ MEDIGO AUTOMATION",
        "Tài Liệu Hướng Dẫn Chi Tiết Dành Cho Người Dùng Không Chuyên & Lập Trình Viên"
    )
    
    style_heading_1(doc, "1. Giới Thiệu Tổng Quan Hệ Sinh Thái Công Cụ Medigo")
    doc.add_paragraph(
        "Hệ thống tự động hóa Medigo bao gồm 3 phân hệ độc lập và 1 Cổng điều phối trung tâm, "
        "tạo thành một chu trình khép kín từ đầu tháng đến khi hoàn tất thanh toán bảng lương:"
    )
    doc.add_paragraph("1. TOOL_KPISHEET (Cổng 8767): Bóc tách đề xuất đầu tháng của quản lý HCM & HN ➔ Tạo gói Kế hoạch KPI chuẩn.")
    doc.add_paragraph("2. TOOL_KPI (Cổng 8766): Xử lý hóa đơn bán hàng KiotViet ➔ Xuất báo cáo kết quả KPI Dược sĩ & Cửa hàng.")
    doc.add_paragraph("3. TOOL_BANGLUONG (Cổng 8768): Xử lý chấm công & kết quả KPI ➔ Xuất bảng lương tổng hợp 12 module.")
    doc.add_paragraph("4. PORTAL TỔNG HỢP (Cổng 8765): Cổng giao diện hợp nhất điều hành toàn bộ quy trình.")

    style_heading_1(doc, "2. Bảng Danh Mục File & Cổng Port Sử Dụng")
    headers = ["Phân Hệ", "Cổng Port", "File Chạy 1-Click", "File Đầu Vào", "File Đầu Ra"]
    rows = [
        ["TOOL_KPISHEET", "8767", "CHAY_TOOL_KPISHEET.bat", "Đề xuất HCM & Đề xuất HN", "NHÀ THUỐC THÁNG X 2026.xlsx"],
        ["TOOL_KPI", "8766", "CHAY_TOOL_KPI.bat", "Hóa đơn KiotViet + Gói kế hoạch", "baocaokpi_thangX_hoanthien.xlsx"],
        ["TOOL_BANGLUONG", "8768", "CHAY_TOOL_BANGLUONG.bat", "Chấm công KiotViet + Báo cáo KPI", "BẢNG LƯƠNG THÁNG X 2026.xlsx"],
        ["PORTAL TỔNG", "8765", "CHAY_HE_THONG.bat", "Toàn bộ file dữ liệu", "Gói Kế hoạch + KPI + Bảng lương"]
    ]
    create_table(doc, headers, rows, [1.5, 0.9, 1.8, 1.8, 1.8])

    style_heading_1(doc, "3. Bảng Tóm Tắt Quy Trình 3 Bước Dành Cho Người Non-Code")
    doc.add_paragraph("Mỗi công cụ đều được trang bị tệp .bat khởi động 1-click. Người dùng không cần cài đặt phức tạp:")
    doc.add_paragraph("• Bước 1: Nhấp đúp vào file .bat tương ứng để mở giao diện Web.")
    doc.add_paragraph("• Bước 2: Kéo thả các file Excel được yêu cầu vào khung nạp dữ liệu.")
    doc.add_paragraph("• Bước 3: Bấm nút Xử lý và tải file kết quả Excel đã được tính toán và tô màu hoàn thiện.")

    style_heading_1(doc, "4. Hướng Dẫn Kiến Trúc & Tích Hợp Dành Cho Lập Trình Viên")
    doc.add_paragraph("Hệ thống tuân thủ các nguyên tắc thiết kế phần mềm sạch (Clean Architecture):")
    doc.add_paragraph("• Zero-Dependency Server: Sử dụng http.server tiêu chuẩn kết hợp bộ bóc tách Multipart form-data tự xây dựng, không bị phụ thuộc vào Flask/FastAPI.")
    doc.add_paragraph("• Smart Auto-Detection: Mọi engine đều có cơ chế đọc mẫu header và ngày tháng từ nội dung file, loại bỏ hoàn toàn việc bị nhầm lẫn giữa các kỳ tháng.")
    doc.add_paragraph("• Unicode RFC 5987 / RFC 6266: Xử lý triệt để lỗi mã hóa tiếng Việt trong header HTTP khi tải file.")

    add_callout(
        doc,
        "• Khi gặp lỗi cổng mạng đang bị chiếm dụng, kiểm tra tác vụ nền hoặc đổi port trong file app.py.\n"
        "• Dữ liệu nhân viên có dấu tiếng Việt cần được giữ nguyên định dạng Unicode UTF-8 để đảm bảo các hàm VLOOKUP/SUMIFS khớp chính xác 100%.\n"
        "• Tất cả file đầu ra đều lưu tự động vào thư mục output/ và đồng bộ sang plans/ hoặc goc/.",
        "TỔNG KẾT & KHẮC PHỤC SỰ CỐ NHANH",
        "info"
    )

    out_path = os.path.join(BASE_DIR, 'HUONG_DAN_SU_DUNG_TOAN_BO_HE_THONG_MEDIGO_KPI.docx')
    doc.save(out_path)
    print(f"--> Đã tạo: {out_path}")

if __name__ == '__main__':
    generate_kpisheet_doc()
    generate_kpi_doc()
    generate_bangluong_doc()
    generate_master_doc()
    print("=== TẤT CẢ 4 FILE DOCX ĐÃ ĐƯỢC TẠO THÀNH CÔNG ===")
