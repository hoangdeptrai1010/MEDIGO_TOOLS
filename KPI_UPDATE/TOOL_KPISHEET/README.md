# HƯỚNG DẪN SỬ DỤNG TOOL_KPISHEET (PORT 8767)

> **File Word hoàn chỉnh**: [`HUONG_DAN_SU_DUNG_TOOL_KPISHEET.docx`](file:///d:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/HUONG_DAN_SU_DUNG_TOOL_KPISHEET.docx)

---

## 1. Tổng Quan Về TOOL_KPISHEET
**TOOL_KPISHEET** là công cụ chuyên biệt để giải quyết bài toán khởi tạo gói kế hoạch đầu tháng:
- Tự động bóc tách danh sách nhân sự, chức danh, trung bình bill mục tiêu và doanh thu khoán từ 2 file đề xuất định dạng tự do của Quản lý Hồ Chí Minh (`KPI CNT HCM Tháng XX.xlsx`) và Quản lý Hà Nội (`e_xuat_KPI_Quy_X.XX.xlsx`).
- Tự động dựng thành file kế hoạch chuẩn (`NHÀ THUỐC THÁNG X 2026.xlsx`) với **4 sheet chuẩn** (`data`, `kpi dược sĩ`, `kpi nhà thuốc`, `Dự án TX`, `Hot Bill HN`), đầy đủ công thức Excel tự động (`IF`, `VLOOKUP`, `CEILING`, `SUMIFS`) và màu sắc phân cấp chuyên nghiệp Medigo.

---

## 2. Cấu Trúc File Trong Thư Mục `TOOL_KPISHEET/`

| Tên File / Thư Mục | Vai Trò / Chức Năng | Chi Tiết Kỹ Thuật |
| :--- | :--- | :--- |
| **`app.py`** | Máy chủ Web API (Cổng `8767`) | Hỗ trợ HTTP Multipart Upload, tải file trực tiếp với chuẩn RFC 5987 Unicode UTF-8. |
| **`builder_engine.py`** | Lõi Engine Bóc Tách & Dựng Excel | Smart Header Parser, tự động nhận diện tháng, sinh công thức IF, VLOOKUP, CEILING, SUMIFS. |
| **`kpi_styling.py`** | Module Tô Màu & Định Dạng Medigo | Palette màu Navy, Soft Blue, Emerald Green, Soft Yellow, Orange, Border chuẩn. |
| **`index.html`** | Giao diện Web Glassmorphism | Giao diện tối hiện đại, hỗ trợ Drag-and-Drop, tự động nhận diện tháng từ tên file, xem trước bảng dữ liệu. |
| **`CHAY_TOOL_KPISHEET.bat`** | File chạy 1-click cho người dùng | Tự động khởi động máy chủ web và mở trình duyệt tại `http://localhost:8767`. |
| **`uploads/`** | Thư mục lưu file nạp tạm thời | Lưu trữ các file đề xuất người dùng tải lên từ giao diện web. |
| **`output/`** | Thư mục chứa file kết quả đầu ra | Lưu file gói kế hoạch được tạo ra: `NHÀ THUỐC THÁNG X 2026.xlsx`. |

---

## 3. Cách Cài Đặt (Environment Setup)
Yêu cầu Python 3.10+ hoặc `uv`. Cài đặt thư viện:
```bash
pip install openpyxl
# Hoặc nếu dùng uv:
uv pip install openpyxl
```

---

## 4. Hướng Dẫn Sử Dụng Cho Người Không Chuyên (Non-Code)
- **Bước 1**: Nhấp đúp chuột vào file [`CHAY_TOOL_KPISHEET.bat`](file:///d:/MEDIGO/KPI_UPDATE/TOOL_KPISHEET/CHAY_TOOL_KPISHEET.bat).
- **Bước 2**: Trình duyệt web tự động mở trang **`http://localhost:8767`**. Kéo thả 2 file đề xuất HCM và HN vào 2 khung kéo thả.
- **Bước 3**: Hệ thống tự động nhận diện tháng (ví dụ: Tháng 9). Bấm nút **"TỰ ĐỘNG BÓC TÁCH & TẠO GÓI KẾ HOẠCH KPI THÁNG"**, sau đó bấm **"Tải Về Gói Kế Hoạch KPI"** để nhận file Excel.

---

## 5. Hướng Dẫn Dành Cho Lập Trình Viên (Developer / Code)
```python
from builder_engine import generate_kpisheet_package

# Khởi tạo gói kế hoạch Tháng 9 từ mã nguồn Python
result = generate_kpisheet_package(
    hcm_file_path='KPI CNT HCM Tháng 09.xlsx',
    hn_file_path='e_xuat_KPI_Quy_3.26.xlsx',
    month_num=9,
    project_folder='../thang9/Pharmacy_retail_Store_KPIs_September 2026'
)
print("Tạo thành công:", result['filename'], "Tổng nhân sự:", result['staff_count'])
```

---

## 6. Lưu Ý Quan Trọng
1. **Cơ chế Auto-Detect Month**: Hệ thống tự động quét regex tên file và Sheet names (`KPI tháng 09`, `T9.26`). Đưa file Tháng 9 vào sẽ luôn xuất đúng Tháng 9, không bao giờ bị lệch sang Tháng 8.
2. **Đồng bộ tự động**: File sau khi tạo thành công sẽ tự động đồng bộ sang `plans/`, `goc/`, và `TOOL_KPI/plans/`.
3. **Xử lý xung đột file**: Nếu file Excel đang mở trong Microsoft Excel, engine tự động lưu sang đuôi `_new.xlsx` tránh lỗi PermissionError.
