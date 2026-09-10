# HƯỚNG DẪN SỬ DỤNG TOOL_KPI (PORT 8766)

> **File Word hoàn chỉnh**: [`HUONG_DAN_SU_DUNG_TOOL_KPI.docx`](file:///d:/MEDIGO/KPI_UPDATE/TOOL_KPI/HUONG_DAN_SU_DUNG_TOOL_KPI.docx)

---

## 1. Tổng Quan Về TOOL_KPI
**TOOL_KPI** là hệ thống tính toán KPI bán hàng cốt lõi:
- Tiếp nhận dữ liệu Hóa đơn bán hàng KiotViet (`DanhSachChiTietHoaDon_...xlsx`) và Danh sách trả hàng (`DanhSachChiTietTraHang_...xlsx`).
- Tính toán chi tiết doanh thu thực tế (Offline, Online, Tổng), số giao dịch, trung bình bill, xét đạt chỉ tiêu 3 mức (80%, 90%, 100%).
- Tính thưởng chương trình nhóm hàng dự án (CK, Combo Liều, NY3, Mini KAT, Hot Bill HN), và xuất báo cáo hoàn thiện `baocaokpi_thangX_hoanthien.xlsx`.

---

## 2. Cấu Trúc File Trong Thư Mục `TOOL_KPI/`

| Tên File / Thư Mục | Vai Trò / Chức Năng | Chi Tiết Kỹ Thuật |
| :--- | :--- | :--- |
| **`app.py`** | Máy chủ Web API (Cổng `8766`) | Xử lý nạp hóa đơn, trả hàng, kế hoạch qua Web UI và gửi file tải về. |
| **`kpi_engine.py`** | Lõi Tính Toán KPI (Core Engine) | Phân loại kênh Online/Offline, tính thưởng dự án theo slide/quy chế, gộp doanh số xoay ca, đối soát tài chính. |
| **`build_kpi_report.py`** | Script chạy CLI độc lập | Cho phép chạy dòng lệnh nhanh: `python build_kpi_report.py --month 9 --invoice ...` |
| **`build_monthly_plan_packages.py`** | Trình dựng kế hoạch dự án | Bóc tách quy chế từ file PowerPoint PPTX và Excel nhóm hàng. |
| **`index.html`** | Giao diện Web Trực Quan | Dashboard nạp file, hiển thị biểu đồ và bảng tiến độ. |
| **`CHAY_TOOL_KPI.bat`** | File chạy 1-click cho người dùng | Khởi chạy máy chủ và tự động mở trình duyệt tại `http://localhost:8766`. |
| **`plans/`** | Thư mục chứa gói kế hoạch KPI | Chứa `KeHoachKPI_2026-08.xlsx`, `NHÀ THUỐC THÁNG 9 2026.xlsx`. |
| **`goc/`** | Thư mục chứa template gốc các tháng | Chứa `NHÀ THUỐC THÁNG 8 2026.xlsx`, `NHÀ THUỐC THÁNG 9 2026.xlsx`. |

---

## 3. Cách Cài Đặt (Environment Setup)
```bash
pip install openpyxl
# Hoặc nếu dùng uv:
uv pip install openpyxl
```

---

## 4. Hướng Dẫn Sử Dụng Cho Người Không Chuyên (Non-Code)
- **Bước 1**: Nhấp đúp chuột vào file [`CHAY_TOOL_KPI.bat`](file:///d:/MEDIGO/KPI_UPDATE/TOOL_KPI/CHAY_TOOL_KPI.bat).
- **Bước 2**: Trình duyệt mở trang **`http://localhost:8766`**. Kéo thả file Hóa đơn bán hàng KiotViet và file Trả hàng (nếu có).
- **Bước 3**: Bấm nút **"XỬ LÝ VÀ XUẤT BÁO CÁO KPI"** và tải file báo cáo hoàn thiện về.

---

## 5. Hướng Dẫn Dành Cho Lập Trình Viên (Developer / Code)
```bash
# Chạy dòng lệnh độc lập:
python build_kpi_report.py --month 9 \
  --invoice "thang9/DATA/DanhSachChiTietHoaDon_...xlsx" \
  --return_file "thang9/DATA/DanhSachChiTietTraHang_...xlsx"
```

```python
# Gọi từ Python script:
from kpi_engine import execute_kpi_engine
stats = execute_kpi_engine(
    hoadon_file='DanhSachChiTietHoaDon.xlsx',
    trahang_file='DanhSachChiTietTraHang.xlsx',
    template_file='goc/NHÀ THUỐC THÁNG 9 2026.xlsx',
    plan_file='plans/NHÀ THUỐC THÁNG 9 2026.xlsx',
    output_file='baocaokpi_thang9_hoanthien.xlsx'
)
```

---

## 6. Lưu Ý Quan Trọng
1. **Xử lý Nhân sự Xoay ca**: 100% doanh số nhóm hàng dự án của nhân sự làm việc tại nhiều chi nhánh sẽ được gộp tự động về chi nhánh chính qua bảng `STAFF_MAIN_BRANCH_MAP`.
2. **Đối soát tài chính**: Engine tự động kiểm tra phương trình: `Doanh thu Hóa đơn - Hàng trả lại == Tổng Doanh thu tính KPI` (Lệch 0 đồng).
3. **Tự động nhận diện Data Tháng**: Khi nạp hóa đơn Tháng 9 nhưng dùng chương trình Tháng 8, engine sẽ tự mở dải ngày Tháng 9 và sinh Sheet `Dự án T9` chuẩn xác.
