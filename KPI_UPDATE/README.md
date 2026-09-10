# 💊 MEDIGO PHARMACY - HỆ THỐNG ĐIỀU HÀNH & TỰ ĐỘNG HÓA KPI VÀ BẢNG LƯƠNG

Hệ thống được chia thành **2 công cụ chuyên biệt độc lập** đặt trong 2 thư mục riêng biệt:

---

## 📂 1. CẤU TRÚC HỆ THỐNG

```text
d:\MEDIGO\KPI_UPDATE\
│
├── 📁 TOOL_KPI/                        <-- TOOL 1: TẠO BÁO CÁO KPI & DỰ ÁN NHÀ THUỐC
│   ├── app.py                          (Server Web App KPI - Cổng 8765)
│   ├── index.html                      (Giao diện kéo thả Hóa đơn & Dự án)
│   ├── kpi_engine.py                   (Core Engine tính KPI theo từng Sheet chi nhánh)
│   ├── build_monthly_plan_packages.py  (Engine dựng gói kế hoạch KPI)
│   ├── build_kpi_report.py             (Script CLI chạy nhanh)
│   ├── CHAY_TOOL_KPI.bat               (Click chạy Web App KPI)
│   └── README.md                       (Hướng dẫn chi tiết Tool KPI)
│
├── 📁 TOOL_BANGLUONG/                  <-- TOOL 2: TỰ ĐỘNG HÓA BẢNG LƯƠNG & CHẤM CÔNG
│   ├── app.py                          (Server Web App Bảng Lương - Cổng 8766)
│   ├── index.html                      (Giao diện kéo thả Chấm công & Bảng Lương)
│   ├── payroll_timecards.py            (Module bóc tách Chấm công, Giờ công, Ngày công)
│   ├── payroll_minikat.py              (Module tính MiniKAT theo từng cửa hàng)
│   ├── payroll_candate.py              (Module tính thưởng hàng cận date 5%)
│   ├── payroll_project_reward.py       (Module tính thưởng dự án & Hot Bill)
│   ├── payroll_whatsapp.py             (Module tính hoa hồng WhatsApp)
│   ├── payroll_kpi.py                  (Module tính thưởng KPI doanh số)
│   ├── payroll_ck_points.py            (Module tính thưởng hàng điểm CK & cắt liều)
│   ├── payroll_deductions.py           (Module tính giảm trừ & phạt vi phạm)
│   ├── payroll_maps.py                 (Module tính thưởng Google Maps)
│   ├── payroll_main_summary.py         (Module hoàn thiện Sheet BẢNG LƯƠNG chuẩn công thức)
│   ├── build_payroll_report.py         (Orchestrator tổng hợp toàn bộ 16 sheets)
│   ├── CHAY_TOOL_BANGLUONG.bat         (Click chạy Web App Bảng Lương)
│   └── README.md                       (Hướng dẫn chi tiết Tool Bảng Lương)
│
├── CHAY_TOOL_KPI.bat                   (Shortcut chạy Tool KPI từ thư mục gốc)
├── CHAY_TOOL_BANGLUONG.bat             (Shortcut chạy Tool Bảng Lương từ thư mục gốc)
└── CHAY_APP.bat                        (Shortcut chạy ứng dụng tổng hợp)
```

---

## ⚙️ 2. YÊU CẦU CÀI ĐẶT

Cài đặt các gói thư viện Python cần thiết:
```bash
pip install pandas openpyxl
```

*(Hoặc dùng uv nếu có)*:
```bash
uv run --with pandas --with openpyxl python app.py
```

---

## 🚀 3. HƯỚNG DẪN KHỞI CHẠY NHANH

1. **Khởi chạy Tool KPI:**
   - Click đúp: **`CHAY_TOOL_KPI.bat`** *(Mở `http://localhost:8765`)*.
   - Nạp file `DanhSachChiTietHoaDon_...xlsx` &rarr; Bấm **"Xử lý & Xuất Báo Cáo KPI"** &rarr; Tải file `baocaokpi_thangX_hoanthien.xlsx`.

2. **Khởi chạy Tool Bảng Lương:**
   - Click đúp: **`CHAY_TOOL_BANGLUONG.bat`** *(Mở `http://localhost:8766`)*.
   - Nạp file `BangChiTietChamCong_thangX.xlsx` và file KPI vừa tạo &rarr; Bấm **"Tính Toán & Xuất Bảng Lương"** &rarr; Tải file `BANGLUONGTHANGX_hoanthien.xlsx`.

Xem hướng dẫn chi tiết trong từng thư mục:
- [Hướng dẫn Tool KPI](file:///d:/MEDIGO/KPI_UPDATE/TOOL_KPI/README.md)
- [Hướng dẫn Tool Bảng Lương](file:///d:/MEDIGO/KPI_UPDATE/TOOL_BANGLUONG/README.md)
