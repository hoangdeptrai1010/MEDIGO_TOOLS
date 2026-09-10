# 💰 HƯỚNG DẪN SỬ DỤNG TOOL TỰ ĐỘNG HÓA BẢNG LƯƠNG & CHẤM CÔNG MEDIGO

Hệ thống chuyên biệt bóc tách dữ liệu máy chấm công vân tay, tự động tính toán 16 sheet thành phần (Giờ công, Ngày công, MiniKAT, Cận date, WhatsApp, KPI, Thưởng CK, Giảm trừ) và hoàn thiện Sheet trung tâm **BẢNG LƯƠNG** với 100% công thức Excel chuẩn sạch.

---

## 1. ⚙️ YÊU CẦU CÀI ĐẶT & MÔI TRƯỜNG

### A. Cài đặt thư viện Python
Mở Terminal / Command Prompt và chạy:
```bash
pip install pandas openpyxl
```

*(Tùy chọn nâng cao nếu dùng uv):*
```bash
uv run --with pandas --with openpyxl python app.py
```

### B. Tự động tính toán số liệu Excel COM (Tùy chọn)
- File `recalc_workbook.ps1` sử dụng PowerShell và ứng dụng Microsoft Excel trên Windows để tự động mở, tính toán toàn bộ công thức và lưu sẵn giá trị số thực tế vào file.
- Không cần cài thêm gì nếu máy đã có Microsoft Excel.

---

## 2. 🚀 CÁCH KHỞI CHẠY HỆ THỐNG

### Cách 1: Chạy qua Giao diện Web (Khuyên dùng)
- Click đúp vào file: **`CHAY_TOOL_BANGLUONG.bat`**
- Trình duyệt sẽ tự động mở tại địa chỉ: `http://localhost:8768`
- Kéo thả file Chấm công, file KPI, file Hóa đơn và bấm **"TÍNH TOÁN & XUẤT BẢNG LƯƠNG"**.

### Cách 2: Chạy dòng lệnh (CLI / Terminal)
```bash
python build_payroll_report.py --month 8 --timecard "path/to/BangChiTietChamCong_thang8.xlsx" --kpi "path/to/baocaokpi_thang8_hoanthien.xlsx" --invoice "path/to/DanhSachChiTietHoaDon_3182026.xlsx"
```

---

## 3. 🧩 CHI TIẾT TỪNG MODULE FILE `.PY` (THEO TỪNG SHEET)

Mỗi sheet và nghiệp vụ tính toán trong bảng lương được bóc tách thành **1 file `.py` riêng biệt**:

| Tên File Module | Sheet Phụ Trách | Chức Năng & Quy Tắc Tính Toán |
| :--- | :--- | :--- |
| **`payroll_timecards.py`** | `Giờ công`<br>`Ngày công` | **Bóc tách máy chấm công vân tay:**<br>• Đọc giờ vào/ra từ `BangChiTietChamCong_thangX.xlsx`.<br>• Ca > 4h tính 1 công; Ca <= 4h chỉ cộng dồn số giờ.<br>• Phân định ca ngày vs ca đêm (tối qua đêm 22:00-06:00). |
| **`payroll_minikat.py`** | `MiniKat - HN`<br>`MiniKat - HCM` | **Tính thưởng dự án MiniKAT:**<br>• Quét hóa đơn & trừ trả hàng 3 nhóm SKUs: *PartySmart, KAT, LadyCare*.<br>• Tính thưởng từng dược sĩ (10k - 20k/sp).<br>• Tính thưởng đại diện Cửa Hàng Trưởng theo chi nhánh. |
| **`payroll_candate.py`** | `Cận date` | **Tính thưởng hàng cận date:**<br>• Quét các dòng sản phẩm có HSD <= 6 tháng so với ngày bán.<br>• Tính **5% doanh số cận date** cho từng dược sĩ theo chi nhánh. |
| **`payroll_project_reward.py`**| `Dự án` | **Thưởng dự án nhóm hàng & Hot Bill:**<br>• Đổ thưởng nhóm hàng (CK, Combo...) từ file KPI.<br>• Đổ thưởng chương trình Hot Bill Hà Nội (15-31/8). |
| **`payroll_whatsapp.py`** | `whatsapp` | **Hoa hồng đơn hàng WhatsApp:**<br>• Quét hóa đơn có ghi chú chứa từ khóa `whatsapp`.<br>• Tính tỷ lệ chi nhánh (>= 280M: 6%, >= 200M: 4%, >= 150M: 3%, < 150M: 1.5%).<br>• Tính hoa hồng cho từng dược sĩ bán. |
| **`payroll_kpi.py`** | `KPI` | **Thưởng KPI Doanh số:**<br>• Đổ Doanh thu mục tiêu, Doanh thu thực tế, % Hoàn thành.<br>• Áp dụng công thức tính thưởng KPI theo bậc đạt. |
| **`payroll_ck_points.py`** | `Thưởng CK` | **Thưởng Hàng điểm & Cắt liều:**<br>• Cập nhật thưởng hàng điểm 50,000 đ/hộp cho nhân sự HN.<br>• Áp dụng hệ số đánh giá hoa hồng (0.8 - 1.0). |
| **`payroll_deductions.py`** | `KPI trừ` | **Các khoản giảm trừ & cấn trừ:**<br>• Hàng hết date, vi phạm quy chế, hóa đơn App-Kiot, Ecom.<br>• Cấn trừ chi phí đồng phục (ví dụ: 500,000đ). |
| **`payroll_maps.py`** | `Map` | **Thưởng chương trình Google Maps:**<br>• Đổ số lượng đánh giá Maps và mức thưởng theo nhân sự & chi nhánh. |
| **`payroll_main_summary.py`**| `BẢNG LƯƠNG` | **Hoàn thiện Sheet Trung Tâm BẢNG LƯƠNG:**<br>• Gắn 100% công thức Excel sạch (Cột E đến AP).<br>• Tự động tính tăng ca OT (16h nếu >= 30 công, 8h nếu = 29 công).<br>• Tính phụ cấp ca đêm tối đa 1.5M (chuẩn 28 công).<br>• Tạo dòng Tổng cộng (Row 74) & Bảng tổng hợp theo 11 Chi nhánh (Row 77-87). |
| **`build_payroll_report.py`** | *Toàn bộ* | **File Tổng Hợp (Orchestrator):** Điều phối chạy toàn bộ 10 module trên theo đúng thứ tự, chuẩn hóa công thức OpenXML và gọi Excel COM tính toán. |
| **`app.py`** & **`index.html`** | *Giao diện Web* | Máy chủ backend (cổng 8766) và giao diện kéo thả file, xem kết quả trực quan. |
| **`CHAY_TOOL_BANGLUONG.bat`** | *Khởi chạy* | Script Windows click chạy nhanh Web App Bảng Lương. |

---

## 4. 📥 CÁC FILE ĐẦU VÀO (INPUT) CẦN THIẾT

1. **File Chấm Công Chi Tiết (`BangChiTietChamCong_thangX.xlsx`):** *(Bắt buộc)*
   - Xuất từ máy chấm công vân tay.
2. **File Báo Cáo KPI (`baocaokpi_thangX_hoanthien.xlsx`):** *(Bắt buộc)*
   - File kết quả xuất ra từ **`TOOL_KPI`**.
3. **File Hóa Đơn & Trả Hàng KiotViet:** *(Tùy chọn/Khuyên dùng)*
   - Dùng để tính toán chính xác MiniKAT, WhatsApp và Hàng cận date.
4. **Template Mẫu Bảng Lương:** *(Tùy chọn)*
   - Sử dụng template chuẩn có sẵn (`tinhcongnhungthuongchia.xlsx`).

---

## 5. 📤 KẾT QUẢ ĐẦU RA (OUTPUT)

File Excel hoàn thiện: **`BANGLUONGTHANGX_hoanthien.xlsx`** gồm 16 sheet hoàn chỉnh:
- Sheet **`BẢNG LƯƠNG`**: Bảng lương tổng hợp chi tiết theo từng nhân viên, từng chi nhánh, đầy đủ cột giờ ngày, giờ đêm, tăng ca OT, phụ cấp đêm, thưởng dự án, KPI, CK, WhatsApp, trừ phạt và thực lĩnh.
- Toàn bộ 15 sheet chi tiết đính kèm: `Giờ công`, `Ngày công`, `Tăng ca lễ`, `Công Đào tạo`, `MiniKat - HN`, `MiniKat - HCM`, `whatsapp`, `Dự án`, `KPI`, `Thưởng CK`, `Cận date`, `Map`, `KPI trừ`...
