# Walkthrough: Tự Động Hóa Bảng Lương & Chuẩn Hóa Chấm Công Toàn Bộ Nhân Sự (Tháng 8/2026)

## Cập Nhật Mới Nhất: Đồng Bộ Giờ Công & Ngày Công Chuẩn Cho Toàn Bộ 70 Dòng Nhân Sự

### 1. Bóc Tách Trực Tiếp Dữ Liệu Từ Máy Chấm Công Gốc
- **File nguồn**: `thang8/DATA/BangChiTietChamCong_thang8.xlsx`
- Đã xử lý **1,745 ca chấm công chi tiết** của tất cả nhân sự thuộc 11 chi nhánh toàn hệ thống.
- **Quy tắc tính ca chuẩn xác**:
  1. **Ca chỉ tính là 1 ca khi thời gian làm việc > 4 tiếng (> 4.0h)**.
  2. Các ca từ 1 đến 4 tiếng ($\le 4$h) không tính là 1 ca, nhưng toàn bộ số giờ làm việc vẫn được cộng đầy đủ vào tổng giờ công.
  3. Xóa bỏ hoàn toàn hiện tượng trùng lặp ca / nhân bản dòng từ file template cũ.

---

### 2. Cấu Trúc Sheet `Giờ công` & `Ngày công` Sau Khi Đồng Bộ

#### Sheet `Giờ công`:
- **Bảng chi tiết (Cột A:E)**: Chứa 230 nhóm ca thực tế của từng nhân viên theo chi nhánh, định dạng rõ ràng:
  - Cột A: Chi nhánh
  - Cột B: Tên nhân viên
  - Cột C: Tên ca
  - Cột D: Thời gian dạng chuỗi (`XhYp`)
  - Cột E: Số giờ công thực tế
- **Bảng tổng hợp (Cột G:K)**: Tính động bằng công thức:
  - Cột I (Giờ ca ngày): `=SUMIFS(E:E, B:B, H{r}, A:A, G{r}) - J{r}`
  - Cột J (Giờ ca đêm): `=SUMIFS(E:E, B:B, H{r}, C:C, "*đêm*", A:A, G{r})`
  - Cột K (Tổng giờ): `=I{r}+J{r}`

#### Sheet `Ngày công`:
- **Bảng chi tiết (Cột A:I)**: Chứa 1,745 dòng ca làm việc chi tiết từng ngày trong tháng 8:
  - Cột G (Ca $\le 4$h): `=IF(AND(ISNUMBER(F{r}), F{r}<=4), 1, 0)`
  - Cột H (Công thực tế): `=IF(AND(ISNUMBER(F{r}), F{r}>4, COUNTIFS($A$2:A{r}, A{r}, $B$2:B{r}, B{r}, $C$2:C{r}, C{r}, $F$2:F{r}, ">4")=1), 1, 0)`
  - Cột I (Ca đêm chuẩn): `=IF(AND(ISNUMBER(F{r}), F{r}>4, ISNUMBER(SEARCH("đêm", D{r}))), 1, 0)`
- **Bảng tổng hợp (Cột K:O)**:
  - Cột M (Ngày công chuẩn): `=SUMIFS(H:H, A:A, K{r}, B:B, L{r})`
  - Cột N (Giờ công): `='Giờ công'!K{r}`
  - Cột O (Trung bình giờ/công): `=IF(M{r}>0, N{r}/M{r}, "")`

---

### 3. Chuẩn Hóa 70 Dòng Trong Sheet `BẢNG LƯƠNG`
- **Cột G (Số giờ ca ngày)**: `=SUMIFS('Giờ công'!I:I, 'Giờ công'!G:G, B{r}, 'Giờ công'!H:H, C{r})`
- **Cột H (Số giờ ca đêm)**: `=SUMIFS('Giờ công'!J:J, 'Giờ công'!G:G, B{r}, 'Giờ công'!H:H, C{r})`
- **Cột I (Chuyên cần)**: `=IF(ROW()=MATCH(C{r}, $C$1:$C$72, 0), IF(SUMIF($C$3:$C$72, C{r}, $R$3:$R$72)>=30, 16, IF(SUMIF($C$3:$C$72, C{r}, $R$3:$R$72)>=29, 8, "")), "")`
- **Cột J (Số ngày làm ngày)**: `=ROUND(G{r}/24, 1)` (quy đổi 24h = 1 ngày, 8h = 0.3 ngày)
- **Cột K (Số ca làm đêm)**: `=SUMIFS('Ngày công'!I:I, 'Ngày công'!A:A, B{r}, 'Ngày công'!B:B, C{r})`
- **Cột R (Số ngày công thực tế)**: `=SUMIFS('Ngày công'!M:M, 'Ngày công'!K:K, B{r}, 'Ngày công'!L:L, C{r})`
- **Cột V (Phụ cấp CHT)**: `=IF(AND(ROW()=MATCH(C{r}, $C$1:$C$72, 0), OR(TRIM(D{r})="CHT", TRIM(D{r})="Q.CHT")), 500000/31*SUMIF($C$3:$C$72, C{r}, $R$3:$R$72), 0)`
- **Cột X (Phụ cấp ca đêm)**: `=IF(ISNUMBER(SEARCH("DSCD", D{r})), MIN(1500000, ROUND(1500000/28*K{r}, 0)), IF(K{r}>20, MIN(1500000, ROUND(1500000/28*K{r}, 0)), 0))`
- **Cột AA (Thưởng MiniKat DS)**: `=IF(ROW()=MATCH(C{r}, $C$1:$C$72, 0), IFERROR(VLOOKUP(C{r}, 'MiniKat - HN'!$A:$L, 12, FALSE), 0) + IFERROR(VLOOKUP(C{r}, 'MiniKat - HCM'!$A:$J, 10, FALSE), 0), 0)`
- **Cột AB (Thưởng MiniKat CHT)**: `=IF(OR(TRIM(D{r})="CHT", TRIM(D{r})="Q.CHT"), SUMIFS('MiniKat - HN'!$Y:$Y, 'MiniKat - HN'!$N:$N, B{r}) + SUMIFS('MiniKat - HCM'!$U:$U, 'MiniKat - HCM'!$M:$M, B{r}), 0)`
- **Cột AC (Thưởng WhatsApp)**: `=IF(ROW()=MATCH(C{r}, $C$1:$C$72, 0), SUMIFS(whatsapp!$D:$D, whatsapp!$A:$A, C{r}), 0)`

---

### 4. Bảng Kết Quả Đối Soát Một Số Nhân Sự Tiêu Biểu

| STT | Chi nhánh | Tên nhân viên | Giờ ca ngày (G) | Giờ ca đêm (H) | Ngày làm ngày (J) | Ca đêm chuẩn >4h (K) | Ngày công thực tế (R) |
|---|---|---|---|---|---|---|---|
| **64** | **Đường Láng** | **Nguyễn Thị Tâm** | 112.18h | 176.26h | **4.7** ngày | **22** ca | **28** ngày |
| **61** | **Đường Láng** | **Hứa Thị Kim Thoa** | 165.92h | 47.16h | **6.9** ngày | **6** ca | **23** ngày |
| **62** | **Đường Láng** | **Lê Thị Soạn** | 0.00h | 247.77h | **0.0** ngày | **31** ca | **31** ngày |
| **53** | **Lê Bình** | **Hồ Thị Minh Hòa** | 272.06h | 6.88h | **11.3** ngày | **1** ca | **28** ngày |
| **1** | **Trường Sa** | **Hồ Thị Minh Hòa** | 15.69h | 0.00h | **0.7** ngày | **0** ca | **1** ngày |
| **11** | **Đỗ Quang Đẩu** | **Trần Thiên Phát** | 0.00h | 245.44h | **0.0** ngày | **31** ca | **31** ngày |
| **55** | **Lê Bình** | **Lê Thị Xuyến** | 16.39h | 243.48h | **0.7** ngày | **30** ca | **30** ngày |
| **--** | **TỔNG CỘNG** | **Toàn hệ thống** | **7,745.69h** | **5,123.39h** | **322.8** ngày | **628** ca | **1,464** ngày |

---

## File Kết Quả Đã Xuất
1. `d:\MEDIGO\KPI_UPDATE\thang8\bangluong_thang8_hoanthien.xlsx`
2. `d:\MEDIGO\KPI_UPDATE\thang8\BANGLUONGTHANG8.xlsx` (Đã đồng bộ hoàn chỉnh)
3. `d:\MEDIGO\KPI_UPDATE\baocaokpi_thang8_hoanthien.xlsx` (Đã đồng bộ toàn bộ bảng màu chuẩn nhà thuốc)
4. `d:\MEDIGO\KPI_UPDATE\TOOL_KPISHEET\output\NHÀ THUỐC THÁNG 9 2026.xlsx` (Gói kế hoạch KPI tháng 9 chuẩn 5 sheet)

---

## Nâng Cấp Giao Diện Báo Cáo KPI Chuẩn Nhà Thuốc (Đẹp, Trực Quan, Chuẩn Màu)
- **Tích hợp module định dạng tự động `kpi_styling.py`** vào cả 2 engine: `TOOL_KPISHEET` (khởi tạo gói kế hoạch) và `TOOL_KPI` (tính toán báo cáo KPI thực tế).
- **Phối màu nhận diện phân vùng chuẩn**:
  - **Navy Blue (`#1F4E79`)**: Cột thông tin nhân sự và chi nhánh (chữ trắng đậm).
  - **Soft Blue (`#DDEBF7`)**: Doanh thu offline, chỉ tiêu & trung bình bill.
  - **Soft Purple (`#E8D8F8`)**: Sản phẩm dự án, combo, chiết khấu.
  - **Soft Amber / Yellow (`#FFF2CC` & `#FFFBEA`)**: Mốc KPI, Doanh thu/ngày, `% Hoàn thành KPI`.
  - **Emerald Green (`#C6EFCE` & `#EBF9F1`)**: Thưởng KPI, Thưởng CHT, Thưởng dự án (chữ xanh đậm `#15803D`).
  - **Peach / Coral (`#FCE4D6`)**: Doanh thu online và nhóm hàng điểm.
  - **Zebra striping (`#FAFAFA`)** + Viền mảnh + Tự động căn lề (Left/Center/Right) + Tự động định dạng số (`#,##0`, `0.0%`).

---

## Đối Soát Dữ Liệu Đầu Vào vs Đầu Ra (Data Integrity Verification)
- **Smart Dynamic Header Mapping**: Tự động quét dòng tiêu đề (Header row detection) và nhận diện thông minh tên các cột thay vì fix cứng vị trí cột (`Nhà thuốc`, `Nhân viên`, `Chức danh`, `Trung bình bill`, `KPIs Doanh thu/tháng`, `NT tháng`).
- **Kết quả đối soát 100%**:
  - **Tháng 9**: Khớp chính xác 100% toàn bộ **51 nhân sự** và **11 nhà thuốc** từ file đề xuất HCM (`KPI CNT HCM Tháng 09.xlsx`) và HN (`e_xuat_KPI_Quy_3.26.xlsx`).
  - **Tháng 8**: Khớp chính xác 100% toàn bộ **52 nhân sự** và **11 nhà thuốc** trên toàn hệ thống.
  - Toàn bộ công thức Excel liên kết đa tầng (`VLOOKUP`, `SUMIF`, `CEILING`, `DAY`) khớp tuyệt đối với file gốc nhà thuốc.
