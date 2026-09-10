# 📚 TỔNG HỢP CÔNG THỨC GỐC CHUẨN THÁNG 7/2026: BẢNG LƯƠNG & BÁO CÁO KPI

> Đây là bộ công thức đối sánh chuẩn mực (Baseline Benchmark) được trích xuất 100% từ file Gốc Tháng 7 (BẢNG LƯƠNG THÁNG 7 2026.xlsx và NHÀ THUỐC THÁNG 7 2026.xlsx). Bộ công thức này là căn cứ gốc để thiết lập kế hoạch và chuyển giao sang Tháng 8, Tháng 9 khi có chương trình/dự án mới.

## 🏛️ PHẦN 1: TOÀN BỘ 16 SHEET BẢNG LƯƠNG THÁNG 7 (GỐC CHUẨN)

### 📌 SHEET: Bảng đánh giá (999 dòng × 28 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | 24H |
| **B** (2) | Doanh thu | Trống / Tùy chọn | None |
| **C** (3) | Chi phí | Trống / Tùy chọn | None |
| **D** (4) | Lợi nhuận | **Công thức** | =B2-C2 |
| **E** (5) | Cột E | **Công thức** | =D2/B2 |
| **F** (6) | Thành tích | Dữ liệu nhập | Không đạt |
| **G** (7) | Lương | Dữ liệu nhập | 72000000 |
| **H** (8) | Mặt bằng | Dữ liệu nhập | 17000000 |
| **I** (9) | Ds chuyên môn | Dữ liệu nhập | 3500000 |
| **J** (10) | Điện + nước | Dữ liệu nhập | 3400000 |
| **K** (11) | Linh tinh | Dữ liệu nhập | 6000000 |
| **L** (12) | SL đơn ONl | Dữ liệu nhập | 1238 |
| **M** (13) | Platform software (8k/1hd) | **Công thức** | =L2*8000 |
| **N** (14) | Ebitda | **Công thức** | =D2-G2-H2-I2-J2-K2-M2 |
| **O** (15) | Ebitda - Team CNT (160M) | **Công thức** | =N2-16000000 |
| **P** (16) | Ebitda - Comp (300M) | **Công thức** | =O2-30000000 |

---

### 📌 SHEET: BẢNG LƯƠNG (1043 dòng × 53 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Cột A | **Công thức** | =SEQUENCE(COUNTA(B3:B111),1) |
| **B** (2) | Cột B | Dữ liệu nhập | 24H |
| **C** (3) | Cột C | Dữ liệu nhập | Hồ Thị Minh Hòa |
| **D** (4) | Cột D | Dữ liệu nhập | DSBC |
| **E** (5) | Tổng thời gian làm việc | **Công thức** | =G3+I3+L3+O3+S3+T3 |
| **F** (6) | Cột F | **Công thức** | =H3+P3 |
| **G** (7) | Quản lý và cửa hàng trưởng ghi nhận vào đây - công thức tổng tự động không cần chỉnh | **Công thức** | =SUMIFS('Giờ công'!I:I,'Giờ công'!G:G,B3,'Giờ công'!H:H,C3) |
| **H** (8) | Cột H | **Công thức** | =SUMIFS('Giờ công'!J:J,'Giờ công'!G:G,B3,'Giờ công'!H:H,C3) |
| **I** (9) | Cột I | **Công thức** | =ifs(R3>=30,16,R3=29,8,true,"") |
| **J** (10) | Cột J | Dữ liệu nhập | 18 |
| **K** (11) | Cột K | Dữ liệu nhập | 6 |
| **N** (14) | Cột N | **Công thức** | =SUMIFS('Tăng ca lễ'!M:M,'Tăng ca lễ'!K:K,B3,'Tăng ca lễ'!L:L,C3) |
| **O** (15) | Cột O | **Công thức** | =IF(or(D3 = "PT",D3 = "DSHV"), SUMIFS('Tăng ca lễ'!O:O,'Tăng ca lễ'!K:K,B3,'Tăng ca lễ'!L:L,C3)*2, SUMIFS('Tăng ca lễ'!O:O,'Tăng ca lễ'!K:K,B3,'Tăng ca lễ'!L:L,C3)*3) |
| **P** (16) | Cột P | **Công thức** | =IF(or(D3 = "PT", D3 = "DSHV"), SUMIFS('Tăng ca lễ'!P:P,'Tăng ca lễ'!K:K,B3,'Tăng ca lễ'!L:L,C3)*2, SUMIFS('Tăng ca lễ'!P:P,'Tăng ca lễ'!K:K,B3,'Tăng ca lễ'!L:L,C3)*3) |
| **R** (18) | Quản lý ghi nhận | **Công thức** | =SUMIFS('Ngày công'!M:M,'Ngày công'!K:K,B3,'Ngày công'!L:L,C3) |
| **S** (19) | Cột S | **Công thức** | =SUMIFS('Công Đào tạo'!G:G,'Công Đào tạo'!E:E,B3,'Công Đào tạo'!F:F,C3) |
| **T** (20) | Cột T | Dữ liệu nhập | 0 |
| **V** (22) | Cột V | **Công thức** | =500000/30*17 |
| **X** (24) | Cột X | **Công thức** | =IFERROR(__xludf.DUMMYFUNCTION("IF($C3="""","""",  IF(   TRIM($D3)=""DSCD"",   MIN(    1500000,    ROUND(1500000/28*MIN(N($J3),28),0)   ),   IF(    AND(     REGEXMATCH(TRIM($D3),""^(DSXC\|DSBC\|CHT\|DSTV)$""),     N($H3)>=160    ),    MIN(     1500000,     ROUND(1500000/28*MIN(N($K3),28),0"&")    ),    0   )  ) )"),0.0) |
| **AA** (27) | Cột AA | **Công thức** | =IF($C3="","",  IFERROR(   LET(    cnChinh,    _xlfn.XLOOKUP(     TRIM($C3),     TRIM(KPI!$B$2:$B1043),     KPI!$A$2:$A1043,     ""    ),    IF(     TRIM($B3)=TRIM(cnChinh),     _xlfn.XLOOKUP(      TRIM($C3),      TRIM('MiniKat - HCM'!$A$2:$A1043),      'MiniKat - HCM'!$J$2:$J1043,      ""     ),     ""    )   ),   ""  ) ) |
| **AC** (29) | Cột AC | **Công thức** | =IFERROR(_xlfn.XLOOKUP(C3, whatsapp!A:A,whatsapp!D:D), " ") |
| **AD** (30) | Cột AD | **Công thức** | =IF($C3="","",  IFERROR(   LET(    cnChinh,    _xlfn.XLOOKUP(     TRIM($C3),     TRIM(KPI!$B$2:$B1043),     KPI!$A$2:$A1043,     ""    ),    IF(     TRIM($B3)=TRIM(cnChinh),     _xlfn.XLOOKUP(      TRIM($C3),      TRIM('Dự án'!$B$2:$B1043),      'Dự án'!$C$2:$C1043,      ""     ),     ""    )   ),   ""  ) ) |
| **AE** (31) | Cột AE | **Công thức** | =SUMIFS(KPI!C:C,KPI!A:A,B3,KPI!B:B,C3) |
| **AF** (32) | Cột AF | **Công thức** | =SUMIFS('Thưởng CK'!C:C,'Thưởng CK'!A:A,B3,'Thưởng CK'!B:B,C3) |
| **AG** (33) | Cột AG | **Công thức** | =SUMIFS('Cận date'!C:C,'Cận date'!A:A,B3,'Cận date'!B:B,C3) |
| **AH** (34) | Cột AH | **Công thức** | =SUMIFS(Map!D:D,Map!B:B,B3,Map!A:A,C3) |
| **AJ** (36) | Cột AJ | **Công thức** | =SUM(AK3:AO3) |
| **AK** (37) | KPI trừ | **Công thức** | =SUMIFS('KPI trừ'!C:C,'KPI trừ'!A:A,B3,'KPI trừ'!B:B,C3) |
| **AL** (38) | Cột AL | **Công thức** | =SUMIFS('KPI trừ'!D:D,'KPI trừ'!A:A,B3,'KPI trừ'!B:B,C3) |
| **AM** (39) | Cột AM | **Công thức** | =SUMIFS('KPI trừ'!E:E,'KPI trừ'!A:A,B3,'KPI trừ'!B:B,C3) |
| **AN** (40) | Cột AN | **Công thức** | =SUMIFS('KPI trừ'!F:F,'KPI trừ'!A:A,B3,'KPI trừ'!B:B,C3) |
| **AO** (41) | Cột AO | **Công thức** | =SUMIFS('KPI trừ'!G:G,'KPI trừ'!A:A,B3,'KPI trừ'!B:B,C3) |
| **AP** (42) | Cột AP | **Công thức** | =IFERROR(__xludf.DUMMYFUNCTION("IFERROR(TEXTJOIN("" "", TRUE, FILTER('KPI trừ'!H:H, ('KPI trừ'!A:A=B3) * ('KPI trừ'!B:B=C3))), """")"),"") |
| **AQ** (43) | Cột AQ | **Công thức** | =IFERROR(__xludf.DUMMYFUNCTION("IFERROR(  MAX(   QUERY(    FILTER(     {'Ngày công'!$C$2:$C$5000,'Ngày công'!$F$2:$F$5000},     TRIM('Ngày công'!$B$2:$B$5000)=TRIM($C3)    ),    ""select sum(Col2)     where Col1 is not null     group by Col1     label sum(Col2) ''"",    0   )  ),  0 )"),16.65) |

---

### 📌 SHEET: MiniKat - HN (1000 dòng × 21 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Người bán | Dữ liệu nhập | Đinh Thị Khánh Ly |
| **B** (2) | SL Party Smart | Dữ liệu nhập | 0 |
| **C** (3) | SL KAT bán | Dữ liệu nhập | 4 |
| **D** (4) | SL LadyCare bán | Dữ liệu nhập | 5 |
| **E** (5) | Doanh thu LadyCare | Dữ liệu nhập | 1045000 |
| **F** (6) | Doanh Thu KAT | Dữ liệu nhập | 1800000 |
| **G** (7) | Thưởng KAT | Dữ liệu nhập | 0 |
| **H** (8) | Thưởng Party | Dữ liệu nhập | 0 |
| **I** (9) | Thưởng LadyCare | Dữ liệu nhập | 0 |
| **J** (10) | Tổng Thưởng | Dữ liệu nhập | 0 |
| **M** (13) | Chi nhánh | Dữ liệu nhập | NT Đường Láng 247 |
| **N** (14) | SL KAT bán | Dữ liệu nhập | 3 |
| **O** (15) | SL PARTY SMART bán | Dữ liệu nhập | 16 |
| **P** (16) | Thưởng KAT Tổng | Dữ liệu nhập | 0 |
| **Q** (17) | Thưởng PS 180 hộp | Dữ liệu nhập | 0 |
| **R** (18) | PS đầu 150 hộp | Dữ liệu nhập | 0 |
| **S** (19) | Doanh thu LadyCare | Dữ liệu nhập | 4912000 |
| **T** (20) | Thưởng LadyCare | Dữ liệu nhập | 0 |
| **U** (21) | Tổng thưởng | Dữ liệu nhập | 0 |

---

### 📌 SHEET: MiniKat - HCM (1000 dòng × 21 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Người bán | Dữ liệu nhập | Hồ Thị Minh Hòa |
| **B** (2) | SL Party Smart | Dữ liệu nhập | 14 |
| **C** (3) | SL KAT bán | Dữ liệu nhập | 2 |
| **D** (4) | SL LadyCare bán | Dữ liệu nhập | 2 |
| **E** (5) | Doanh thu LadyCare | Dữ liệu nhập | 338000 |
| **F** (6) | Doanh Thu KAT | Dữ liệu nhập | 960000 |
| **G** (7) | Thưởng KAT | Dữ liệu nhập | 0 |
| **H** (8) | Thưởng Party | Dữ liệu nhập | 0 |
| **I** (9) | Thưởng LadyCare | Dữ liệu nhập | 0 |
| **J** (10) | Tổng Thưởng | Dữ liệu nhập | 0 |
| **M** (13) | Chi nhánh | Dữ liệu nhập | NT 24H |
| **N** (14) | SL KAT bán | Dữ liệu nhập | 12 |
| **O** (15) | SL PARTY SMART bán | Dữ liệu nhập | 80 |
| **P** (16) | Thưởng KAT Tổng | Dữ liệu nhập | 0 |
| **Q** (17) | Thưởng PS 180 hộp | Dữ liệu nhập | 0 |
| **R** (18) | PS đầu 150 hộp | Dữ liệu nhập | 0 |
| **S** (19) | Doanh thu LadyCare | Dữ liệu nhập | 4332000 |
| **T** (20) | Thưởng LadyCare | Dữ liệu nhập | 0 |
| **U** (21) | Tổng thưởng | Dữ liệu nhập | 0 |

---

### 📌 SHEET: whatsapp (979 dòng × 25 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Whatsapp - Người bán | Dữ liệu nhập | Đinh Thị Khánh Ly |
| **B** (2) | Doanh thu | Dữ liệu nhập | 19255000 |
| **C** (3) | Hệ số thưởng | Dữ liệu nhập | 0.015 |
| **D** (4) | Thưởng | **Công thức** | =B2*C2 |
| **F** (6) | Ứng dụng nhắn tin Whatsapp | Trống / Tùy chọn | None |

---

### 📌 SHEET: Cận date (927 dòng × 16 cột)

> 💡 **Quy tắc bóc tách Cận Date từ Hóa đơn**: Sản phẩm được xác định là cận date khi có Hạn sử dụng (HSD) cách thời điểm đơn bán trong vòng **6 tháng** ($0 \le \text{HSD} - \text{Ngày Bán} \le 183\text{ ngày}$) hoặc sản phẩm có gắn tiền tố `CD` / `CẬN`. Thưởng cận date được tính bằng **5%** trên tổng doanh thu cận date bán được (`=I2*5%`).

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | 24H |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **C** (3) | Thưởng cận Date | **Công thức** | =_xlfn.XLOOKUP(B2,H:H,J:J,"") |
| **H** (8) | Tên nhân viên | Dữ liệu nhập | Bùi Thị Thanh Thủy |
| **I** (9) | Doanh thu cận date | Dữ liệu nhập | 20750 |
| **J** (10) | Thưởng cận date | **Công thức** | =I2*5% |

---

### 📌 SHEET: KPI (994 dòng × 27 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | 24H |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **C** (3) | Thưởng KPI | **Công thức** | =_xlfn.XLOOKUP(B2,H:H,L:L,0) |
| **H** (8) | Tên nhân viên | Dữ liệu nhập | Lê Ngọc Anh |
| **I** (9) | KPI | Dữ liệu nhập | 178000000 |
| **J** (10) | Doanh thu thực | Dữ liệu nhập | 178400 |
| **K** (11) | % KPI | Dữ liệu nhập | 0.82 |
| **L** (12) | Thưởng KPI | Dữ liệu nhập | 4128808 |

---

### 📌 SHEET: Thưởng CK (992 dòng × 22 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | 24H |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **C** (3) | Total thưởng | **Công thức** | =E2*D2 |
| **D** (4) | Hệ số thưởng | **Công thức** | =_xlfn.XLOOKUP(B2,H:H,L:L,1) |
| **E** (5) | Thưởng CK - Nghĩa điền | Dữ liệu nhập | 876500 |
| **H** (8) | Tên nhân viên | Dữ liệu nhập | Lê Ngọc Anh |
| **I** (9) | Total thưởng | **Công thức** | =J2*K2 |
| **J** (10) | Hệ số | Dữ liệu nhập | 1 |
| **K** (11) | Dự án | Dữ liệu nhập | 0 |
| **L** (12) | Hệ số thưởng HH | **Công thức** | =IF(AND(J2=0.8,K2=0),0.9,1) |

---

### 📌 SHEET: Dự án (991 dòng × 16 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | 24H |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **C** (3) | Dự án CK - 2M | **Công thức** | =_xlfn.XLOOKUP(B2,H:H,I:I,"") |
| **H** (8) | Nhân viên | Dữ liệu nhập | Lê Ngọc Anh |
| **I** (9) | Total thưởng | **Công thức** | =J2*K2*L2 |
| **J** (10) | Hệ số | Dữ liệu nhập | 1 |
| **K** (11) | Dự án | Dữ liệu nhập | 0 |
| **L** (12) | Hệ số App | Dữ liệu nhập | 1 |

---

### 📌 SHEET: Giờ công (2985 dòng × 27 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | Hàng Bông |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Nguyễn Mạnh Tuấn |
| **C** (3) | Tên ca | Dữ liệu nhập | Ca Sáng |
| **D** (4) | Mỗi ca | Dữ liệu nhập | 2h0p |
| **E** (5) | Quy đổi | **Công thức** | =IF(ISNUMBER(SEARCH("h",D2)),      LEFT(D2,SEARCH("h",D2)-1) + IFERROR(MID(D2,SEARCH("h",D2)+1,LEN(D2)-SEARCH("h",D2)-1)/60, 0),      IF(ISNUMBER(SEARCH("p",D2)), LEFT(D2,SEARCH("p",D2)-1)/60, "") ) |
| **G** (7) | Chi nhánh | Dữ liệu nhập | 24H |
| **H** (8) | Tên nhân viên | Dữ liệu nhập | Hồ Thị Minh Hòa |
| **I** (9) | Sáng | **Công thức** | =SUMIFS(E:E,B:B,H2,A:A,G2) - J2 |
| **J** (10) | Đêm | **Công thức** | =SUMIFS(E:E, B:B, H2 , C:C, "*đêm*",A:A,G2) |
| **K** (11) | Tổng | **Công thức** | =I2+J2 |
| **N** (14) | Tên chi nhánh | Dữ liệu nhập | Hàng Bông |
| **O** (15) | Số giờ theo chi nhánh | **Công thức** | =SUMIF(A:A,N2,E:E) |

---

### 📌 SHEET: Ngày công (2677 dòng × 30 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | Hàng Bông |
| **B** (2) | Tên Nhân viên | Dữ liệu nhập | Nguyễn Mạnh Tuấn |
| **C** (3) | Ngày | Dữ liệu nhập | 2026-07-02 00:00:00 |
| **D** (4) | Tên ca | Dữ liệu nhập | Ca đêm ( 2 )(22:00 - 06:00)(ca qua đêm) |
| **E** (5) | Giờ làm thực tế | Dữ liệu nhập | 7h58p |
| **F** (6) | Giờ công chuẩn | **Công thức** | =IF(ISNUMBER(SEARCH("h",E2)),      LEFT(E2,SEARCH("h",E2)-1) + IFERROR(MID(E2,SEARCH("h",E2)+1,LEN(E2)-SEARCH("h",E2)-1)/60, 0),      IF(ISNUMBER(SEARCH("p",E2)), LEFT(E2,SEARCH("p",E2)-1)/60, "") ) |
| **G** (7) | Trừ công | **Công thức** | =IF(AND(ISNUMBER(F2), F2<4), 1, 0) |
| **H** (8) | Trùng trừ công | **Công thức** | =IF(AND(G2=1, COUNTIFS(A:A, A2, B:B, B2, C:C, C2) > 1), 1, 0) |
| **K** (11) | Chi nhánh | Dữ liệu nhập | 24H |
| **L** (12) | Tên nhân viên | Dữ liệu nhập | Hồ Thị Minh Hòa |
| **M** (13) | Ngày công chuẩn | **Công thức** | =IFERROR(__xludf.DUMMYFUNCTION("COUNTUNIQUEIFS(C:C,A:A,K2,B:B,L2)-SUMIFS(G:G,A:A,K2,B:B,L2)+SUMIFS(H:H,A:A,K2,B:B,L2)"),18.0) |
| **N** (14) | Giờ công | **Công thức** | ='Giờ công'!K2 |
| **O** (15) | Trung bình giờ/công | **Công thức** | =IF(M2>0,N2/M2,"") |
| **Q** (17) | Cột Q | **Công thức** | =UNIQUE(A2:A5668) |
| **R** (18) | sheet ngày công | **Công thức** | =SUMIF(A:A,Q2,F:F) |
| **S** (19) | Sheet giờ công | **Công thức** | =SUMIF(K:K,Q2,N:N) |
| **T** (20) | Cột T | **Công thức** | =R2-S2 |

---

### 📌 SHEET: KPI trừ (999 dòng × 22 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | 24H |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **C** (3) | Hàng hết date | Trống / Tùy chọn | None |
| **D** (4) | Quy chế quy trình | Trống / Tùy chọn | None |
| **E** (5) | HD App - Kiot | Trống / Tùy chọn | None |
| **F** (6) | Ecom | Trống / Tùy chọn | None |
| **G** (7) | Khác | Dữ liệu nhập | 500000 |
| **H** (8) | Ghi chú | Dữ liệu nhập | Cấn trừ chi phí đồng phục |

---

### 📌 SHEET: Tăng ca lễ (2942 dòng × 32 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Trống / Tùy chọn | None |
| **B** (2) | Tên nhân viên | Trống / Tùy chọn | None |
| **C** (3) | Ngày | Trống / Tùy chọn | None |
| **D** (4) | Tên ca | Trống / Tùy chọn | None |
| **E** (5) | Giờ làm thực tế | Trống / Tùy chọn | None |
| **F** (6) | Quy đổi | **Công thức** | =IF(ISNUMBER(SEARCH("h",E2)),      LEFT(E2,SEARCH("h",E2)-1) + IFERROR(MID(E2,SEARCH("h",E2)+1,LEN(E2)-SEARCH("h",E2)-1)/60, 0),      IF(ISNUMBER(SEARCH("p",E2)), LEFT(E2,SEARCH("p",E2)-1)/60, "") ) |
| **G** (7) | Trừ công | **Công thức** | =IF(AND(ISNUMBER(F2), F2<2), 1, 0) |
| **H** (8) | Trùng trừ công | **Công thức** | =IF(AND(G2=1, COUNTIFS(A:A, A2, B:B, B2, C:C, C2) > 1), 1, 0) |
| **K** (11) | Chi nhánh | Dữ liệu nhập | 24H |
| **L** (12) | Tên nhân viên | Dữ liệu nhập | Hồ Thị Minh Hòa |
| **M** (13) | Ngày công chuẩn | **Công thức** | =IFERROR(__xludf.DUMMYFUNCTION("COUNTUNIQUEIFS(C:C,A:A,K2,B:B,L2)-SUMIFS(G:G,A:A,K2,B:B,L2)+SUMIFS(H:H,A:A,K2,B:B,L2)"),0.0) |
| **N** (14) | Tổng | **Công thức** | =O2+P2 |
| **O** (15) | Số giờ lễ ca ngày ( x 3 ) | **Công thức** | =SUMIFS(F:F,B:B,L2,A:A,K2) - P2 |
| **P** (16) | Số giờ lễ ca đêm ( x 3 ) | **Công thức** | =SUMIFS(F:F, B:B, L2 , D:D, "*đêm*",A:A,K2) |
| **Q** (17) | Ngày lễ vẫn đi làm (ghi ngày cụ thể) | Trống / Tùy chọn | None |

---

### 📌 SHEET: Công Đào tạo (1007 dòng × 24 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **B** (2) | Thời gian đào tạo | Dữ liệu nhập | 0 |
| **C** (3) | Quy đổi | **Công thức** | =B2/60 |
| **E** (5) | Chi nhánh | Dữ liệu nhập | 24H |
| **F** (6) | Tên nhân viên | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **G** (7) | Giờ đào tạo | **Công thức** | =SUMIF(A:A,F2,C:C) |

---

### 📌 SHEET: Map (1000 dòng × 20 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Tên nhân viên | Trống / Tùy chọn | None |
| **B** (2) | Chi nhánh | Trống / Tùy chọn | None |
| **C** (3) | SL đánh giá Tháng 06/26 | Trống / Tùy chọn | None |
| **D** (4) | Thưởng MAP | **Công thức** | =(C2+E2)*20000 |
| **G** (7) | Cột G | Dữ liệu nhập | không có gửi đánh giá maps |

---

### 📌 SHEET: Nghỉ phép năm - HR (998 dòng × 18 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Chi nhánh | Dữ liệu nhập | Hàng Bông |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Lê Thị Ngọc Ánh |
| **C** (3) | Số giờ nghỉ phép năm | Trống / Tùy chọn | None |
| **D** (4) | Ngày phép năm  (ghi ngày cụ thể) | Trống / Tùy chọn | None |

---

## 📊 PHẦN 2: TOÀN BỘ CÁC SHEET BÁO CÁO KPI THÁNG 7 (GỐC CHUẨN)

### 📌 SHEET: data (972 dòng × 128 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Nhà thuốc | Dữ liệu nhập | Trường Sa |
| **B** (2) | Tên nhân viên | Dữ liệu nhập | Hồ Thị Minh Hòa |
| **C** (3) | Số giao dịch off | Dữ liệu nhập | 318.0 |
| **D** (4) | Doanh thu off | Dữ liệu nhập | 21517000.0 |
| **E** (5) | Số giao dịch onl | Dữ liệu nhập | 221.0 |
| **F** (6) | Doanh thu onl | Dữ liệu nhập | 38107899.0 |
| **G** (7) | Tổng giao dịch | Dữ liệu nhập | 539.0 |
| **H** (8) | Doanh thu tổng | Dữ liệu nhập | 59624899.0 |
| **J** (10) | 2026-07-31 00:00:00 | Trống / Tùy chọn | None |

---

### 📌 SHEET: kpi dược sĩ (944 dòng × 38 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Dữ liệu cập nhật đến ngày | Dữ liệu nhập | Hàng Bông |
| **B** (2) | Cột B | Dữ liệu nhập | Lê Ngọc Anh |
| **C** (3) | =data!J1 | Dữ liệu nhập | CHT |
| **D** (4) | Cột D | Dữ liệu nhập | 150000.0 |
| **E** (5) | Cột E | **Công thức** | =AE3 |
| **F** (6) | Cột F | **Công thức** | =IFERROR(   ROUND(     LET(       nt,$A3,       nv,$B3,       songay,DAY(EOMONTH($C$1,0)),       doanhthungay,$I3,       duan_nv,SUMIFS('Dự án T7'!$D:$D,'Dự án T7'!$A:$A,nt,'Dự án T7'!$B:$B,nv),       duan_nt,SUMIFS('Dự án T7'!$D:$D,'Dự án T7'!$A:$A,nt),       tyle,duan_nv/duan_nt,       muc2_ngay,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$P:$P),       muc3_ngay,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$T:$T),       hangdiem1,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$I:$I),       hangdiem2,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$N:$N),       hangdiem3,_xlfn.XLOOKUP(nt,'kpi nhà thuốc'!$B:$B,'kpi nhà thuốc'!$R:$R),       SWITCH(TRUE,         doanhthungay>=muc3_ngay,hangdiem3,         doanhthungay>=muc2_ngay,hangdiem2,         TRUE,hangdiem1       )/songay*tyle     ),   -3), 0) |
| **G** (7) | Cột G | **Công thức** | =_xlfn.XLOOKUP(B3,'Dự án T7'!B:B,'Dự án T7'!D:D,0) |
| **H** (8) | Cột H | **Công thức** | =Q3 |
| **I** (9) | Cột I | **Công thức** | =(AD3/DAY($C$1)) |
| **J** (10) | Cột J | **Công thức** | =_xlfn.XLOOKUP(B3,'kpi nhà thuốc'!A:A,'kpi nhà thuốc'!F:F,(AD3/DAY($C$1))/U3) |
| **K** (11) | Cột K | **Công thức** | =_xlfn.XLOOKUP(B3,'kpi nhà thuốc'!A:A,'kpi nhà thuốc'!G:G,IF(AND(C3="BC",G3>=F3),     IFS(         AND(I3>=T3,E3>=D3), AD3*1.5%,         AND(I3>=T3,E3<D3), AD3*1.3%,         AND(I3>=S3,E3>=D3), AD3*1.3%,         AND(I3>=S3,E3<D3), AD3*1.1%,         AND(I3>=R3,E3>=D3), AD3*1.2%,         AND(I3>=R3,E3<D3), AD3*1%,         TRUE, ""     ),     "" )) |
| **L** (12) | Cột L | **Công thức** | =IF(C3="DSTV","", IF(J3<0.6, 0.8, 1)) |
| **M** (13) | Cột M | **Công thức** | =IFERROR(   _xlfn.XLOOKUP(     $A3&$B3,     'Dự án T7'!$A:$A&'Dự án T7'!$B:$B,     'Dự án T7'!$L:$L,     0   )*$L3, 0) |
| **N** (14) | Cột N | **Công thức** | =K3+M3*L3 |
| **O** (15) | Mức hoành thành KPI doanh thu/ngày | **Công thức** | =CEILING(R3,100000) |
| **P** (16) | Cột P | **Công thức** | =CEILING(S3,100000) |
| **Q** (17) | Cột Q | **Công thức** | =CEILING(T3,100000) |
| **R** (18) | Mức hoành thành KPI doanh thu/ngày | **Công thức** | =U3*$R$2 |
| **S** (19) | Cột S | **Công thức** | =U3*$S$2 |
| **T** (20) | Cột T | **Công thức** | =U3*$T$2 |
| **U** (21) | Cột U | **Công thức** | =V3/DAY('kpi nhà thuốc'!$T$1) |
| **V** (22) | Cột V | Dữ liệu nhập | 178000000.0 |
| **W** (23) | Cột W | **Công thức** | =SUMIF(data!$B:$B,B3,data!$C:$C) |
| **X** (24) | Cột X | **Công thức** | =SUMIF(data!$B:$B,B3,data!$D:$D) |
| **Y** (25) | Cột Y | **Công thức** | =IF(W3>0,X3/W3,0) |
| **Z** (26) | Cột Z | **Công thức** | =SUMIF(data!$B:$B,B3,data!$E:$E) |
| **AA** (27) | Cột AA | **Công thức** | =SUMIF(data!$B:$B,B3,data!$F:$F) |
| **AB** (28) | Cột AB | **Công thức** | =IF(Z3>0,AA3/Z3,0) |
| **AC** (29) | Cột AC | **Công thức** | =W3+Z3 |
| **AD** (30) | Cột AD | **Công thức** | =X3+AA3 |
| **AE** (31) | Cột AE | **Công thức** | =IF(AC3>0,AD3/AC3,0) |

---

### 📌 SHEET: kpi nhà thuốc (1000 dòng × 38 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Cột A | Dữ liệu nhập | Nguyễn Trần Ngọc Phương |
| **B** (2) | Cột B | Dữ liệu nhập | Trường Sa |
| **C** (3) | Cột C | Dữ liệu nhập | 640000000.0 |
| **D** (4) | Dữ liệu cập nhật đến ngày | **Công thức** | =AE3 |
| **E** (5) | Cột E | **Công thức** | =D3/DAY('kpi dược sĩ'!$C$1) |
| **F** (6) | ='kpi dược sĩ'!C1 | **Công thức** | =E3/T3 |
| **G** (7) | Cột G | **Công thức** | =IF(     W3<>"Đạt",     0,     IFS(         E3>=T3, D3*0.5%+1000000,         E3>=P3, D3*0.4%+700000,         E3>=K3, D3*0.35%+500000,         TRUE, 0     ) ) |
| **H** (8) | Cột H | Dữ liệu nhập | 0.8 |
| **I** (9) | Cột I | **Công thức** | =J3*0.25 |
| **J** (10) | Cột J | **Công thức** | =C3*H3 |
| **K** (11) | Cột K | **Công thức** | =T3*H3 |
| **L** (12) | Cột L | **Công thức** | =CEILING(K3,100000) |
| **M** (13) | Cột M | Dữ liệu nhập | 0.9 |
| **N** (14) | Cột N | **Công thức** | =O3*0.25 |
| **O** (15) | Cột O | **Công thức** | =C3*M3 |
| **P** (16) | Cột P | **Công thức** | =T3*M3 |
| **Q** (17) | Cột Q | Dữ liệu nhập | 1.0 |
| **R** (18) | Cột R | **Công thức** | =S3*0.25 |
| **S** (19) | Cột S | **Công thức** | =C3*Q3 |
| **T** (20) | 2026-07-31 00:00:00 | **Công thức** | =C3/DAY($T$1) |
| **U** (21) | Cột U | **Công thức** | =SUMIF('Dự án T7'!$A:$A,B3,'Dự án T7'!D:D)*DAY($F$1) |
| **V** (22) | Cột V | **Công thức** | =U3/D3 |
| **W** (23) | Cột W | **Công thức** | =IF(OR(F3="",V3=""),"",IF(AND(F3>=80%,V3>=25%),"Đạt","Không đạt")) |
| **X** (24) | Cột X | **Công thức** | =SUMIF(data!$A:$A,B3,data!$C:$C) |
| **Y** (25) | Cột Y | **Công thức** | =SUMIF(data!$A:$A,B3,data!$D:$D) |
| **Z** (26) | Cột Z | **Công thức** | =Y3/X3 |
| **AA** (27) | Cột AA | **Công thức** | =SUMIF(data!$A:$A,B3,data!$E:$E) |
| **AB** (28) | Cột AB | **Công thức** | =SUMIF(data!$A:$A,B3,data!$F:$F) |
| **AC** (29) | Cột AC | **Công thức** | =AB3/AA3 |
| **AD** (30) | Cột AD | **Công thức** | =X3+AA3 |
| **AE** (31) | Cột AE | **Công thức** | =Y3+AB3 |
| **AF** (32) | Cột AF | **Công thức** | =AE3/AD3 |

---

### 📌 SHEET: Dự án T7 (939 dòng × 32 cột)

| Cột | Tên trường / Tiêu đề cột | Loại | Công thức Excel Chuẩn Tháng 7 / Giá trị mẫu |
|:---:|---|:---:|---|
| **A** (1) | Nhà thuốc | Dữ liệu nhập | Hàng Bông |
| **B** (2) | Nhân viên | Dữ liệu nhập | Lê Ngọc Anh |
| **C** (3) | Chức danh | Dữ liệu nhập | CHT |
| **D** (4) | CK + Combo + NY3/ngày | **Công thức** | =SUM(E3:G3)/DAY($A$1) |
| **E** (5) | NY3 | Dữ liệu nhập | 0.0 |
| **F** (6) | Chiết khấu | Dữ liệu nhập | 0.0 |
| **G** (7) | Combo Liều | Dữ liệu nhập | 0.0 |
| **H** (8) | Chiết khấu/ngày | **Công thức** | =F3/DAY($A$1) |
| **I** (9) | Combo Liều/ngày | **Công thức** | =G3/DAY($A$1) |
| **J** (10) | Thưởng dự án | **Công thức** | =IFERROR(   IF(     OR($A3="Hàng Bông",$A3="Đường Láng"),     IFS(       AND($H3>=3400000,$I3>=800000),4000000,       AND($H3>=3000000,$I3>=700000),2800000,       AND($H3>=2600000,$I3>=550000),2000000,       AND($H3>=2200000,$I3>=450000),1200000,       TRUE,0     ),     IFS(       AND($H3>=1500000,$I3>=700000),4000000,       AND($H3>=1200000,$I3>=650000),2800000,       AND($H3>=1000000,$I3>=580000),2200000,       AND($H3>=850000,$I3>=480000),1600000,       AND($H3>=630000,$I3>=350000),1000000,       TRUE,0     )   ), 0) |
| **K** (11) | Thưởng thêm | **Công thức** | =IFERROR(   IF(     AND(       $J3=0,       IF(OR($A3="Hàng Bông",$A3="Đường Láng"),$D3>=2000000,$D3>=950000)     ),     500000,     0   ), 0) |
| **L** (12) | Total dự án | **Công thức** | =IFERROR($J3+$K3,0) |

---
