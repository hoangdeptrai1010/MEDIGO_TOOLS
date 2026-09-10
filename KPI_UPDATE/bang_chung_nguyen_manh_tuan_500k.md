# BẰNG CHỨNG XÁC MINH: NGUYỄN MẠNH TUẤN NHẬN DƯ 500,000Đ TIỀN THƯỞNG DỰ ÁN THÁNG 8/2026

## 1. Căn cứ văn bản quy chế chính thức
* **Văn bản**: Slide trình chiếu thể lệ Dự Án Bán Lẻ Hà Nội Tháng 8/2026.
* **Đường dẫn tệp gốc**: `thang8/Pharmacy_retail_Store_KPIs_August 2026/Hà Nội/Dự án/Chương trình/Hanoi_Projects_August_2026_v1.pptx` (Slide 1).
* **Nội dung nguyên văn quy định**:
  1. **Điều kiện hưởng thưởng khuyến khích 500K**:  
     > *"Tổng doanh thu **CK + Combo liều/DS/ngày $\ge$ 2M** $\rightarrow$ Thưởng khuyến khích dược sĩ **500K**"*
  2. **Ghi chú bắt buộc (in chữ đỏ)**:  
     > *"Lưu ý: **Nhóm CK, Combo liều bao gồm doanh thu nhóm mã hàng Chiết khấu, Combo liều (không ghi nhận doanh thu các mã LL)**"*  
     *(Quy chế chỉ tính 2 nhóm là Chiết Khấu và Combo Liều, hoàn toàn không cho phép cộng nhóm NY3 - Thuốc kê đơn)*.

---

## 2. Số liệu thực tế của Dược sĩ Nguyễn Mạnh Tuấn
*(Trích xuất từ Dòng 5 sheet `Dự án T8` - Tính trên 31 ngày làm việc của Tháng 8/2026)*

| Cột | Hạng mục doanh thu | Số tiền thực đạt (VNĐ) | Doanh thu trung bình/ngày (VNĐ/ngày) |
| :---: | :--- | :---: | :---: |
| **Cột F** | Doanh số Chiết khấu (CK) | 46,975,942 | 1,515,353 |
| **Cột G** | Doanh số Combo Liều | 11,695,475 | 377,273 |
| **F + G** | **Tổng CK + Combo Liều (Theo đúng Slide)** | **58,671,417** | **1,892,626** |
| | **Mức sàn chỉ tiêu trên Slide** | | **2,000,000** |
| | **Chênh lệch so với chỉ tiêu** | | **-107,374 (Thiếu 107k/ngày)** |
| | **KẾT LUẬN THEO QUY CHẾ** | | ❌ **KHÔNG ĐẠT (Thưởng = 0 VNĐ)** |

---

## 3. Nguyên nhân phát sinh sai lệch 500,000đ trong file Excel
Lỗi phát sinh do công thức tính toán tại Dòng 5 sheet `Dự án T8`:
1. **Tại Ô D5 (Doanh thu ngày xét thưởng)**:
   * Công thức cài đặt sẵn: `=SUM(E5:G5)/DAY($A$1)`
   * Ô E5 chứa nhóm **NY3 = 7,086,966 VNĐ**.
   * Việc cộng thêm cột E (NY3) đã vô tình kéo tổng doanh thu từ **58,671,417đ** lên **65,758,383đ**.
   * Doanh thu ngày bị đội lên: $65,758,383 / 31 = \mathbf{2,121,238\text{ VNĐ/ngày}}$.
2. **Tại Ô K5 (Thưởng thêm 500k)**:
   * Công thức: `=IFERROR(IF(AND($J5=0, IF(OR($A5="Hàng Bông",$A5="Đường Láng"), $D5>=2000000, $D5>=950000)), 500000, 0), 0)`
   * Vì $D5 = 2,121,238đ \ge 2,000,000đ$, điều kiện trả về `TRUE` $\rightarrow$ Kích hoạt mức thưởng **500,000 VNĐ**.

---

## 4. Bảng đối chứng với các Dược sĩ khác cùng nhận mốc 500k
Tất cả 6 nhân sự khác nhận thưởng thêm 500k đều đạt thực chất bằng `CK + Combo` mà không cần đến NY3:
* **Đinh Thị Lan Anh** (Hàng Bông): Riêng CK đạt 2,021,214đ/ngày ($\ge 2.0\text{M}$).
* **Đinh Thị Khánh Ly** (Hàng Bông): Riêng CK đạt 3,208,549đ/ngày ($\ge 2.0\text{M}$).
* **Hoàng Thanh Thủy** (Đỗ Quang Đẩu): CK + Combo đạt 1,043,436đ/ngày ($\ge 950\text{k}$).
* **Trần Hoàng Khánh** (Nguyễn Chí Thanh): CK + Combo đạt 1,013,216đ/ngày ($\ge 950\text{k}$).
* **Cao Trọng Nhân** (Nguyễn Thị Thập): CK + Combo đạt 1,001,531đ/ngày ($\ge 950\text{k}$).
* **Nguyễn Thị Hương Giang** (Nguyễn Văn Quá): CK + Combo đạt 1,011,051đ/ngày ($\ge 950\text{k}$).

---

## 5. Quyền lợi thực tế hợp lệ của Nguyễn Mạnh Tuấn
* **Thưởng Dự án chính (Ô J5)**: 0 VNĐ.
* **Thưởng thêm khuyến khích (Ô K5)**: 0 VNĐ (Không đạt điều kiện $\ge 2.0\text{M}$).
* **Thưởng Hot Bill Hà Nội (Ô L5)**: **100,000 VNĐ** *(Hợp lệ 100% với 2 hóa đơn giao nhanh đạt chuẩn ngày 25/08 và 28/08)*.
* **Tổng tiền thưởng Dự án đúng quy chế (Ô M5)**: **100,000 VNĐ**.
