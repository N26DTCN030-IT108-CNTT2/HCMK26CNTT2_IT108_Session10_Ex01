# HCMK26CNTT2_IT108_Session10_Ex01
Bài tập 1
# Bộ Hồ Sơ Đối Soát Dự Án Ví Điện Tử PayKX (Session 12 Ex01)

Bài tập phân tích và chuẩn hóa tài liệu đối soát giao dịch tháng cho dự án ví điện tử PayKX dành cho sinh viên.

---

# Phần 1. Báo Cáo Phân Tích Dữ Liệu (Analysis Report)

## 1. Phân tích Record và Field trong Bảng Giao Dịch Excel

Dựa trên bảng tổng hợp giao dịch Tong_Hop_Phi_Ngan_Hang.xlsx:

* **Định nghĩa Field (Trường / Cột dữ liệu):**
  - Cột Ma_Giao_Dich: Chứa mã định danh duy nhất của từng giao dịch (Ví dụ: GD001, GD002).
  - Cột Ten_Nguoi_Dung: Trữ lượng thông tin chứa tên sinh viên thực hiện giao dịch (Lấy từ hàm `VLOOKUP`).
  - Cột So_Tien_VND: Giá trị tiền tệ của giao dịch thanh toán.
  - Cột Kenh_Thanh_Toan: Cổng/Ngân hàng xử lý thanh toán (MoMo, Vietcombank, ShopeePay).

* **Định nghĩa Record (Bản ghi / Dòng dữ liệu):**
  - Ví dụ Record Dòng 2: GD001 | Nguyễn Văn An | 150,000 | MoMo
  - Ý nghĩa: Tập hợp toàn bộ thuộc tính mô tả trọn vẹn cho **chính xác 1 sự kiện thanh toán** của sinh viên Nguyễn Văn An thông qua kênh MoMo với số tiền 150.000 VNĐ.

---

## 2. Báo Cáo Quy Trình Kiểm Thử 4 Bước Khi Gặp Lỗi `#N/A` (Troubleshooting)

Khi hàm VLOOKUP tra cứu tên người dùng trả về lỗi #N/A, quy trình kiểm tra 4 điểm được thực hiện như sau:

1. **Điểm 1 - Kiểm tra sự tồn tại của Mã tra cứu:** So sánh xem giá trị Ma_Giao_Dich ở bảng gốc có thực sự nằm trong cột đầu tiên của bảng danh mục tra cứu (DanhMucUser) hay không.
2. **Điểm 2 - Kiểm tra bẫy khoảng trắng (Dùng hàm TRIM):** Kiểm tra xem mã ở một trong hai bảng có bị dính khoảng trắng ẩn ở đầu/cuối hay không (ví dụ: 'GD001 ' khác với 'GD001'). Khắc phục bằng cách bọc TRIM(A2).
3. **Điểm 3 - Kiểm tra tham số khớp chính xác (FALSE/ 0):** Đảm bảo tham số thứ 4 của VLOOKUP được đặt là FALSE hoặc 0 để ép tìm kiếm chính xác tuyệt đối theo chuỗi Unicode.
4. **Điểm 4 - Kiểm tra cố định địa chỉ vùng tra cứu (Dấu `$`) và bọc IFERROR:** Đảm bảo Table_array được cố định bằng dấu $
5. (DanhMucUser!A$2:B$5) để tránh trôi vùng dữ liệu khi kéo công thức. Bọc thêm hàm =IFERROR(VLOOKUP(...), "Không tìm thấy") để hiển thị thông báo thay vì lỗi rác.

---

# Phần 2. Mô Tả Chi Tiết Bộ Sản Phẩm Trải Nghiệm

## 1. Tài liệu SRS Mô tả Tính năng Hoàn tiền (`SRS_PayKX_Feature.docx`)
* **Phân cấp Heading Styles chuẩn:**
  - Tiêu đề chính: Gán nhãn `Heading 1` (Size 16, Font Calibri, In đậm, Màu Navy `#102C57`).
  - Tiêu đề mục con (1.1, 1.2, 2.1, 2.2): Gán nhãn `Heading 2` (Size 13, Màu Slate Blue `#35598E`).
  - Nội dung mô tả: Nhãn `Normal`.
* **Mục lục tự động (Auto TOC):** Tích hợp trường Table of Contents tự động thu thập và cập nhật số trang thời gian thực.

---

## 2. Bảng Tổng Hợp Phí Ngân Hàng (`Tong_Hop_Phi_Ngan_Hang.xlsx`)
* **Tuân thủ 5 Quy tắc Vàng Database:**
  - 1 hàng Header duy nhất ở Dòng 1, không gộp ô (Merge Cell).
  - Không chứa dòng/cột trống giữa các bản ghi.
  - Dòng tổng cộng nằm hoàn toàn bên ngoài bảng dữ liệu thô.
* **Công thức VLOOKUP + TRIM chống lỗi:**
  ```excel
  =IFERROR(VLOOKUP(TRIM(A2), DanhMucUser!A$2:B$5, 2, FALSE), "Không tìm thấy")
``` 











  
