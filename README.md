# 🚀 HỆ THỐNG GẮN MÃ KẾ TOÁN TỰ ĐỘNG (AUTO-ACCOUNTING MAPPER)

Ứng dụng chuyên nghiệp chạy trên máy cá nhân để hỗ trợ đồng bộ dữ liệu kế toán: **Tự động gán mã vật tư, hàng hóa và đối tác (Khách hàng/Nhà cung cấp)** vào các bảng dữ liệu phát sinh gồm:
- Bảng kê hóa đơn mua vào.
- Bảng kê hóa đơn bán ra.
- Sổ kho, báo cáo Xuất - Nhập - Tồn.
- Nhật ký sao kê giao dịch ngân hàng.

---

## 🎨 CHỨC NĂNG CHÍNH

1. **Gắn mã hàng hóa thông minh**: So khớp diễn giải hóa đơn với danh mục vật tư bằng thuật toán kết hợp đối chiếu quy cách (`500ml`, `10x20`, `A4`, `PCB40`, `Phi 6`,...), trùng lặp từ khóa và tính khoảng cách chuỗi mờ (**Fuzzy String Distance**).
2. **Gắn mã và phân loại đối tác**: Trích xuất mã số thuế, số tài khoản ngân hàng để gán mã khách hàng (`KHxxx`) / nhà cung cấp (`NCCxxx`), đồng thời tự động nhận diện từ viết tắt tên công ty (`TNHH`, `MTV`, `CP`,...).
3. **Phân tích thông minh sao kê ngân hàng**: Phán đoán chính xác các luồng tiền thu (khách hàng), chi (nhà cung cấp), tự động lọc bỏ và ghi nhận các chi phí hệ thống (phí ngân hàng, tiền thuế, chi lương nhân viên, luân chuyển tiền nội bộ,...).
4. **Màn hình kiểm duyệt tương tác (Interactive Editor)**: Cho phép sửa đổi trực tiếp các đề xuất của máy tính, thêm ghi chú kiểm tra và sinh mã mới linh hoạt ngay trên giao diện web.
5. **Tổng hợp đối chiếu chéo**: So khớp doanh thu hóa đơn bán ra với tiền thực tế thu về ngân hàng để kiểm tra chênh lệch công nợ phải thu của từng đối tác.

---

## 🛠️ HƯỚNG DẪN CÀI ĐẶT & CHẠY TRÊN MÁY CÁ NHÂN (LOCAL APP)

### 📌 Yêu cầu chuẩn bị
Máy tính của bạn cần cài đặt sẵn **Python 3.9+** (khuyên dùng Python 3.10 hoặc 3.11).

### 🖥️ Các bước thực hiện chi tiết:

1. **Mở Command Prompt (Windows) hoặc Terminal (macOS/Linux)** tại thư mục chứa dự án.
2. **Tạo môi trường ảo (Khuyên dùng)** nhằm tránh xung đột thư viện:
   ```bash
   python -m venv venv
   ```
3. **Kích hoạt môi trường ảo**:
   - Ở Windows:
     ```bash
     venv\Scripts\activate
     ```
   - Ở macOS/Linux:
     ```bash
     source venv/bin/activate
     ```
4. **Cài đặt các gói thư viện cần thiết** từ tệp `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```
5. **Chạy ứng dụng bằng Streamlit**:
   ```bash
   streamlit run app.py
   ```
6. Trình duyệt web sẽ tự động mở trang ứng dụng tại địa chỉ: `http://localhost:8501`.

---

## 📈 CHI TIẾT THUẬT TOÁN TÍNH ĐIỂM TƯƠNG THÍCH (0 - 100 ĐIỂM)

### 📦 Logic Gắn Mã Hàng Hóa:
- **Khớp tên tuyệt đối** (sau chuẩn hóa không dấu, chữ thường): `+60 điểm`
- **Fuzzy Ratio** (Tỷ lệ khớp mờ ký tự theo thuật toán RapidFuzz):
  - Khớp trên `90%`: `+50 điểm`
  - Khớp từ `75%` đến `90%`: `+35 điểm`
- **Trùng khớp từ khóa nhận diện** trong danh mục: `+20 điểm`
- **Trùng khớp Đơn vị tính (ĐVT)**: `+10 điểm`
- **Trùng khớp Nhóm hàng**: `+10 điểm`
- **Khớp thông số quy cách trích xuất** (ví dụ: `A4`, `5L`, `PCB40`, `Phi 6`): `+20 điểm`
- *Điểm tối đa được giới hạn ở mức 100.*

### 👥 Logic Gắn Mã Khách Hàng / Nhà Cung Cấp:
- **Trùng khớp Mã Số Thuế (MST)**: `+60 điểm` (Xác thực tuyệt đối)
- **Trùng khớp Số tài khoản ngân hàng**: `+50 điểm`
- **Trùng tên tuyệt đối** sau khi chuẩn hóa (loại bỏ từ như *Công ty, TNHH, MTV, CP*): `+50 điểm`
- **Fuzzy Ratio** (Độ gần giống tên riêng):
  - Khớp trên `90%`: `+40 điểm`
  - Khớp từ `75%` đến `90%`: `+25 điểm`
- **Trùng từ khóa đối tác**: `+20 điểm`
- **Nội dung diễn giải có chứa tên đối tượng**: `+20 điểm`
- *Điểm tối đa được giới hạn ở mức 100.*

---

## 🏗️ CÁC TỘT DỮ LIỆU ĐƯỢC BỔ SUNG TRONG FILE EXCEL ĐẦU RA

Ứng dụng bảo đảm giữ nguyên 100% tất cả các cột dữ liệu gốc của bạn, không bao giờ làm mất mát hay ghi đè thông tin cũ, chỉ bổ sung thêm các cột phân tích tự động ở cuối bảng:
- **`Mã hàng hóa` / `Mã đối tượng`**: Mã chuẩn được gán từ danh mục hoặc mã mới tự động sinh.
- **`Tên hàng hóa chuẩn` / `Tên đối tượng chuẩn`**: Tên hiển thị chuẩn chỉ trong cơ sở dữ liệu.
- **`Tìm thấy từ nguồn đối chiếu`**: Mô tả tóm tắt lý do hoặc tên đối tác thô tìm thấy.
- **`Tỷ lệ tương thích (%)`**: Điểm số tin cậy do thuật toán quyết định.
- **`Mức độ tương thích`**: Phân loại mức độ bằng nhãn `Cao`, `Cần kiểm tra` hoặc `Mã mới tinh`.
- **`Lý do gắn mã`**: Trình bày từng thành phần điểm cộng một cách rành mạch để kế toán dễ đối chiếu.
- **`Trạng thái xử lý`**: Thể hiện trạng thái `TỰ ĐỘNG GẮN`, `DUYỆT THỦ CÔNG`, `TẠO MÃ MỚI` hoặc `BỎ QUA`.
- **`Ghi chú kiểm tra/Sửa đổi`**: Trường trống dành riêng cho kế toán viết nhận định chỉnh lý sau này.
