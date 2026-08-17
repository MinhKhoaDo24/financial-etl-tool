# 📊 Hệ Thống Chuẩn Hóa & Gộp Báo Cáo Giao Dịch JB + PBSV

Ứng dụng Web Streamlit giúp tự động hóa quá trình đọc, chuẩn hóa, đối chiếu danh sách khách hàng và gộp báo cáo giao dịch chứng khoán từ hai nguồn **JB** và **PBSV** thành một file báo cáo duy nhất.

---

## 🌟 Tính Năng Nổi Bật

1. **Xử Lý Dữ Liệu Linh Hoạt & Đa Nguồn**:
   - Hỗ trợ chọn **nhiều file cùng lúc** cho từng loại báo cáo (.xlsx, .xls, .csv).
   - **Đầu vào tối thiểu linh hoạt**: Chỉ cần file **Danh sách Khách hàng (Master)** và **ít nhất 1 file giao dịch** (JB hoặc PBSV) là ứng dụng đã có thể tự động xử lý.

2. **Chuẩn Hóa & Nhận Diện Mã Công Ty Rút Gọn**:
   - Tự động nhận diện và chuyển đổi cột `Tên Công ty` sang mã rút gọn chuẩn: **`JB`** hoặc **`PBSV`**.
   - Lọc chính xác các giao dịch ở trạng thái *"Hoàn thành"* đối với dữ liệu PBSV.
   - Tự động tính toán các trường derived như `% Phí` nếu thiếu, và luôn tính `Tổng giá trị giao dịch (sau phí + thuế)` theo đúng 1 công thức chuẩn: `= Tổng giá trị giao dịch (trước phí + thuế) + Phí + Thuế` cho mọi dòng (không phân biệt nguồn JB/PBSV hay Mua/Bán).

3. **Giao Diện Dashboard Hiện Đại & Tùy Biến Chuyên Sâu**:
   - **Thanh Sidebar điều khiển**: Gom gọn khu vực tải file và nút thao tác giúp không gian làm việc chính rộng rãi.
   - **Cấu trúc Thẻ (Card Container) & Tab Workflows**: Phân chia thông tin rõ ràng qua 3 Tab:
     - 📋 *Tab 1: Bảng Kết Quả Chính* (Xem & sửa trực tiếp dạng Excel).
     - ⚠️ *Tab 2: Dòng Không Khớp Khách Hàng* (Kiểm tra & bổ sung thủ công).
     - 📥 *Tab 3: Tải Xuất Báo Cáo* (Tải file CSV và Excel).
   - **Cấu hình Theme dễ dàng**: Tùy chỉnh màu sắc, font chữ và chế độ Sáng/Tối qua file `.streamlit/config.toml` được chú thích chi tiết bằng tiếng Việt.

4. **Tự Động Đồng Bộ Chỉnh Sửa Tay**:
   - Cho phép click đúp sửa trực tiếp bất kỳ ô dữ liệu nào trên bảng web.
   - Nút tải file **CSV** và **Excel (.xlsx)** cam kết tự động xuất ra **dữ liệu mới nhất đã bao gồm toàn bộ thao tác sửa tay**.

---

## 📁 Cấu Trúc Thư Mục Dự Án

```text
├── app.py                      # Giao diện Web chính (Streamlit UI, Tabs, Cards, Sidebar)
├── merge_processor.py          # Module ETL đọc file, chuẩn hóa, join đối chiếu & gộp dữ liệu
├── .streamlit/
│   └── config.toml             # Cấu hình Theme Streamlit (Màu sắc, Font, Light/Dark mode)
├── requirements.txt            # Danh sách thư viện Python phụ thuộc
├── DATA MODEL.png              # Sơ đồ kiến trúc Data Model
└── README.md                   # Tài liệu hướng dẫn dự án
```

---

## 🛠️ Hướng Dẫn Khởi Chạy Tại Local

### 1. Cài đặt thư viện phụ thuộc
Mở terminal tại thư mục dự án và chạy lệnh:
```bash
pip install -r requirements.txt
```

### 2. Chạy ứng dụng Streamlit
```bash
streamlit run app.py
```
Trình duyệt sẽ tự động mở trang web tại địa chỉ: `http://localhost:8501`.

---

## 🎨 Hướng Dẫn Tùy Chỉnh Giao Diện (Theme)

Bạn có thể thay đổi màu sắc ứng dụng theo nhận diện thương hiệu bằng cách mở file [`.streamlit/config.toml`](.streamlit/config.toml) và chỉnh sửa các mã màu HEX:

- `primaryColor`: Màu nút bấm chính (`Xử lý`), thanh highlight.
- `backgroundColor`: Màu nền trang chính (mặc định: `#F8FAFC`).
- `secondaryBackgroundColor`: Màu nền của Sidebar và các thẻ Card (mặc định: `#FFFFFF`).
- `textColor`: Màu chữ (mặc định: `#1E293B`).
- `font`: Font chữ hiển thị (`"sans serif"`, `"serif"`, `"monospace"`).

---

## ☁️ Hướng Dẫn Deploy Lên Streamlit Community Cloud (Miễn Phí)

1. **Đẩy mã nguồn lên GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Deploy Streamlit JB PBSV ETL Tool"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
   git push -u origin main
   ```
2. **Deploy trên Streamlit Cloud**:
   - Truy cập [share.streamlit.io](https://share.streamlit.io) và đăng nhập bằng GitHub.
   - Chọn **New app** -> Chọn Repository và nhánh `main`.
   - Cấu hình **Main file path**: `app.py`.
   - Bấm **Deploy!** để lấy đường link công khai.
