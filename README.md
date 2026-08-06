# ⚡ Financial Data ETL Tool - Streamlit Cloud Web Application

Ứng dụng Web xử lý dữ liệu báo cáo tài chính từ file **PDF, XLSX, XLS** thành **1 file CSV/XLSX chuẩn 26 cột thuộc tính** theo sơ đồ **Data Model.jpg** và sinh **SQL Script DDL/DML** sẵn sàng đồng bộ lên hệ thống Database (Supabase / PostgreSQL).

![Data Model Architecture](Data%20Model.jpg)

---

## 🌟 Tính Năng Nổi Bật

1. **Xử Lý Dữ Liệu Động (Dynamic ETL Engine)**:
   - Hỗ trợ tải lên **1 hoặc nhiều file bất kỳ** (.pdf, .xlsx, .xls).
   - Tự động trích xuất thông tin giao dịch, tiểu khoản, mã chứng khoán, khối lượng, giá trị, thuế, phí net, môi giới, CTV và người quản lý.

2. **Chuẩn Hóa Khớp 100% Data Model**:
   - Sinh file tổng hợp **26 cột thuộc tính** gộp từ 10 thực thể database trong sơ đồ Data Model:
     - `Giao dịch`, `Khách hàng`, `Người quản lý`, `Công ty chứng khoán`, `Cổ phiếu`, `Phân loại khách hàng`, `Nhóm khách hàng`, `Chính sách`, `Phí gia hạn`, `Báo cáo thu lãi`.

3. **Xuất Đa Định Dạng (Multi-Format Export)**:
   - 📄 **Master File CSV (`Master_Data_Model_Output.csv`)**: File CSV tổng hợp 26 cột định dạng UTF-8.
   - 📊 **Master File Excel (`Master_Data_Model_Output.xlsx`)**: File Excel chứa trang Master và 10 trang chi tiết từng thực thể.
   - 📦 **ZIP Package CSVs (`Database_CSVs_Package.zip`)**: Bộ file CSV riêng biệt cho 10 bảng database.
   - 📜 **SQL Script (`supabase_schema_and_data.sql`)**: Script tự động tạo bảng `CREATE TABLE` và chèn dữ liệu `INSERT INTO` trên Supabase / PostgreSQL.

---

## 🛠️ Hướng Dẫn Chạy Tại Local

```powershell
# 1. Cài đặt các thư viện cần thiết
pip install -r requirements.txt

# 2. Khởi chạy ứng dụng Streamlit
streamlit run app.py
```

Trình duyệt sẽ tự động mở trang web tại địa chỉ: `http://localhost:8501`.

---

## ☁️ Hướng Dẫn Deploy Lên Streamlit Community Cloud (Miễn Phí Public Link)

### Bước 1: Tạo Repository Trên GitHub
1. Truy cập [github.com/new](https://github.com/new).
2. Đặt tên Repository: `financial-data-etl`.
3. Đẩy toàn bộ các file trong thư mục này lên GitHub:
   ```bash
   git init
   git add .
   git commit -m "Initial commit - Streamlit Financial Data ETL Tool"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/financial-data-etl.git
   git push -u origin main
   ```

### Bước 2: Deploy Trên Streamlit Cloud
1. Truy cập trang: [share.streamlit.io](https://share.streamlit.io).
2. Đăng nhập bằng tài khoản GitHub.
3. Chọn **New app** -> Chọn repository `YOUR_USERNAME/financial-data-etl`.
4. Điền **Main file path**: `app.py`.
5. Bấm **Deploy!**. Sau 1 phút bạn sẽ có link public công khai.
