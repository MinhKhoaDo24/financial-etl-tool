import streamlit as st
import pandas as pd
import io
import zipfile
import os
import etl_processor as etl

# Set Streamlit Page Configuration
st.set_page_config(
    page_title="Dynamic Financial Data ETL Tool",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Aesthetics
st.markdown("""
<style>
    .main {
        background-color: #0e1117;
        font-family: 'Inter', -apple-system, sans-serif;
    }
    .header-box {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .header-title {
        color: #f8fafc;
        font-size: 26px;
        font-weight: 700;
        margin: 0 0 8px 0;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .header-subtitle {
        color: #94a3b8;
        font-size: 14px;
        margin: 0;
    }
    .metric-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
    }
    .metric-value {
        color: #38bdf8;
        font-size: 22px;
        font-weight: 700;
    }
    .metric-label {
        color: #94a3b8;
        font-size: 13px;
        margin-top: 4px;
    }
    .badge {
        background-color: #2563eb;
        color: #ffffff;
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 600;
    }
    .highlight-download {
        background: #1e293b;
        border: 2px solid #2563eb;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Default Sample Files
DEFAULT_XLS_PATH = os.path.join(os.path.dirname(__file__), "Bao cao giao dich 04.08.26.xls")
DEFAULT_PDF_PATH = os.path.join(os.path.dirname(__file__), "RE1002 T7 LMH.pdf")
DEFAULT_IMG_PATH = os.path.join(os.path.dirname(__file__), "Data Model.jpg")

# Render Sidebar
st.sidebar.title("⚙️ Tùy Chọn Tải Dữ Liệu")
st.sidebar.markdown("---")

uploaded_files = st.sidebar.file_uploader(
    "📤 Tải lên file báo cáo (.pdf, .xlsx, .xls):",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Có thể tải lên 1 hoặc nhiều file PDF/Excel bất kỳ"
)

use_default = st.sidebar.checkbox("📁 Sử dụng file mẫu trong folder (Demo)", value=not bool(uploaded_files))

st.sidebar.markdown("---")
st.sidebar.markdown("### 🗄️ Đích đẩy dữ liệu Database")
db_target = st.sidebar.selectbox(
    "Định dạng xuất dữ liệu:",
    ["Single Master File (CSV / XLSX)", "Multi-sheet Excel (.xlsx)", "ZIP Package CSVs (.zip)", "Supabase SQL Script (.sql)"]
)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Gợi ý:** Bạn có thể xuất 1 file CSV hoặc XLSX duy nhất chứa đầy đủ các cột thông tin từ Data Model để đẩy lên Database.")

# Header
st.markdown("""
<div class="header-box">
    <div class="header-title">
        <span>⚡ Tool Xử Lý Dữ Liệu Tài Chính - Định Dạng Data Model</span>
        <span class="badge">Streamlit ETL v1.1</span>
    </div>
    <div class="header-subtitle">
        Tự động chuyển đổi báo cáo PDF, XLSX hoặc XLS thành 1 file CSV/XLSX duy nhất chứa đầy đủ các cột thuộc tính quy định trong sơ đồ Data Model.
    </div>
</div>
""", unsafe_allow_html=True)

# Collect processing inputs
excel_tx_all = []
pdf_tx_all = []
processed_file_names = []

if uploaded_files:
    for f in uploaded_files:
        fname = f.name
        fbytes = f.getvalue()
        processed_file_names.append(fname)
        if fname.lower().endswith(('.xlsx', '.xls')):
            _, txs = etl.parse_excel_data(fbytes, filename=fname)
            excel_tx_all.extend(txs)
        elif fname.lower().endswith('.pdf'):
            _, txs = etl.parse_pdf_data(fbytes, filename=fname)
            pdf_tx_all.extend(txs)
elif use_default:
    if os.path.exists(DEFAULT_XLS_PATH):
        _, txs = etl.parse_excel_data(DEFAULT_XLS_PATH, filename="Bao cao giao dich 04.08.26.xls")
        excel_tx_all.extend(txs)
        processed_file_names.append("Bao cao giao dich 04.08.26.xls")
    if os.path.exists(DEFAULT_PDF_PATH):
        _, txs = etl.parse_pdf_data(DEFAULT_PDF_PATH, filename="RE1002 T7 LMH.pdf")
        pdf_tx_all.extend(txs)
        processed_file_names.append("RE1002 T7 LMH.pdf")

if not excel_tx_all and not pdf_tx_all:
    st.warning("⚠️ Vui lòng tải lên ít nhất 1 file (.pdf, .xlsx hoặc .xls) hoặc tích chọn 'Sử dụng file mẫu' ở bảng bên trái.")
    st.stop()

# Build Relational Database & Master Flat Table
db_tables = etl.build_relational_database(excel_tx_all, pdf_tx_all)
master_df = etl.build_master_flat_table(db_tables)

# Display Metrics
col1, col2, col3, col4, col5 = st.columns(5)
total_tx = len(db_tables['Giao_dich'])
total_val = db_tables['Giao_dich']['Giá trị giao dịch'].sum() if not db_tables['Giao_dich'].empty else 0
total_cust = len(db_tables['Khach_hang'])
total_stock = len(db_tables['Co_phieu'])
total_files = len(processed_file_names)

with col1:
    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_files}</div><div class="metric-label">File đã xử lý</div></div>', unsafe_allow_html=True)
with col2:
    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_tx}</div><div class="metric-label">Tổng bản ghi</div></div>', unsafe_allow_html=True)
with col3:
    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_val:,.0f} đ</div><div class="metric-label">Tổng giá trị</div></div>', unsafe_allow_html=True)
with col4:
    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_cust}</div><div class="metric-label">Khách hàng</div></div>', unsafe_allow_html=True)
with col5:
    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_stock}</div><div class="metric-label">Mã cổ phiếu</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# QUICK DOWNLOAD SECTION AT TOP
st.markdown("""
<div class="highlight-download">
    <h3 style="margin-top:0; color:#38bdf8;">📥 Tải Nhanh File Kết Quả Đã Chuẩn Hóa Theo Data Model</h3>
    <p style="color:#94a3b8; font-size:14px;">File CSV và XLSX dưới đây chứa đầy đủ các cột thuộc tính hợp nhất từ tất cả các bảng trong sơ đồ Data Model.</p>
</div>
""", unsafe_allow_html=True)

dl_col1, dl_col2, dl_col3 = st.columns(3)

# Master CSV Byte stream
master_csv_bytes = master_df.to_csv(index=False).encode('utf-8-sig')

# Master XLSX Byte stream
excel_master_buffer = io.BytesIO()
with pd.ExcelWriter(excel_master_buffer, engine='openpyxl') as writer:
    master_df.to_excel(writer, sheet_name="Master_Data_Model", index=False)
    for tbl_name, df_tbl in db_tables.items():
        df_tbl.to_excel(writer, sheet_name=tbl_name[:31], index=False)
master_excel_bytes = excel_master_buffer.getvalue()

with dl_col1:
    st.download_button(
        label="📄 Tải File CSV Tổng Hợp (Master Data.csv)",
        data=master_csv_bytes,
        file_name="Master_Data_Model_Output.csv",
        mime="text/csv",
        use_container_width=True
    )

with dl_col2:
    st.download_button(
        label="📊 Tải File Excel Tổng Hợp (Master Data.xlsx)",
        data=master_excel_bytes,
        file_name="Master_Data_Model_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with dl_col3:
    sql_script_content = etl.generate_sql_script(db_tables)
    st.download_button(
        label="📜 Tải Script SQL Supabase (.sql)",
        data=sql_script_content.encode('utf-8'),
        file_name="supabase_schema_and_data.sql",
        mime="text/plain",
        use_container_width=True
    )

st.markdown("<br>", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 1. Xem Bảng Dữ Liệu Master (Data Model)",
    "📋 2. Bảng Theo Schema Riêng (10 Entities)",
    "🔍 3. Chi Tiết Trích Xuất Thô",
    "📐 4. Sơ Đồ Data Model (ERD)",
    "☁️ 5. Hướng Dẫn Deploy Web Tool"
])

# TAB 1: Master Table View
with tab1:
    st.subheader("📋 Bảng Tổng Hợp Chứa Đầy Đủ Tất Cả Cột Theo Data Model")
    st.caption("Bảng dữ liệu phẳng đã được JOIN giữa Giao dịch, Khách hàng, Người quản lý, Công ty chứng khoán & Cổ phiếu.")
    
    st.dataframe(master_df, use_container_width=True)
    st.markdown(f"**Tổng số cột thông tin:** `{len(master_df.columns)}` cột | **Tổng số dòng:** `{len(master_df)}` dòng")

# TAB 2: Separate Entities
with tab2:
    st.subheader("📋 10 Bảng Dữ Liệu Tách Biệt Theo Thực Thể Database")
    selected_table = st.selectbox("Chọn bảng cần xem:", list(db_tables.keys()), index=7)
    df_selected = db_tables[selected_table]
    
    st.dataframe(df_selected, use_container_width=True)
    st.download_button(
        label=f"⬇️ Tải CSV cho bảng `{selected_table}`",
        data=df_selected.to_csv(index=False).encode('utf-8-sig'),
        file_name=f"{selected_table}.csv",
        mime="text/csv"
    )

# TAB 3: Raw Extracted Data
with tab3:
    st.subheader("🔍 Dữ Liệu Thô Trích Xuất")
    if excel_tx_all:
        st.markdown("#### 📄 Dữ liệu từ File Excel")
        st.dataframe(pd.DataFrame(excel_tx_all), use_container_width=True)
    if pdf_tx_all:
        st.markdown("#### 📑 Dữ liệu từ File PDF")
        st.dataframe(pd.DataFrame(pdf_tx_all), use_container_width=True)

# TAB 4: Data Model Image
with tab4:
    st.subheader("📐 Sơ đồ Schema Relational Data Model")
    if os.path.exists(DEFAULT_IMG_PATH):
        st.image(DEFAULT_IMG_PATH, caption="Relational Database Model Schema (Data Model.jpg)", use_column_width=True)

# TAB 5: Deployment Guide
with tab5:
    st.subheader("☁️ Hướng Dẫn Deploy Streamlit Web Tool Công Khai")
    st.markdown("""
    1. Upload toàn bộ các file (`app.py`, `etl_processor.py`, `requirements.txt`, `Data Model.jpg`) lên **GitHub**.
    2. Đăng nhập [share.streamlit.io](https://share.streamlit.io) bằng tài khoản GitHub.
    3. Chọn **New app**, nhập tên repository và bấm **Deploy**. Bạn sẽ nhận được đường link web trực tuyến chia sẻ cho mọi người dùng.
    """)

st.markdown("---")
st.caption("⚡ Financial Data ETL Tool | Streamlit Web Engine.")
