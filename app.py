import streamlit as st
import pandas as pd
import io
import zipfile
import os
import etl_processor as etl

# ────────────────────────────────────────────
# Page Config
# ────────────────────────────────────────────
st.set_page_config(
    page_title="Financial Data ETL Tool",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ────────────────────────────────────────────
# CSS
# ────────────────────────────────────────────
st.markdown("""
<style>
    .main { font-family: 'Inter', -apple-system, sans-serif; }
    .header-box {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 22px 26px;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0,0,0,.3);
    }
    .header-title { color:#f8fafc; font-size:24px; font-weight:700; margin:0 0 6px 0; }
    .header-subtitle { color:#94a3b8; font-size:13px; margin:0; }
    .metric-card {
        background:#1e293b; border:1px solid #334155;
        border-radius:10px; padding:14px; text-align:center;
    }
    .metric-value { color:#38bdf8; font-size:20px; font-weight:700; }
    .metric-label { color:#94a3b8; font-size:12px; margin-top:4px; }
    .badge {
        background:#2563eb; color:#fff;
        padding:3px 9px; border-radius:999px;
        font-size:11px; font-weight:600;
    }
    .dl-box {
        background:#1e293b; border:2px solid #2563eb;
        border-radius:12px; padding:18px; margin-bottom:18px;
    }
    .edit-hint {
        background:#0f2d1f; border:1px solid #22c55e;
        border-radius:8px; padding:10px 14px; color:#86efac;
        font-size:13px; margin-bottom:12px;
    }
</style>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────
# Paths
# ────────────────────────────────────────────
DEFAULT_XLS_PATH = os.path.join(os.path.dirname(__file__), "Bao cao giao dich 04.08.26.xls")
DEFAULT_PDF_PATH = os.path.join(os.path.dirname(__file__), "RE1002 T7 LMH.pdf")
DEFAULT_PBSV_XLSX_PATH = os.path.join(os.path.dirname(__file__), "PBSV.xlsx")
DEFAULT_IMG_PATH = os.path.join(os.path.dirname(__file__), "DATA MODEL.png")

# ────────────────────────────────────────────
# Sidebar
# ────────────────────────────────────────────
st.sidebar.title("⚙️ Tải & Xử Lý Dữ Liệu")
st.sidebar.markdown("---")

uploaded_files = st.sidebar.file_uploader(
    "📤 Tải lên file báo cáo (.pdf, .xlsx, .xls):",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Hỗ trợ: file giao dịch XLS/XLSX (khớp lệnh hoặc lịch sử đặt lệnh), file PDF RE1002, hoặc file Data Model nhiều sheet"
)

use_default = st.sidebar.checkbox(
    "📁 Sử dụng file mẫu Demo",
    value=not bool(uploaded_files)
)

st.sidebar.markdown("---")
st.sidebar.info("💡 **Tip chỉnh sửa:** Sau khi dữ liệu hiển thị trên bảng, bạn có thể **click vào bất kỳ ô nào** để sửa trước khi tải xuống.")

# ────────────────────────────────────────────
# Header
# ────────────────────────────────────────────
st.markdown("""
<div class="header-box">
    <div class="header-title">
        ⚡ Financial Data ETL Tool
        <span class="badge">v2.0 – Editable</span>
    </div>
    <div class="header-subtitle">
        Xử lý báo cáo giao dịch PDF / XLS / XLSX (khớp lệnh hoặc lịch sử đặt lệnh — tự động lọc giao dịch <strong>Hoàn thành</strong>).
        Dữ liệu có thể <strong>chỉnh sửa trực tiếp</strong> trên bảng trước khi xuất CSV / XLSX / SQL.
    </div>
</div>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────
# Detect file mode & build db_tables
# ────────────────────────────────────────────
excel_tx_all = []
pdf_tx_all = []
processed_file_names = []
is_direct_model = False        # True when file is PBSV multi-sheet model

db_tables_raw = {}             # from ETL or direct load

if uploaded_files:
    for f in uploaded_files:
        fname = f.name
        fbytes = f.getvalue()
        processed_file_names.append(fname)

        if fname.lower().endswith(('.xlsx', '.xls')):
            if etl.is_multi_sheet_data_model(fbytes):
                # ── PBSV multi-sheet direct load ──
                is_direct_model = True
                loaded = etl.parse_multi_sheet_excel(fbytes, filename=fname)
                # Merge into db_tables_raw (overwrite per table)
                for k, v in loaded.items():
                    if k not in db_tables_raw or db_tables_raw[k].empty:
                        db_tables_raw[k] = v
                    else:
                        db_tables_raw[k] = pd.concat([db_tables_raw[k], v], ignore_index=True)
            else:
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
    if os.path.exists(DEFAULT_PBSV_XLSX_PATH):
        # PBSV.xlsx = báo cáo "Lịch sử đặt lệnh" — parser tự lọc chỉ giữ giao dịch Hoàn thành
        _, txs = etl.parse_excel_data(DEFAULT_PBSV_XLSX_PATH, filename="PBSV.xlsx")
        excel_tx_all.extend(txs)
        processed_file_names.append("PBSV.xlsx")
    if os.path.exists(DEFAULT_PDF_PATH):
        _, txs = etl.parse_pdf_data(DEFAULT_PDF_PATH, filename="RE1002 T7 LMH.pdf")
        pdf_tx_all.extend(txs)
        processed_file_names.append("RE1002 T7 LMH.pdf")

if not is_direct_model and not excel_tx_all and not pdf_tx_all:
    st.warning("⚠️ Vui lòng tải lên ít nhất 1 file (.pdf, .xlsx, .xls) hoặc tích chọn 'Sử dụng file mẫu Demo'.")
    st.stop()

# Build relational tables
if is_direct_model and db_tables_raw:
    db_tables = db_tables_raw
else:
    db_tables = etl.build_relational_database(excel_tx_all, pdf_tx_all)

master_df = etl.build_master_flat_table(db_tables)

# ────────────────────────────────────────────
# Metrics bar
# ────────────────────────────────────────────
df_gd = db_tables.get('Giao_dich', pd.DataFrame())
total_tx  = len(df_gd)
total_val = pd.to_numeric(df_gd.get('Giá trị giao dịch', pd.Series(dtype=float)), errors='coerce').sum() if not df_gd.empty else 0
total_cust  = len(db_tables.get('Khach_hang', pd.DataFrame()))
total_stock = len(db_tables.get('Co_phieu', pd.DataFrame()))
total_files = len(processed_file_names)

cols = st.columns(5)
for c, val, lbl in zip(cols,
    [total_files, total_tx, f"{total_val:,.0f} đ", total_cust, total_stock],
    ["File đã xử lý", "Tổng giao dịch", "Tổng giá trị", "Khách hàng", "Mã cổ phiếu"]):
    c.markdown(f'<div class="metric-card"><div class="metric-value">{val}</div><div class="metric-label">{lbl}</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if processed_file_names:
    mode_tag = "📊 Multi-sheet Data Model" if is_direct_model else "🔄 ETL từ báo cáo"
    st.success(f"✅ {mode_tag} — Đã xử lý: **{', '.join(processed_file_names)}**")

# ────────────────────────────────────────────
# SESSION STATE: store editable tables
# ────────────────────────────────────────────
# We store edited versions in session_state so edits persist across rerenders.
if "edited_tables" not in st.session_state:
    st.session_state.edited_tables = {}
if "edited_master" not in st.session_state:
    st.session_state.edited_master = None

# Initialize session state on first load or when new data arrives
file_key = ",".join(processed_file_names) + str(is_direct_model)
if st.session_state.get("_last_file_key") != file_key:
    st.session_state.edited_tables = {k: v.copy() for k, v in db_tables.items()}
    st.session_state.edited_master = master_df.copy()
    st.session_state["_last_file_key"] = file_key

# ────────────────────────────────────────────
# Download helpers — always use edited state
# ────────────────────────────────────────────
def get_master_csv():
    return st.session_state.edited_master.to_csv(index=False).encode('utf-8-sig')

def get_master_xlsx():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        st.session_state.edited_master.to_excel(writer, sheet_name="Master_Data_Model", index=False)
        for tbl_name, df_tbl in st.session_state.edited_tables.items():
            df_tbl.to_excel(writer, sheet_name=tbl_name[:31], index=False)
    return buf.getvalue()

def get_zip_csv():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for tbl_name, df_tbl in st.session_state.edited_tables.items():
            zf.writestr(f"{tbl_name}.csv", df_tbl.to_csv(index=False).encode('utf-8-sig'))
    return buf.getvalue()

def get_sql():
    return etl.generate_sql_script(st.session_state.edited_tables).encode('utf-8')

# Quick download bar
st.markdown('<div class="dl-box"><b style="color:#38bdf8">📥 Tải Nhanh File Kết Quả</b> &nbsp;—&nbsp; <span style="color:#94a3b8;font-size:13px">Phản ánh mọi thay đổi bạn đã chỉnh sửa bên dưới</span></div>', unsafe_allow_html=True)

dc1, dc2, dc3, dc4 = st.columns(4)
with dc1:
    st.download_button("📄 CSV Tổng Hợp", get_master_csv(), "Master_Data_Model.csv", "text/csv", use_container_width=True)
with dc2:
    st.download_button("📊 Excel Multi-Sheet", get_master_xlsx(), "Master_Data_Model.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with dc3:
    st.download_button("📦 ZIP CSVs", get_zip_csv(), "Database_CSVs.zip", "application/zip", use_container_width=True)
with dc4:
    st.download_button("📜 SQL Script", get_sql(), "supabase_schema_data.sql", "text/plain", use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# ────────────────────────────────────────────
# Tabs
# ────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "✏️ 1. Chỉnh Sửa Bảng Master (Data Model)",
    "✏️ 2. Chỉnh Sửa Từng Bảng Riêng (Theo Thực Thể)",
    "🔍 3. Chi Tiết Trích Xuất Thô",
    "📐 4. Sơ Đồ Data Model (ERD)",
    "☁️ 5. Hướng Dẫn Deploy"
])

# ─── TAB 1: Master editable view ───
with tab1:
    st.subheader("✏️ Bảng Tổng Hợp Master — Click Vào Ô Để Chỉnh Sửa")
    st.markdown('<div class="edit-hint">💡 Click vào bất kỳ ô nào trong bảng bên dưới để sửa giá trị. Thay đổi sẽ được phản ánh ngay vào file tải xuống.</div>', unsafe_allow_html=True)

    edited_master = st.data_editor(
        st.session_state.edited_master,
        use_container_width=True,
        num_rows="dynamic",      # Allow adding/deleting rows
        key="master_editor"
    )
    # Persist edits to session_state
    st.session_state.edited_master = edited_master

    st.markdown(f"**{len(edited_master.columns)} cột** | **{len(edited_master)} dòng**")

# ─── TAB 2: Individual entity editor ───
with tab2:
    st.subheader("✏️ Chỉnh Sửa Từng Bảng Dữ Liệu Theo Thực Thể")
    st.markdown('<div class="edit-hint">💡 Chọn bảng từ dropdown, chỉnh sửa trực tiếp trên lưới, sau đó nhấn nút Lưu Thay Đổi để áp dụng.</div>', unsafe_allow_html=True)

    available_tables = {k: v for k, v in st.session_state.edited_tables.items() if not v.empty}
    selected_table = st.selectbox(
        "Chọn bảng cần chỉnh sửa:",
        list(available_tables.keys()),
        index=min(8, len(available_tables)-1) if available_tables else 0
    )

    if selected_table and selected_table in st.session_state.edited_tables:
        edited_tbl = st.data_editor(
            st.session_state.edited_tables[selected_table],
            use_container_width=True,
            num_rows="dynamic",
            key=f"editor_{selected_table}"
        )

        save_col, dl_col = st.columns([1, 2])
        with save_col:
            if st.button("💾 Lưu thay đổi bảng này", type="primary", use_container_width=True):
                st.session_state.edited_tables[selected_table] = edited_tbl
                # Rebuild master from updated tables
                st.session_state.edited_master = etl.build_master_flat_table(st.session_state.edited_tables)
                st.success(f"✅ Đã lưu thay đổi cho bảng `{selected_table}`!")

        with dl_col:
            st.download_button(
                label=f"⬇️ Tải CSV bảng `{selected_table}`",
                data=edited_tbl.to_csv(index=False).encode('utf-8-sig'),
                file_name=f"{selected_table}.csv",
                mime="text/csv",
                use_container_width=True
            )

# ─── TAB 3: Raw extracted data ───
with tab3:
    st.subheader("🔍 Dữ Liệu Thô Trích Xuất")
    if excel_tx_all:
        st.markdown("#### 📄 Dữ liệu từ File Excel")
        st.dataframe(pd.DataFrame(excel_tx_all), use_container_width=True)
    if pdf_tx_all:
        st.markdown("#### 📑 Dữ liệu từ File PDF")
        st.dataframe(pd.DataFrame(pdf_tx_all), use_container_width=True)
    if is_direct_model:
        st.info("📊 File này là Data Model nhiều sheet — không qua bước ETL trung gian. Dữ liệu gốc được hiển thị trực tiếp ở Tab 2.")

# ─── TAB 4: ERD ───
with tab4:
    st.subheader("📐 Sơ đồ Schema Relational Data Model")
    if os.path.exists(DEFAULT_IMG_PATH):
        try:
            st.image(DEFAULT_IMG_PATH, caption="Relational Database Model Schema (DATA MODEL.png)", use_container_width=True)
        except Exception:
            st.image(DEFAULT_IMG_PATH, caption="Relational Database Model Schema (DATA MODEL.png)")
    else:
        st.info("Sơ đồ DATA MODEL.png (chỉ hiển thị khi chạy tại local).")

# ─── TAB 5: Deploy guide ───
with tab5:
    st.subheader("☁️ Hướng Dẫn Deploy Streamlit Web Tool Công Khai")
    st.markdown("""
    1. Upload các file (`app.py`, `etl_processor.py`, `requirements.txt`, `DATA MODEL.png`) lên **GitHub**.
    2. Đăng nhập [share.streamlit.io](https://share.streamlit.io) bằng tài khoản GitHub.
    3. Chọn **New app** → chọn repository → nhập `app.py` → bấm **Deploy!**
    
    Sau 1-2 phút bạn sẽ có link công khai chia sẻ cho mọi người dùng trực tuyến!
    """)

st.markdown("---")
st.caption("⚡ Financial Data ETL Tool v2.0 | Streamlit Web Engine.")
