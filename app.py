# -*- coding: utf-8 -*-
"""
=====================================================================
app.py - STREAMLIT TOOL: CHUẨN HÓA & GỘP BÁO CÁO GIAO DỊCH JB + PBSV
=====================================================================
File này đã được nâng cấp cấu trúc giao diện (Sidebar + Cards + Tabs)
và chèn đầy đủ comment hướng dẫn chi tiết từng thành phần giao diện.
"""

import io
import json

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

import merge_processor as mp

# =====================================================================
# CẤU TRÚC 1: CAU HINH TRANG & BỐ CỤC TỔNG THỂ (PAGE CONFIG)
# =====================================================================
# st.set_page_config phải là lệnh Streamlit đầu tiên trong file.
# - page_title: Tiêu đề hiển thị trên Tab trình duyệt
# - page_icon: Biểu tượng Favicon trên tab trình duyệt
# - layout: "wide" (rộng toàn màn hình) hoặc "centered" (thu gọn ở giữa)
# - initial_sidebar_state: "expanded" (mở sidebar) hoặc "collapsed" (thu gọn)
st.set_page_config(
    page_title="Gộp Báo Cáo Giao Dịch JB + PBSV",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =====================================================================
# CẤU TRÚC 2: CUSTOM CSS (LÀM ĐẸP GIAO DIỆN SÂU BẰNG CSS)
# =====================================================================
# Bạn có thể tự do thêm/sửa mã CSS bên dưới để tùy biến font chữ, shadow, màu card,...
st.markdown(
    """
    <style>
    /* 1. Đội kiểu dáng Banner Header chính */
    .main-header {
        background: linear-gradient(135deg, #0052D4 0%, #4364F7 50%, #6FB1FC 100%);
        padding: 20px 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(0, 82, 212, 0.15);
    }
    .main-header h1 {
        color: white !important;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
    }
    .main-header p {
        color: #E0E7FF !important;
        margin: 6px 0 0 0;
        font-size: 0.95rem;
    }
    
    /* 2. Đổi màu và kiểu dáng cho các thẻ Metric Thống kê */
    [data-testid="stMetric"] {
        background-color: var(--background-color);
        padding: 14px 18px;
        border-radius: 10px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 2px 5px rgba(0,0,0,0.02);
    }
    [data-testid="stMetricValue"] {
        font-size: 1.6rem !important;
        font-weight: 700;
        color: #0066CC !important;
    }
    
    /* 3. Tùy chỉnh Nút bấm Primary (Xử lý) trong Sidebar */
    .stButton > button[kind="primary"] {
        border-radius: 8px;
        font-weight: 600;
        font-size: 1rem;
        padding: 10px 16px;
        box-shadow: 0 4px 12px rgba(0, 102, 204, 0.25);
    }
    </style>
""",
    unsafe_allow_html=True,
)

# =====================================================================
# CẤU TRÚC 3: THANH SIDEBAR BÊN TRÁI (BẢNG ĐIỀU KHIỂN & TẢI FILE)
# =====================================================================
# Sử dụng `with st.sidebar:` để đưa toàn bộ khu vực Cấu hình & Tải file 
# sang thanh bên trái, giải phóng không gian màn hình chính cho Bảng Excel.
with st.sidebar:
    st.image("https://img.icons8.com/color/96/analytics.png", width=60)
    st.title("⚙️ Bảng Điều Khiển")
    st.caption("Tải lên các file báo cáo để bắt đầu xử lý gộp dữ liệu.")

    st.divider()

    # --- KHU VỰC UPLOAD FILE ---
    st.subheader("1️⃣ File Đầu Vào")

    jb_files = st.file_uploader(
        "📄 File JB input (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="jb_uploader",
        accept_multiple_files=True,
        help="Báo cáo 'Kết quả khớp lệnh và phí giao dịch' từ JB. Có thể chọn nhiều file.",
    )

    pbsv_files = st.file_uploader(
        "📄 File PBSV input (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="pbsv_uploader",
        accept_multiple_files=True,
        help="Báo cáo 'Lịch sử đặt lệnh' từ PBSV. Chỉ giữ giao dịch 'Hoàn thành'. Có thể chọn nhiều file.",
    )

    master_files = st.file_uploader(
        "👥 DS Khách hàng (2 sheet JB & PB)",
        type=["xlsx", "xls"],
        key="master_uploader",
        accept_multiple_files=True,
        help="Cần có 2 sheet: 1 sheet chứa 'JB' và 1 sheet chứa 'PB'.",
    )

    st.divider()

    # --- NÚT XỬ LÝ (PRIMARY BUTTON) ---
    # Nút bấm được đặt dạng use_container_width=True để full chiều rộng sidebar
    process_clicked = st.button("🚀 Bắt đầu Xử lý", type="primary", use_container_width=True)

    # --- TRỢ GIÚP / HƯỚNG DẪN Ở SIDEBAR ---
    with st.expander("❓ Hướng dẫn sử dụng"):
        st.markdown(
            """
        1. Tải lên **DS Khách hàng** và ít nhất **1 file giao dịch** (JB hoặc PBSV).
        2. Nhấn nút **🚀 Bắt đầu Xử lý**.
        3. Xem thống kê và chỉnh sửa dữ liệu trực tiếp ở màn hình chính.
        4. Chuyển sang Tab **Tải xuống** để tải file Excel/CSV.
        """
        )

# =====================================================================
# CẤU TRÚC 4: MÀN HÌNH CHÍNH (MAIN CONTENT AREA)
# =====================================================================

# --- 4.1. BANNER TIÊU ĐỀ CHÍNH ---
st.markdown(
    """
    <div class="main-header">
        <h1>📊 Hệ Thống Chuẩn Hóa & Gộp Báo Cáo Giao Dịch</h1>
        <p>Tự động đối soát, khớp danh sách khách hàng và xuất báo cáo chuẩn JB + PBSV</p>
    </div>
""",
    unsafe_allow_html=True,
)

# --- 4.2. QUẢN LÝ TRẠNG THÁI (SESSION STATE) ---
# Session state giúp giữ nguyên kết quả xử lý ngay cả khi người dùng tương tác với bảng
if "result" not in st.session_state:
    st.session_state.result = None
if "edited_output" not in st.session_state:
    st.session_state.edited_output = None
if "edited_mismatch" not in st.session_state:
    st.session_state.edited_mismatch = None

# --- 4.3. XỬ LÝ SỰ KIỆN KHI BẤM NÚT "XỬ LÝ" ---
if process_clicked:
    if not master_files:
        st.warning("⚠️ Vui lòng tải lên file Danh sách khách hàng (Master) ở thanh Sidebar trước khi bấm Xử lý.")
    elif not (jb_files or pbsv_files):
        st.warning("⚠️ Vui lòng tải lên ít nhất 1 file giao dịch đầu vào (JB hoặc PBSV) ở thanh Sidebar.")
    else:
        with st.spinner("⏳ Đang chuẩn hóa dữ liệu và xử lý gộp file..."):
            result = mp.process_files(
                [(f, f.name) for f in jb_files] if jb_files else [],
                [(f, f.name) for f in pbsv_files] if pbsv_files else [],
                [(f, f.name) for f in master_files] if master_files else [],
            )
        st.session_state.result = result
        st.session_state.edited_output = (
            result["output_df"].copy() if result["output_df"] is not None else None
        )
        st.session_state.edited_mismatch = result["mismatch_df"].copy()

# --- 4.4. HIỂN THỊ KHI CHƯA CÓ DỮ LIỆU ---
result = st.session_state.result

if result is None:
    # Sử dụng st.container với border=True để tạo Thẻ thông báo chào mừng
    with st.container(border=True):
        st.info("👈 Hãy chọn các file đầu vào từ **Thanh Sidebar bên trái** và bấm **🚀 Bắt đầu Xử lý**.")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.markdown("##### 📁 File JB Input")
            st.caption("Báo cáo khớp lệnh & phí từ hệ thống JB.")
        with col_b:
            st.markdown("##### 📁 File PBSV Input")
            st.caption("Lịch sử đặt lệnh đã hoàn thành từ PBSV.")
        with col_c:
            st.markdown("##### 👥 Danh sách Khách hàng")
            st.caption("File Excel gồm 2 sheet khách hàng JB và PB.")
    st.stop()

# --- 4.5. XỬ LÝ BÁO LỖI VÀ CẢNH BÁO ---
if result["errors"]:
    st.error("❌ **Lỗi cấu trúc file nghiêm trọng:**")
    for err in result["errors"]:
        st.markdown(f"- {err}")
    st.stop()

if result.get("file_errors"):
    st.error(f"❌ Có **{len(result['file_errors'])} file bị lỗi định dạng** (các file hợp lệ khác vẫn được xử lý):")
    for err in result["file_errors"]:
        st.markdown(f"- {err}")

if result["warnings"]:
    with st.expander(f"⚠️ {len(result['warnings'])} cảnh báo chi tiết trong quá trình gộp dữ liệu", expanded=False):
        for w in result["warnings"]:
            st.markdown(f"- {w}")

# --- HÀM HỖ TRỢ CHUYỂN ĐỔI DỮ LIỆU ĐÃ CHỈNH SỬA RA BYTES ĐỂ TẢI DOWN ---
def _to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def _to_xlsx_bytes(df: pd.DataFrame, mismatch_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Output", index=False)
        if len(mismatch_df) > 0:
            mismatch_df.to_excel(writer, sheet_name="Khong_khop", index=False)
    return buf.getvalue()


def _to_tsv_string(df: pd.DataFrame, use_vn_decimal: bool = True, include_header: bool = False) -> str:
    """Chuyển đổi DataFrame thành chuỗi Tab-Separated Values (TSV) để dán chuẩn vào Google Sheets bản tiếng Việt.
    Dùng dấu phẩy ',' cho phần thập phân (vd: 0,75) để không bị Google Sheets hiểu nhầm dấu '.' là phân cách nghìn và nhảy thành 75.
    Mặc định include_header=False để loại bỏ dòng tiêu đề khi copy.
    """
    df_clean = df.copy()
    num_cols = ["Khối lượng đặt", "Khối lượng khớp", "Giá khớp", "Tổng giá ", "% Phí", "Phí", "Thuế", "Tổng thuần"]
    for col in num_cols:
        if col in df_clean.columns:
            s_num = pd.to_numeric(df_clean[col], errors="coerce").fillna(0.0)

            def _format_num(x):
                if pd.isna(x):
                    return ""
                fx = float(x)
                if fx.is_integer():
                    return str(int(fx))
                formatted = f"{fx:.4f}".rstrip("0").rstrip(".")
                if use_vn_decimal:
                    formatted = formatted.replace(".", ",")
                return formatted

            df_clean[col] = s_num.apply(_format_num)

    return df_clean.to_csv(sep="\t", index=False, header=include_header)


def copy_to_clipboard_button(text_data: str, label: str = "📋 Copy Dữ Liệu Cho Google Sheets", key: str = "copy_btn"):
    """Tạo nút JavaScript copy dữ liệu thẳng vào Clipboard trình duyệt."""
    js_data = json.dumps(text_data)
    html_code = f"""
    <div style="font-family: system-ui, -apple-system, sans-serif;">
        <button id="{key}" onclick="copyData_{key}()" style="
            width: 100%;
            background: linear-gradient(135deg, #107c41 0%, #1f9a55 100%);
            color: white;
            border: none;
            padding: 10px 16px;
            font-size: 14px;
            font-weight: 600;
            border-radius: 8px;
            cursor: pointer;
            box-shadow: 0 3px 8px rgba(16, 124, 65, 0.25);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            transition: all 0.2s ease;
        ">
            {label}
        </button>
        <div id="status_{key}" style="margin-top: 6px; font-size: 12px; color: #107c41; font-weight: 600; text-align: center; display: none;">
            ✅ Đã copy dữ liệu! Mở Google Sheets và bấm Ctrl + V để dán.
        </div>
    </div>
    <script>
    function copyData_{key}() {{
        const text = {js_data};
        navigator.clipboard.writeText(text).then(function() {{
            const status = document.getElementById("status_{key}");
            const btn = document.getElementById("{key}");
            const oldBg = btn.style.background;
            const oldText = btn.innerHTML;
            btn.style.background = "#198754";
            btn.innerHTML = "✅ Đã Copy!";
            status.style.display = "block";
            setTimeout(function() {{
                btn.style.background = oldBg;
                btn.innerHTML = oldText;
                status.style.display = "none";
            }}, 3500);
        }}).catch(function(err) {{
            alert("Lỗi sao chép dữ liệu: " + err);
        }});
    }}
    </script>
    """
    components.html(html_code, height=70)


def push_to_google_sheet_webhook(df: pd.DataFrame, webhook_url: str) -> tuple[bool, str]:
    """Gửi dữ liệu DataFrame qua Webhook Google Apps Script để tự động cập nhật Google Sheet dưới dạng KIỂU SỐ (Numbers)."""
    try:
        import numpy as np
        headers = df.columns.tolist()

        def _clean_val(v):
            if pd.isna(v) or v is None:
                return ""
            if isinstance(v, (int, float, np.integer, np.floating)):
                if np.isnan(v) or np.isinf(v):
                    return 0
                if float(v).is_integer():
                    return int(v)
                return float(v)
            s = str(v).strip()
            if "," in s and "." not in s:
                try:
                    return float(s.replace(",", "."))
                except ValueError:
                    pass
            try:
                val = float(s)
                if val.is_integer():
                    return int(val)
                return val
            except ValueError:
                return s

        rows = [[_clean_val(cell) for cell in row] for row in df.values]
        payload = {
            "headers": headers,
            "rows": rows
        }
        response = requests.post(webhook_url, json=payload, timeout=30)
        if response.status_code == 200:
            return True, "🎉 Tự động đẩy dữ liệu sang Google Sheet thành công!"
        else:
            return False, f"⚠️ Lỗi phản hồi từ Google Apps Script (HTTP {response.status_code}): {response.text}"
    except Exception as e:
        return False, f"❌ Lỗi kết nối Webhook: {str(e)}"





# =====================================================================
# CẤU TRÚC 5: KHU VỰC THỐNG KÊ DASHBOARD (METRICS CARD CONTAINER)
# =====================================================================
# Gom các số liệu KPI vào một Card container có khung viền bo góc đẹp mắt
with st.container(border=True):
    st.markdown("#### 📈 Thống Kê Tổng Quan Dữ Liệu")
    stats = result["stats"]
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Giao dịch JB", f"{stats.get('so_gd_jb', 0):,} dòng")
    s2.metric("GD PBSV (Hoàn thành)", f"{stats.get('so_gd_pbsv', 0):,} dòng")
    s3.metric("Tổng dòng Output", f"{stats.get('tong_so_gd', 0):,} dòng")
    s4.metric("Khách chưa khớp", f"{stats.get('so_dong_khong_khop', 0):,} dòng", delta_color="inverse")

st.markdown("<br>", unsafe_allow_html=True)

# =====================================================================
# CẤU TRÚC 6: PHÂN CHIA TAB CÔNG VIỆC (MULTI-TAB WORKFLOW)
# =====================================================================
# Chia giao diện hiển thị kết quả thành 4 Tab chuyên biệt
tab_main, tab_mismatch, tab_export, tab_gsheet = st.tabs(
    [
        "📋 Bảng Kết Quả Chính",
        "⚠️ Dòng Không Khớp Khách Hàng",
        "📥 Tải Xuất Báo Cáo",
        "🟢 Đồng Bộ Google Sheets",
    ]
)

# --- TAB 1: BẢNG DỮ LIỆU CHÍNH (DẠNG EXCEL EDITABLE) ---
with tab_main:
    st.subheader("2️⃣ Xem & Chỉnh Sửa Báo Cáo Dữ Liệu")
    st.caption("💡 Bạn có thể click đúp chuột vào bất kỳ ô nào để chỉnh sửa trực tiếp dữ liệu trước khi tải xuống.")

    # st.data_editor cho phép sửa dữ liệu dạng Excel
    edited_output = st.data_editor(
        st.session_state.edited_output,
        use_container_width=True,
        num_rows="dynamic",
        key="output_editor",
    )
    st.session_state.edited_output = edited_output

    st.success("✅ **Đã cập nhật:** File tải về / copy bên dưới luôn là **dữ liệu mới nhất sau khi bạn sửa tay**.")

    # Khung chứa nút tải nhanh và nút copy trực tiếp tại Tab 1
    with st.container(border=True):
        col_down1, col_down2, col_down3 = st.columns(3)
        with col_down1:
            st.download_button(
                "⬇️ Tải xuống file CSV (.csv)",
                data=_to_csv_bytes(st.session_state.edited_output),
                file_name="output_giao_dich_da_sua.csv",
                mime="text/csv",
                use_container_width=True,
                key="btn_dl_csv_tab1",
            )
        with col_down2:
            st.download_button(
                "⬇️ Tải xuống file Excel (.xlsx)",
                data=_to_xlsx_bytes(st.session_state.edited_output, st.session_state.edited_mismatch),
                file_name="output_giao_dich_da_sua.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_dl_xlsx_tab1",
            )
        with col_down3:
            copy_to_clipboard_button(
                _to_tsv_string(st.session_state.edited_output),
                label="📋 Copy Dán Vào Google Sheets",
                key="btn_copy_tab1",
            )

# --- TAB 2: CÁC DÒNG LỖI / KHÔNG KHỚP KHÁCH HÀNG ---
with tab_mismatch:
    if len(st.session_state.edited_mismatch) > 0:
        st.subheader("⚠️ Danh Sách Giao Dịch Không Khớp Khách Hàng")
        st.caption("Các dòng này chưa tìm thấy thông tin khách hàng tương ứng trong file Master. Bạn có thể bổ sung thủ công tại đây.")

        edited_mismatch = st.data_editor(
            st.session_state.edited_mismatch,
            use_container_width=True,
            num_rows="dynamic",
            key="mismatch_editor",
        )
        st.session_state.edited_mismatch = edited_mismatch
    else:
        st.success("🎉 **Tuyệt vời!** Tất cả các giao dịch đều đã khớp hoàn toàn với danh sách khách hàng.")

# --- TAB 3: TẢI FILE KẾT QUẢ (EXPORT OPTIONS) ---
with tab_export:
    st.subheader("3️⃣ Tải Xuất File Báo Cáo Dữ Liệu")
    st.caption("Chọn định dạng file bạn mong muốn để tải về máy tính (bao gồm tất cả chỉnh sửa tay mới nhất).")

    # Đặt nút tải file trong khung Card viền đẹp mắt
    with st.container(border=True):
        col_d1, col_d2 = st.columns(2)

        with col_d1:
            st.markdown("##### 📄 Định dạng CSV (.csv)")
            st.caption("Dữ liệu bảng chính (sau khi chỉnh sửa) dạng CSV UTF-8.")
            st.download_button(
                "⬇️ Tải File CSV",
                data=_to_csv_bytes(st.session_state.edited_output),
                file_name="output_giao_dich.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with col_d2:
            st.markdown("##### 📊 Định dạng Excel (.xlsx)")
            st.caption("Bao gồm cả Sheet Output chính và Sheet Các dòng không khớp.")
            st.download_button(
                "⬇️ Tải File Excel (.xlsx)",
                data=_to_xlsx_bytes(st.session_state.edited_output, st.session_state.edited_mismatch),
                file_name="output_giao_dich.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# --- TAB 4: ĐỒNG BỘ GOOGLE SHEETS & COPY NHANH ---
with tab_gsheet:
    st.subheader("🟢 Đồng Bộ Dữ Liệu Trực Tiếp Sang Google Sheets")
    st.caption("Sao chép dữ liệu để dán nhanh hoặc tự động ghi trực tiếp vào Google Sheet mà không cần tải file về máy.")

    c1, c2 = st.columns(2)

    with c1:
        with st.container(border=True):
            st.markdown("##### 📋 Phương án 1: Copy Nhanh 1-Click")
            st.caption("Copy dữ liệu (đã bỏ dòng tiêu đề/header) dưới dạng bảng TSV. Mở Google Sheet và nhấn **Ctrl + V** tại ô muốn dán.")
            copy_to_clipboard_button(
                _to_tsv_string(st.session_state.edited_output),
                label="📋 Copy Dữ Liệu Dán Vào Google Sheets",
                key="btn_copy_tab_gsheet",
            )

    with c2:
        with st.container(border=True):
            st.markdown("##### 🚀 Phương án 2: Tự Động Đẩy Sang Google Sheet")
            st.caption("Nhập URL Webhook (Apps Script) của Google Sheet để tự động ghi đè dữ liệu.")
            webhook_url = st.text_input(
                "🔗 URL Google Apps Script Webhook",
                placeholder="https://script.google.com/macros/s/.../exec",
                key="gsheet_webhook_url_input",
            )
            if st.button("🚀 Bắt đầu Đẩy Dữ Liệu sang Google Sheet", type="primary", use_container_width=True):
                if not webhook_url:
                    st.warning("⚠️ Vui lòng nhập Webhook URL của Google Apps Script trước!")
                else:
                    with st.spinner("⏳ Đang kết nối và đẩy dữ liệu sang Google Sheet..."):
                        success, msg = push_to_google_sheet_webhook(st.session_state.edited_output, webhook_url)
                    if success:
                        st.success(msg)
                    else:
                        st.error(msg)

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📖 Hướng dẫn thiết lập Tự động Đẩy Dữ liệu sang Google Sheet (1 phút)"):
        st.markdown(
            """
            **Bước 1:** Mở file Google Sheet của bạn -> Trên thanh menu chọn **Extensions (Tiện ích mở rộng)** -> **Apps Script**.
            
            **Bước 2:** Xóa hết code cũ, dán đoạn mã bên dưới vào và nhấn biểu tượng **Lưu (Save 💾)**:
            """
        )
        st.code(
            """function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Xóa nội dung cũ
    sheet.clearContents();
    
    // Ghi tiêu đề
    if (data.headers && data.headers.length > 0) {
      sheet.appendRow(data.headers);
    }
    
    // Ghi các dòng dữ liệu (tự động ép kiểu chuỗi số thành Number chuẩn)
    if (data.rows && data.rows.length > 0) {
      var cleanedRows = data.rows.map(function(row) {
        return row.map(function(val) {
          if (typeof val === 'string') {
            var trimmed = val.trim();
            if (trimmed !== '' && !isNaN(trimmed)) {
              return Number(trimmed);
            }
          }
          return val;
        });
      });
      sheet.getRange(2, 1, cleanedRows.length, cleanedRows[0].length).setValues(cleanedRows);
    }
    
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}""",
            language="javascript",
        )

        st.markdown(
            """
            **Bước 3:** Nhấn nút **Deploy (Triển khai)** góc trên bên phải -> Chọn **New deployment (Triển khai mới)**.
            - Select type (Chọn loại): **Web app**.
            - Execute as: **Me**.
            - Who has access (Ai có quyền truy cập): **Anyone (Bất kỳ ai)**.
            - Nhấn **Deploy** -> Cấp quyền nếu Google hỏi -> Copy đường link **Web App URL** dán vào ô bên trên.
            """
        )

