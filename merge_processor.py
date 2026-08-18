# -*- coding: utf-8 -*-
"""
merge_processor.py
===================
Chuan hoa va gop 2 bao cao giao dich (JB, PBSV) thanh 1 bang duy nhat,
theo dung cau truc file output chuan, co doi chieu voi danh sach khach hang
de dien cac truong con thieu (Tai khoan luu ky, Ten khach hang, Nguoi quan ly...).

Quy trinh:
    1. Doc va chuan hoa JB input     -> parse_jb_input()
    2. Doc va chuan hoa PBSV input   -> parse_pbsv_input() (chi giu "Hoan thanh")
    3. Doc danh sach khach hang JB/PB (2 sheet) -> parse_customer_master()
    4. Join theo khoa + dien truong thieu -> join_and_normalize()
    5. Tinh cac truong dan xuat ("Tổng thuần", % Phí neu con thieu)
       theo CONG THUC CHUAN DUY NHAT: sau = truoc + Phi + Thue (ap dung cho moi dong,
       khong phan biet nguon JB/PBSV hay Mua/Ban) -> _fill_derived_fields()
    6. Gop JB + PBSV -> build_output()
"""
from __future__ import annotations

import re
import io
from datetime import datetime
from typing import Optional

import numpy as np
import pandas as pd

# ============================================================
# Ten 2 cot gia tri giao dich (truoc/sau phi + thue) - dung dung ten cot nhu trong file
# mau Output.xlsx do user cung cap (nguon chuan duy nhat cho schema output).
# CONG THUC CHUAN DUY NHAT (ap dung cho MOI dong, khong phan biet nguon JB/PBSV
# hay Mua/Ban):
#     COL_SAU = COL_TRUOC + Phí + Thuế
# Xem _fill_derived_fields() de biet chi tiet + buoc validate.
# ============================================================
COL_TRUOC = "Tổng giá "
COL_SAU = "Tổng thuần"

# Gia tri mac dinh dien vao khi khong tra cuu duoc "Nhóm khách hàng" / "Phân loại khách
# hàng" tu danh sach khach hang (o trong, None, NaN, hoac chi chua khoang trang).
DEFAULT_UNKNOWN = "Không xác định"

# ============================================================
# Cau truc output chuan (khop file mau Output.xlsx do user cung cap)
# Luu y: cot "Tai khoan luu ky " co dau cach cuoi la co chu y de khop chinh xac
# template - khong duoc strip().
# ============================================================
OUTPUT_COLUMNS = [
    "Ngày",
    "Số hiệu lệnh",
    "Tài khoản lưu ký ",
    "Số tiểu khoản",
    "Tên khách hàng",
    "Mã Chứng khoán",
    "Giao dịch (Mua/Bán)",
    "Khối lượng đặt",
    "Khối lượng khớp",
    "Giá khớp",
    COL_TRUOC,
    "% Phí",
    "Phí",
    "Thuế",
    COL_SAU,
    "Tên người quản lý",
    "Tên Công ty",
    "Nhóm khách hàng",
    "Phân loại khách hàng",
]

MISMATCH_COLUMNS = [
    "Nguồn",
    "Số tiểu khoản tra cứu",
    "Tên khách hàng (nếu có)",
    "Lý do",
]

COMPANY_NAME_FALLBACK = {
    "JB": "JB",
    "PBSV": "PBSV",
}

class ValidationError(Exception):
    """Loi format / thieu cot bat buoc o file dau vao - hien thi ro cho user, khong crash app."""

# ------------------------------------------------------------------
# Helpers dung chung & Tu khoa nhan dien Cong ty / To chuc
# ------------------------------------------------------------------

COMPANY_KEYWORDS_SUBSTRING = [
    "CÔNG TY", "TỔNG CÔNG TY", "TẬP ĐOÀN", "DOANH NGHIỆP", "CHI NHÁNH",
    "TRUNG TÂM", "HIỆP HỘI", "NGÂN HÀNG", "BẢO HIỂM", "CHỨNG KHOÁN",
    "QUỸ ĐẦU TƯ", "QUỸ MỞ", "QUỸ THÀNH VIÊN", "QUỸ HỘI", "QUỸ", "TỔ CHỨC",
    "HỢP TÁC XÃ", "BAN QUẢN LÝ", "SỞ GIAO DỊCH", "VIỆN NGHỊÊN CỨU", "VIỆN",
    "LIÊN DOANH", "THƯƠNG MẠI CỔ PHẦN", "THƯƠNG MẠI DỊCH VỤ", "CỔ PHẦN",
    "DOANH NGHIỆP TƯ NHÂN", "HOLDING", "HOLDINGS", "CORPORATION", "SECURITIES",
    "INVESTMENT", "INVESTMENTS", "MANAGEMENT", "CAPITAL", "PARTNERS",
    "LIMITED", "FINANCE", "FINANCIAL", "ENTERPRISE", "ASSET MANAGEMENT",
    "ASSET", "COMMERCIAL", "LOGISTICS", "REAL ESTATE", "TECHNOLOGY",
    "DEVELOPMENT", "SERVICES", "TRADING", "SOLUTIONS",
]

COMPANY_KEYWORDS_TOKEN = {
    "CTY", "TCTY", "CTCP", "TNHH", "TMCP", "NHTMCP", "NHTM", "JSC", "DNTN",
    "TCT", "HTX", "CN", "LLC", "PLC", "INC", "LTD", "CO", "CO.", "CORP",
    "GROUP", "BANK", "FUND", "FUNDS", "TRUST", "BV", "NV", "SA", "AG",
    "GMBH", "PTE", "SDN", "BHD", "SRO"
}

ROMAN_NUMERALS = {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"}


def is_company_or_organization(name: str) -> bool:
    """Kiem tra xem ten khach hang co phai la cong ty, to chuc, doanh nghiep hay khong."""
    if not name or not isinstance(name, str):
        return False
    upper_name = name.strip().upper()
    if not upper_name:
        return False

    # Kiem tra cac chuoi con (substring)
    for kw in COMPANY_KEYWORDS_SUBSTRING:
        if kw in upper_name:
            return True

    # Kiem tra cac token / tu doc lap
    tokens = set(re.split(r'[\s,.\-/()]+', upper_name))
    for tk in COMPANY_KEYWORDS_TOKEN:
        if tk in tokens:
            return True

    return False


def format_customer_name(name: str) -> str:
    """Chuan hoa ten khach hang:
    Neu ten bi full IN HOA va KHONG PHAI cong ty, to chuc -> chuyen ve Title Case (Nguyễn Văn A).
    Neu la cong ty/to chuc hoac da co hoa thuong -> giu nguyen.
    """
    if not name or not isinstance(name, str):
        return name if name is not None else ""

    s = name.strip()
    if not s:
        return s

    # Chi dinh dang lai ten FULL IN HOA (s.isupper() == True)
    if s.isupper():
        if not is_company_or_organization(s):
            formatted = s.title()

            # Giu nguyen cac ma La Ma (I, II, III, IV, ...)
            words = formatted.split()
            new_words = []
            for w in words:
                clean_w = re.sub(r'[^\w]', '', w).upper()
                if clean_w in ROMAN_NUMERALS:
                    new_words.append(w.upper())
                else:
                    new_words.append(w)
            return " ".join(new_words)

    return s


def _norm_text(v) -> str:
    """Chuan hoa 1 gia tri o (cell) ve chuoi da strip, phuc vu so sanh/tim header."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _key_str(v) -> Optional[str]:
    """Chuan hoa gia tri dung lam khoa join / ma code (tieu khoan, ma CK...).
    Xu ly truong hop Excel luu so nguyen thanh float (vd 123.0 -> '123')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    return s if s else None


def _to_number(v) -> float:
    """Chuyen ve so (float), tra 0.0 neu rong/khong hop le (ho tro dinh dang chuoi Viet Nam & Quoc te)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)

    s = str(v).strip()
    if not s or s == "-":
        return 0.0

    # Chuan hoa phan phan cach nghin va phap phan
    if "." in s and "," in s:
        # Neu dot o truoc comma (vd: "36.800,50" hoac "3.009.000,00"): VN format
        if s.rfind(".") < s.rfind(","):
            s = s.replace(".", "").replace(",", ".")
        else: # EN format: "36,800.50"
            s = s.replace(",", "")
    elif "." in s:
        # Truong hop chi co dau cham:
        # Vd 1: "36.800.000" -> nhieu dau cham -> phan cach nghin VN
        # Vd 2: "36.800" hoac "9.200" -> phan sau co 3 chu so ngàn -> phan cach nghin VN
        # Vd 3: "1000000.0" hoac "0.15" -> standard decimal float -> giu nguyen
        parts = s.split(".")
        if len(parts) > 2:
            s = s.replace(".", "")
        elif len(parts) == 2:
            if len(parts[1]) == 3 and not parts[1].startswith("00") and parts[0] != "0":
                s = s.replace(".", "")
    elif "," in s:
        # Truong hop chi co dau phay:
        # Vd 1: "0,1" hoac "36,8" -> "0.1", "36.8" (decimal comma VN)
        # Vd 2: "1,000,000" -> (multiple commas -> thousand separator EN)
        parts = s.split(",")
        if len(parts) > 2:
            s = s.replace(",", "")
        elif len(parts) == 2:
            if len(parts[1]) == 3:
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")

    try:
        val = float(s)
        return val if not np.isnan(val) else 0.0
    except (ValueError, TypeError):
        return 0.0


def _format_date_cell(v) -> Optional[str]:
    """Chuan hoa 1 o ngay thang ve chuoi 'dd/mm/yyyy', ho tro ca cell dang datetime
    (Excel luu ngay dang so/date-time) lan cell dang chuoi co san."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime("%d/%m/%Y")
    return _norm_text(v) or None


def _read_raw_table(file, filename: str) -> pd.DataFrame:
    """Doc file tho (khong header) ho tro .xlsx/.xls/.csv. Khong raise cho file rong."""
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    try:
        if ext == "csv":
            raw = file.read() if hasattr(file, "read") else file
            if isinstance(raw, (bytes, bytearray)):
                text = raw.decode("utf-8-sig", errors="replace")
            else:
                text = raw
            return pd.read_csv(io.StringIO(text), header=None, dtype=object)
        elif ext in ("xlsx", "xlsm"):
            return pd.read_excel(file, header=None, dtype=object, engine="openpyxl")
        elif ext == "xls":
            return pd.read_excel(file, header=None, dtype=object, engine="xlrd")
        else:
            raise ValidationError(
                f"File '{filename}' không đúng định dạng được hỗ trợ (.xlsx, .xls, .csv)."
            )
    except ValidationError:
        raise
    except Exception as e:  # noqa: BLE001 - muon bat moi loi doc file de bao ro cho user
        raise ValidationError(f"Không đọc được file '{filename}': {e}") from e


def _find_row_with_marker(df: pd.DataFrame, markers, max_rows: int = 40) -> Optional[int]:
    """Tim dong dau tien (trong max_rows dong dau) co 1 o khop voi 1 trong cac 'markers'
    (so sanh khong phan biet hoa/thuong, da strip)."""
    markers_norm = {m.strip().lower() for m in markers}
    limit = min(max_rows, len(df))
    for i in range(limit):
        row = df.iloc[i]
        for v in row:
            if _norm_text(v).lower() in markers_norm:
                return i
    return None


def _extract_pattern(df: pd.DataFrame, pattern: str, max_rows: int = 15) -> Optional[str]:
    """Quet cac dong dau file de tim 1 cell khop regex, tra ve group(1)."""
    limit = min(max_rows, len(df))
    rx = re.compile(pattern)
    for i in range(limit):
        row = df.iloc[i]
        for v in row:
            text = _norm_text(v)
            if not text:
                continue
            m = rx.search(text)
            if m:
                return m.group(1).strip()
    return None


def _extract_company_name(df: pd.DataFrame, source: str) -> str:
    """Lay ten cong ty rut gon (JB, PBSV) cho dau ra cot 'Tên Công ty'."""
    return COMPANY_NAME_FALLBACK.get(source, source)


def _combine_two_row_header(df: pd.DataFrame, header_row1: int, header_row2: int) -> list[str]:
    """Ghep header 2 dong (vd 'Mua' o dong 1 + 'KL Đặt' o dong 2 -> 'Mua_KL Đặt').
    Dong 1 duoc forward-fill de ap dung cho cac cot con (merged cell)."""
    row1 = df.iloc[header_row1].tolist()
    row2 = df.iloc[header_row2].tolist()
    # forward-fill row1
    filled1 = []
    last = ""
    for v in row1:
        t = _norm_text(v)
        if t:
            last = t
        filled1.append(last)
    combined = []
    for c in range(len(row1)):
        l1 = filled1[c]
        l2 = _norm_text(row2[c])
        if l2:
            combined.append(f"{l1}_{l2}" if l1 else l2)
        else:
            combined.append(l1)
    return combined


# ============================================================
# 1) JB input
# ============================================================

# Ten cot ket hop (sau khi gop header 2 dong) -> ten noi bo
_JB_COL_MAP = {
    "STT": "stt",
    "SHL": "so_hieu_lenh",
    "Số tiểu khoản": "so_tieu_khoan",
    "Tên khách hàng": "ten_khach_hang",
    "Mã chứng khoán": "ma_ck",
    "Mua_KL Đặt": "mua_kl_dat",
    "Mua_KL khớp": "mua_kl_khop",
    "Mua_Giá khớp TB": "mua_gia_khop",
    "Mua_Giá trị": "mua_gia_tri",
    "Bán_KL Đặt": "ban_kl_dat",
    "Bán_KL khớp": "ban_kl_khop",
    "Bán_Giá khớp TB": "ban_gia_khop",
    "Bán_Giá trị": "ban_gia_tri",
    "Phí_%": "phi_pct",
    "Phí_Giá trị": "phi_gia_tri",
    "Thuế_%": "thue_pct",
    "Thuế_Giá trị": "thue_gia_tri",
    "Tổng thuần_Mua": "tong_thuan_mua",
    "Tổng thuần_Bán": "tong_thuan_ban",
    "Môi giới": "moi_gioi",
    "CTV": "ctv",
}

_JB_REQUIRED_INTERNAL = [
    "so_tieu_khoan", "ten_khach_hang", "ma_ck",
    "mua_gia_tri", "ban_gia_tri", "phi_gia_tri", "thue_gia_tri",
]


def parse_jb_input(file, filename: str) -> tuple[pd.DataFrame, list[str]]:
    """Doc + chuan hoa file JB input (bao cao 'Ket qua khop lenh va phi giao dich').

    Tra ve (df_normalized, warnings). Neu file thieu cot bat buoc -> ValidationError.
    """
    warnings: list[str] = []
    raw = _read_raw_table(file, filename)
    if raw.empty:
        raise ValidationError(f"File JB input '{filename}' rỗng, không có dữ liệu.")

    header_row1 = _find_row_with_marker(raw, ["STT"])
    if header_row1 is None:
        raise ValidationError(
            f"File JB input '{filename}' sai định dạng: không tìm thấy dòng tiêu đề chứa cột 'STT'. "
            "Vui lòng kiểm tra đây có đúng là báo cáo 'Kết quả khớp lệnh và phí giao dịch' của JB không."
        )
    header_row2 = header_row1 + 1
    if header_row2 >= len(raw):
        raise ValidationError(f"File JB input '{filename}' sai định dạng: thiếu dòng tiêu đề phụ.")

    combined_headers = _combine_two_row_header(raw, header_row1, header_row2)
    col_index = {}
    for idx, label in enumerate(combined_headers):
        internal = _JB_COL_MAP.get(label)
        if internal and internal not in col_index:
            col_index[internal] = idx

    missing = [k for k in _JB_REQUIRED_INTERNAL if k not in col_index]
    if missing:
        raise ValidationError(
            f"File JB input '{filename}' thiếu các cột bắt buộc: {', '.join(missing)}. "
            f"Các cột tìm thấy: {', '.join(sorted(col_index.keys())) or '(không có)'}."
        )

    company_name = _extract_company_name(raw, "JB")
    tu_ngay = _extract_pattern(raw, r"Từ ngày\s*:?\s*([0-3]?\d/[01]?\d/\d{4})")
    den_ngay = _extract_pattern(raw, r"Đến ngày\s*:?\s*([0-3]?\d/[01]?\d/\d{4})")
    single_report_date = tu_ngay if (tu_ngay and tu_ngay == den_ngay) else None

    def get(row, key):
        i = col_index.get(key)
        return row[i] if i is not None else None

    records = []
    r = header_row2 + 1
    n = len(raw)
    while r < n:
        row = raw.iloc[r]
        if row.isna().all():
            # dong trong ngan cach giua header va du lieu -> bo qua, khong phai het du lieu
            r += 1
            continue
        stt_val = get(row, "stt")
        try:
            int(str(stt_val).strip())
        except (ValueError, TypeError):
            # het du lieu (gap dong 'Tong :' / 'Tong cong')
            break

        so_tieu_khoan = _key_str(get(row, "so_tieu_khoan"))
        if not so_tieu_khoan:
            r += 1
            continue

        mua_gia_tri = _to_number(get(row, "mua_gia_tri"))
        ban_gia_tri = _to_number(get(row, "ban_gia_tri"))
        mua_kl_khop = _to_number(get(row, "mua_kl_khop"))
        ban_kl_khop = _to_number(get(row, "ban_kl_khop"))

        if mua_gia_tri > 0 or mua_kl_khop > 0:
            side = "Mua"
        elif ban_gia_tri > 0 or ban_kl_khop > 0:
            side = "Bán"
        else:
            side = None
            warnings.append(
                f"JB dòng STT {stt_val} (tiểu khoản {so_tieu_khoan}): không xác định được Mua/Bán "
                "(khối lượng khớp = 0 ở cả 2 bên) — dòng vẫn được giữ lại."
            )

        if side == "Mua":
            kl_dat, kl_khop = _to_number(get(row, "mua_kl_dat")), mua_kl_khop
            gia_khop, tong_gia = _to_number(get(row, "mua_gia_khop")), mua_gia_tri
            tong_thuan = _to_number(get(row, "tong_thuan_mua"))
        elif side == "Bán":
            kl_dat, kl_khop = _to_number(get(row, "ban_kl_dat")), ban_kl_khop
            gia_khop, tong_gia = _to_number(get(row, "ban_gia_khop")), ban_gia_tri
            tong_thuan = _to_number(get(row, "tong_thuan_ban"))
        else:
            kl_dat = kl_khop = gia_khop = tong_gia = tong_thuan = np.nan

        # Ngay giao dich: uu tien khoang ngay bao cao (neu Tu ngay == Den ngay);
        # fallback: 8 ky tu dau cua SHL thuong la YYYYMMDD.
        ngay = single_report_date
        shl = _key_str(get(row, "so_hieu_lenh")) or ""
        if not ngay and len(shl) >= 8 and shl[:8].isdigit():
            try:
                d = datetime.strptime(shl[:8], "%Y%m%d")
                ngay = d.strftime("%d/%m/%Y")
            except ValueError:
                ngay = None

        ctv_raw = _norm_text(get(row, "ctv"))
        ctv_name = ctv_raw.split("-", 1)[-1].strip() if "-" in ctv_raw else (ctv_raw or None)

        records.append({
            "source": "JB",
            "ngay": ngay,
            "so_hieu_lenh": shl or None,
            "so_tieu_khoan": so_tieu_khoan,
            "ten_khach_hang_raw": format_customer_name(_norm_text(get(row, "ten_khach_hang"))) or None,
            "ma_ck": _norm_text(get(row, "ma_ck")) or None,
            "loai_gd": side,
            "kl_dat": kl_dat,
            "kl_khop": kl_khop,
            "gia_khop": gia_khop,
            "tong_gia": tong_gia,
            "pct_phi": _to_number(get(row, "phi_pct")),
            "phi": _to_number(get(row, "phi_gia_tri")),
            "thue": _to_number(get(row, "thue_gia_tri")),
            "tong_thuan": tong_thuan,
            "ten_quan_ly_raw": ctv_name,
            "ten_cong_ty": company_name,
        })
        r += 1

    if not records:
        warnings.append(f"File JB input '{filename}' không có dòng giao dịch nào sau khi đọc.")

    return pd.DataFrame(records), warnings


# ============================================================
# 2) PBSV input
# ============================================================

_PBSV_COL_MAP = {
    "Tiểu khoản": "tieu_khoan",
    "Ngày": "ngay",
    "Mua/Bán": "mua_ban",
    "Mã CK": "ma_ck",
    "Trạng thái": "trang_thai",
    "Phí": "phi",
    "Thuế TNCN": "thue_tncn",
    "Thuế TNCN (quyền)": "thue_tncn_quyen",
    "Số hiệu lệnh": "so_hieu_lenh",
    "Thông tin giao dịch chứng khoán_KL đặt": "kl_dat",
    "Thông tin giao dịch chứng khoán_Giá đặt": "gia_dat",
    "Thông tin giao dịch chứng khoán_KL khớp": "kl_khop",
    "Thông tin giao dịch chứng khoán_Giá khớp": "gia_khop",
    "Thông tin giao dịch chứng khoán_Giá trị khớp": "gia_tri_khop",
}

_PBSV_REQUIRED_INTERNAL = [
    "tieu_khoan", "ngay", "mua_ban", "ma_ck", "trang_thai",
    "kl_dat", "kl_khop", "gia_khop", "gia_tri_khop",
]

_TRANG_THAI_HOAN_THANH = "hoàn thành"

# ------------------------------------------------------------------
# Dinh dang PBSV MOI: bao cao "Ket qua khop lenh cua khach hang".
# Khac voi dinh dang cu (2 cot, marker "Tieu khoan"):
#   - Header chi 1 dong (khong phai 2 dong ghep).
#   - Khong co cot "Trang thai" vi moi dong da la giao dich khop lenh hoan tat
#     (khong can loc "Hoan thanh").
#   - Cot "So tai khoan" tren moi dong da la khoa tieu khoan day du (vd
#     "029C000256.00"), khong can ghep voi "So tai khoan: xxx" o dau file.
#   - Xen giua cac dong du lieu la cac dong tong phu ("Tong theo ma CK: ...",
#     "Tong theo ngay:...") va cuoi file la cac dong tong ket - phai BO QUA
#     (khong duoc dung vong lap) vi du lieu con tiep tuc sau cac dong nay.
# ------------------------------------------------------------------
_PBSV_MATCH_COL_MAP = {
    "STT": "stt",
    "Ngày đặt lệnh": "ngay",
    "Mã CK": "ma_ck",
    "Loại lệnh": "mua_ban",
    "Số hiệu lệnh": "so_hieu_lenh",
    "Số tài khoản": "so_tieu_khoan",
    "SL đặt": "kl_dat",
    "SL khớp": "kl_khop",
    "Giá khớp BQ": "gia_khop",
    "Giá trị khớp": "gia_tri_khop",
    "%\nPhí": "pct_phi",
    "Giá trị phí": "phi_gia_tri",
    "Thuế TNCN": "thue_tncn",
    "Thuế quyền": "thue_quyen",
    "Phí CK": "phi_ck",
    "Thành tiền": "thanh_tien",
}

_PBSV_MATCH_REQUIRED_INTERNAL = [
    "stt", "ngay", "ma_ck", "mua_ban", "so_hieu_lenh", "so_tieu_khoan",
    "kl_dat", "kl_khop", "gia_khop", "gia_tri_khop", "phi_gia_tri", "thue_tncn",
]


def parse_pbsv_input(file, filename: str) -> tuple[pd.DataFrame, list[str]]:
    """Doc + chuan hoa file PBSV input. Tu dong nhan dien 1 trong 2 dinh dang:
      - Dinh dang cu 'Lich su dat lenh' (header 2 dong, marker 'Tieu khoan', co cot
        Trang thai -> chi giu 'Hoan thanh').
      - Dinh dang moi 'Ket qua khop lenh cua khach hang' (header 1 dong, marker
        'STT' + 'So tai khoan', khong co cot Trang thai).
    """
    warnings: list[str] = []
    raw = _read_raw_table(file, filename)
    if raw.empty:
        raise ValidationError(f"File PBSV input '{filename}' rỗng, không có dữ liệu.")

    header_row = _find_row_with_marker(raw, ["Tiểu khoản"])
    if header_row is None:
        return _parse_pbsv_match_result(raw, filename, warnings)
    sub_header_row = header_row + 1
    if sub_header_row >= len(raw):
        raise ValidationError(f"File PBSV input '{filename}' sai định dạng: thiếu dòng tiêu đề phụ.")

    combined_headers = _combine_two_row_header(raw, header_row, sub_header_row)
    col_index = {}
    for idx, label in enumerate(combined_headers):
        internal = _PBSV_COL_MAP.get(label)
        if internal and internal not in col_index:
            col_index[internal] = idx

    missing = [k for k in _PBSV_REQUIRED_INTERNAL if k not in col_index]
    if missing:
        raise ValidationError(
            f"File PBSV input '{filename}' thiếu các cột bắt buộc: {', '.join(missing)}. "
            f"Các cột tìm thấy: {', '.join(sorted(col_index.keys())) or '(không có)'}."
        )

    company_name = _extract_company_name(raw, "PBSV")
    so_tai_khoan = _extract_pattern(raw, r"Số tài khoản\s*:?\s*([0-9A-Za-z]+)")
    if not so_tai_khoan:
        warnings.append(
            f"File PBSV input '{filename}': không tìm thấy 'Số tài khoản' ở phần đầu báo cáo — "
            "'Số tiểu khoản' đầy đủ sẽ không xây dựng được, các dòng này sẽ không khớp được khách hàng."
        )

    def get(row, key):
        i = col_index.get(key)
        return row[i] if i is not None else None

    records = []
    skipped_status = 0
    r = sub_header_row + 1
    n = len(raw)
    while r < n:
        row = raw.iloc[r]
        tieu_khoan = _key_str(get(row, "tieu_khoan"))
        if not tieu_khoan:
            r += 1
            continue

        trang_thai = _norm_text(get(row, "trang_thai"))
        if trang_thai.strip().lower() != _TRANG_THAI_HOAN_THANH:
            skipped_status += 1
            r += 1
            continue

        so_tieu_khoan_full = f"{so_tai_khoan}.{tieu_khoan}" if so_tai_khoan else None

        thue = _to_number(get(row, "thue_tncn")) + _to_number(get(row, "thue_tncn_quyen"))

        records.append({
            "source": "PBSV",
            "ngay": _norm_text(get(row, "ngay")) or None,
            "so_hieu_lenh": _key_str(get(row, "so_hieu_lenh")),
            "so_tieu_khoan": so_tieu_khoan_full,
            "ten_khach_hang_raw": None,  # PBSV input khong co cot ten khach hang
            "ma_ck": _norm_text(get(row, "ma_ck")) or None,
            "loai_gd": _norm_text(get(row, "mua_ban")) or None,
            "kl_dat": _to_number(get(row, "kl_dat")),
            "kl_khop": _to_number(get(row, "kl_khop")),
            "gia_khop": _to_number(get(row, "gia_khop")),
            "tong_gia": _to_number(get(row, "gia_tri_khop")),
            "pct_phi": np.nan,   # PBSV khong co san -> tinh sau
            "phi": _to_number(get(row, "phi")),
            "thue": thue,
            "tong_thuan": np.nan,  # PBSV khong co san -> tinh sau
            "ten_quan_ly_raw": None,  # PBSV input khong co, lay tu danh sach KH
            "ten_cong_ty": company_name,
        })
        r += 1

    if skipped_status:
        warnings.append(
            f"File PBSV input '{filename}': đã loại {skipped_status} giao dịch không ở trạng thái "
            "'Hoàn thành' (theo quy tắc xử lý)."
        )
    if not records:
        warnings.append(
            f"File PBSV input '{filename}' không có giao dịch nào ở trạng thái 'Hoàn thành'."
        )

    return pd.DataFrame(records), warnings


def _parse_pbsv_match_result(raw: pd.DataFrame, filename: str, warnings: list[str]) -> tuple[pd.DataFrame, list[str]]:
    """Doc dinh dang PBSV MOI: bao cao 'Ket qua khop lenh cua khach hang' (xem ghi chu o
    _PBSV_MATCH_COL_MAP). Header chi 1 dong; khong loc Trang thai vi moi dong da la giao
    dich khop lenh hoan tat; cac dong "Tong theo ma CK" / "Tong theo ngay" / tong ket cuoi
    file duoc bo qua bang continue (khong break) vi du lieu con tiep tuc sau do.
    """
    header_row = _find_row_with_marker(raw, ["STT"])
    header_texts = {_norm_text(v) for v in raw.iloc[header_row]} if header_row is not None else set()
    if header_row is None or "Số tài khoản" not in header_texts:
        raise ValidationError(
            f"File PBSV input '{filename}' sai định dạng: không tìm thấy dòng tiêu đề chứa cột 'Tiểu khoản' "
            "(định dạng 'Lịch sử đặt lệnh') hoặc 2 cột 'STT' + 'Số tài khoản' (định dạng 'Kết quả khớp lệnh "
            "của khách hàng'). Vui lòng kiểm tra lại file báo cáo PBSV."
        )

    header = raw.iloc[header_row].tolist()
    col_index = {}
    for idx, cell in enumerate(header):
        internal = _PBSV_MATCH_COL_MAP.get(_norm_text(cell))
        if internal and internal not in col_index:
            col_index[internal] = idx

    missing = [k for k in _PBSV_MATCH_REQUIRED_INTERNAL if k not in col_index]
    if missing:
        raise ValidationError(
            f"File PBSV input '{filename}' (định dạng 'Kết quả khớp lệnh') thiếu các cột bắt buộc: "
            f"{', '.join(missing)}. Các cột tìm thấy: {', '.join(sorted(col_index.keys())) or '(không có)'}."
        )

    company_name = _extract_company_name(raw, "PBSV")

    def get(row, key):
        i = col_index.get(key)
        return row[i] if i is not None else None

    records = []
    r = header_row + 1
    n = len(raw)
    while r < n:
        row = raw.iloc[r]
        stt_val = get(row, "stt")
        try:
            int(str(stt_val).strip())
        except (ValueError, TypeError):
            # dong "Tong theo ma CK: ..." / "Tong theo ngay:..." / dong tong ket cuoi file
            # -> bo qua, KHONG dung vong lap vi du lieu giao dich con tiep tuc sau do.
            r += 1
            continue

        so_tieu_khoan = _key_str(get(row, "so_tieu_khoan"))
        if not so_tieu_khoan:
            r += 1
            continue

        phi = _to_number(get(row, "phi_gia_tri")) + _to_number(get(row, "phi_ck"))
        thue = _to_number(get(row, "thue_tncn")) + _to_number(get(row, "thue_quyen"))

        records.append({
            "source": "PBSV",
            "ngay": _format_date_cell(get(row, "ngay")),
            "so_hieu_lenh": _key_str(get(row, "so_hieu_lenh")),
            "so_tieu_khoan": so_tieu_khoan,
            "ten_khach_hang_raw": None,  # dinh dang nay khong co cot ten khach hang
            "ma_ck": _norm_text(get(row, "ma_ck")) or None,
            "loai_gd": _norm_text(get(row, "mua_ban")) or None,
            "kl_dat": _to_number(get(row, "kl_dat")),
            "kl_khop": _to_number(get(row, "kl_khop")),
            "gia_khop": _to_number(get(row, "gia_khop")),
            "tong_gia": _to_number(get(row, "gia_tri_khop")),
            "pct_phi": _to_number(get(row, "pct_phi")),
            "phi": phi,
            "thue": thue,
            "tong_thuan": _to_number(get(row, "thanh_tien")),
            "ten_quan_ly_raw": None,  # PBSV luon lay ten Moi gioi tu danh sach khach hang
            "ten_cong_ty": company_name,
        })
        r += 1

    if not records:
        warnings.append(
            f"File PBSV input '{filename}' (định dạng 'Kết quả khớp lệnh') không có dòng giao dịch nào sau khi đọc."
        )

    return pd.DataFrame(records), warnings


# ============================================================
# 3) Danh sach khach hang (2 sheet: JB + PBSV)
# ============================================================

_MASTER_JB_REQUIRED = ["TÀI KHOẢN LƯU KÝ", "SỐ TIỂU KHOẢN", "TÊN KHÁCH HÀNG", "TÊN MÔI GIỚI"]
_MASTER_PB_REQUIRED = ["TÀI KHOẢN LƯU KÝ", "SỐ TIỂU KHOẢN", "TÊN KHÁCH HÀNG", "TÊN MÔI GIỚI"]


def parse_customer_master(file, filename: str) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Doc file danh sach khach hang JB + PB (2 sheet). Tra ve (master_jb, master_pb, warnings),
    da index theo 'SỐ TIỂU KHOẢN' da chuan hoa (chuoi, khong index() de con giu cot).
    """
    warnings: list[str] = []
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if ext not in ("xlsx", "xlsm", "xls"):
        raise ValidationError(
            f"File danh sách khách hàng '{filename}' phải ở định dạng Excel (.xlsx) có 2 sheet (JB và PB)."
        )
    try:
        engine = "xlrd" if ext == "xls" else "openpyxl"
        xls = pd.ExcelFile(file, engine=engine)
    except Exception as e:  # noqa: BLE001
        raise ValidationError(f"Không đọc được file danh sách khách hàng '{filename}': {e}") from e

    if len(xls.sheet_names) < 2:
        raise ValidationError(
            f"File danh sách khách hàng '{filename}' cần có 2 sheet (JB và PB), "
            f"nhưng chỉ tìm thấy {len(xls.sheet_names)} sheet: {xls.sheet_names}."
        )

    sheet_jb = next((s for s in xls.sheet_names if "jb" in s.lower()), None)
    sheet_pb = next(
        (s for s in xls.sheet_names if s != sheet_jb and ("pb" in s.lower() or "pbsv" in s.lower())),
        None,
    )
    if sheet_jb is None or sheet_pb is None:
        raise ValidationError(
            f"Không xác định được sheet nào là JB / PB trong file '{filename}' "
            f"(tên sheet hiện có: {xls.sheet_names}). Vui lòng đặt tên sheet có chứa 'JB' và 'PB'."
        )

    def _load_sheet(sheet_name, required_cols, label):
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            raise ValidationError(
                f"Sheet '{sheet_name}' ({label}) trong file '{filename}' thiếu cột bắt buộc: "
                f"{', '.join(missing)}. Các cột tìm thấy: {', '.join(df.columns)}."
            )
        df = df.copy()
        if "TÊN KHÁCH HÀNG" in df.columns:
            df["TÊN KHÁCH HÀNG"] = df["TÊN KHÁCH HÀNG"].apply(
                lambda v: format_customer_name(str(v)) if pd.notna(v) and str(v).strip() else v
            )
        df["_key"] = df["SỐ TIỂU KHOẢN"].apply(_key_str)
        df = df[df["_key"].notna()]
        dup = df["_key"].duplicated()
        if dup.any():
            warnings.append(
                f"Sheet '{sheet_name}' ({label}): có {int(dup.sum())} 'SỐ TIỂU KHOẢN' bị trùng lặp — "
                "chỉ dòng đầu tiên được dùng để đối chiếu."
            )
            df = df[~dup]
        return df.set_index("_key", drop=False)

    master_jb = _load_sheet(sheet_jb, _MASTER_JB_REQUIRED, "JB")
    master_pb = _load_sheet(sheet_pb, _MASTER_PB_REQUIRED, "PB")
    return master_jb, master_pb, warnings


# ============================================================
# 4) Join + chuan hoa ve schema output
# ============================================================

def _derive_custody_account(so_tieu_khoan: str, source: str) -> Optional[str]:
    """Suy ra 'TÀI KHOẢN LƯU KÝ' tu 'SỐ TIỂU KHOẢN' theo quy tac rieng cho tung broker -
    dung lam FALLBACK khi khong tim thay khop truc tiep theo SỐ TIỂU KHOẢN trong Master.
    Tra ve None neu khong the suy ra mot cach chac chan (khong doan bua de tranh gan nham
    khach hang/CTV cho giao dich).

        - PBSV: bo phan duoi sau dau '.' cuoi cung (vd "123456.00" -> "123456"). Khong co
          dau '.' -> giu nguyen ca chuoi.
        - JB: bo 2 ky tu cuoi cua so tieu khoan (ma hau to nhu "MG"/"NM"/"M1"), sau do neu
          4 ky tu dau cua phan con lai la "0001" (ma noi bo) -> thay bang "050C" (ma tai
          khoan luu ky chuan, da doi chieu khop 100% voi Master JB hien co). Prefix khac
          "0001" -> chua co quy tac xac nhan -> tra ve None (khong doan).
        - Nguon khac (ngoai JB/PBSV, hien khong co trong he thong) -> copy nguyen so tieu
          khoan lam tai khoan luu ky tam thoi.
    """
    if not so_tieu_khoan:
        return None
    s = so_tieu_khoan.strip()
    if not s:
        return None

    if source == "PBSV":
        return s.rsplit(".", 1)[0].strip() if "." in s else s

    if source == "JB":
        if len(s) <= 2:
            return None
        core = s[:-2]
        if core[:4] == "0001":
            return "050C" + core[4:]
        return None

    return s


def _build_custody_index(master: pd.DataFrame) -> dict:
    """Xay dung index phu theo 'TÀI KHOẢN LƯU KÝ', dung de tra cuu fallback khi khop qua
    _derive_custody_account(). Nhieu dong (nhieu tieu khoan) co the cung 1 tai khoan luu
    ky - cac dong nay theo nguyen tac phai cung 1 khach hang/CTV nen chi can lay dong dau
    tien gap duoc."""
    idx: dict = {}
    if master.empty or "TÀI KHOẢN LƯU KÝ" not in master.columns:
        return idx
    for _, row in master.iterrows():
        k = _key_str(row["TÀI KHOẢN LƯU KÝ"])
        if k and k not in idx:
            idx[k] = row
    return idx


def _join_source(df_norm: pd.DataFrame, master: pd.DataFrame, source: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Doi chieu du lieu giao dich (JB hoac PBSV) voi danh sach khach hang tuong ung. Neu
    khong khop truc tiep theo 'SỐ TIỂU KHOẢN', thu suy ra 'TÀI KHOẢN LƯU KÝ' tu so tieu
    khoan (_derive_custody_account) roi khop lai theo tai khoan luu ky - fallback danh cho
    cac dong ma Master khong liet ke dung tieu khoan giao dich nhung co liet ke tai khoan
    luu ky qua 1 tieu khoan khac cua cung khach hang.
    Tra ve (output_rows_df, mismatch_rows_df, match_stats).
    """
    match_stats = {"truc_tiep": 0, "suy_luan_tai_khoan_luu_ky": 0, "khong_khop": 0, "xoa_thieu_ten": 0}
    if df_norm.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS), pd.DataFrame(columns=MISMATCH_COLUMNS), match_stats

    custody_index = _build_custody_index(master)

    out_rows = []
    mismatch_rows = []

    for _, row in df_norm.iterrows():
        key = row["so_tieu_khoan"]
        m = master.loc[key] if (key is not None and key in master.index) else None
        matched = m is not None
        derived_custody = None

        if matched:
            match_stats["truc_tiep"] += 1
        elif key:
            derived_custody = _derive_custody_account(key, source)
            if derived_custody:
                m2 = custody_index.get(derived_custody)
                if m2 is not None:
                    m = m2
                    matched = True
                    match_stats["suy_luan_tai_khoan_luu_ky"] += 1

        if not matched:
            match_stats["khong_khop"] += 1

        raw_cust = row.get("ten_khach_hang_raw")
        raw_cust_fmt = format_customer_name(raw_cust) if raw_cust else ""

        if not matched:
            if not key:
                ly_do = "Thiếu số tiểu khoản để đối chiếu (kiểm tra 'Số tài khoản' ở đầu báo cáo PBSV)"
            elif derived_custody:
                ly_do = (
                    f"Không tìm thấy trong danh sách khách hàng (đã thử suy ra tài khoản lưu ký "
                    f"'{derived_custody}' từ số tiểu khoản nhưng cũng không khớp)"
                )
            else:
                ly_do = "Không tìm thấy trong danh sách khách hàng (không suy ra được tài khoản lưu ký từ số tiểu khoản)"
            mismatch_rows.append({
                "Nguồn": source,
                "Số tiểu khoản tra cứu": key if key else "(rỗng)",
                "Tên khách hàng (nếu có)": raw_cust_fmt,
                "Lý do": ly_do,
            })

        tai_khoan_luu_ky = m["TÀI KHOẢN LƯU KÝ"] if matched else ""
        master_cust = m["TÊN KHÁCH HÀNG"] if matched else ""
        master_cust_fmt = format_customer_name(master_cust) if (matched and pd.notna(master_cust)) else ""
        ten_khach_hang = raw_cust_fmt or master_cust_fmt

        if source == "JB":
            # Theo yeu cau: JB dung ten CTV lam "Nguoi quan ly". Uu tien gia tri co san
            # tren chinh dong giao dich, fallback ve danh sach khach hang.
            raw_manager = row.get("ten_quan_ly_raw")
            master_manager = m.get("TÊN CTV") if matched and "TÊN CTV" in master.columns else None
            master_manager = _norm_text(master_manager) or None
            if raw_manager:
                ten_quan_ly = raw_manager
            elif master_manager:
                ten_quan_ly = master_manager
            elif matched:
                ten_quan_ly = ""  # khach hang co that nhung khong co CTV duoc gan - khong phai loi khop
            else:
                ten_quan_ly = ""
        else:
            # PBSV: dung ten Moi gioi, luon lay tu danh sach khach hang (input khong co san).
            ten_quan_ly = _norm_text(m.get("TÊN MÔI GIỚI")) if matched else ""

        # "Nhóm khách hàng" / "Phân loại khách hàng": luon lay tu danh sach khach hang
        # (Master). O trong/None/NaN/chi co khoang trang, hoac khong khop duoc khach hang
        # -> dien mac dinh DEFAULT_UNKNOWN thay vi de trong o.
        nhom_kh = _norm_text(m.get("NHÓM KHÁCH HÀNG")) if (matched and "NHÓM KHÁCH HÀNG" in master.columns) else ""
        phan_loai_kh = (
            _norm_text(m.get("PHÂN LOẠI KHÁCH HÀNG"))
            if (matched and "PHÂN LOẠI KHÁCH HÀNG" in master.columns)
            else ""
        )
        nhom_kh = nhom_kh or DEFAULT_UNKNOWN
        phan_loai_kh = phan_loai_kh or DEFAULT_UNKNOWN

        # Theo yeu cau nghiep vu: khong de trong o "Tên khách hàng" trong bang ket qua
        # chinh - neu khong tra cuu duoc ten khach hang (du da khop tieu khoan hay chua)
        # thi XOA CA DONG khoi bang chinh, chi giu vet lai trong mismatch_df de nguoi
        # dung kiem chung. Rieng "Tên người quản lý" (Moi gioi/CTV) thieu thi VAN GIU
        # dong, chi de trong o do (khach hang co that nhung khong co Moi gioi/CTV duoc
        # gan khong phai loi can xoa dong).
        if not ten_khach_hang:
            match_stats["xoa_thieu_ten"] += 1
            if matched:
                mismatch_rows.append({
                    "Nguồn": source,
                    "Số tiểu khoản tra cứu": key if key else "(rỗng)",
                    "Tên khách hàng (nếu có)": raw_cust_fmt,
                    "Lý do": "Đã khớp tiểu khoản nhưng thiếu Tên khách hàng trong danh sách khách hàng — đã xóa dòng khỏi bảng kết quả",
                })
            continue

        out_rows.append({
            "Ngày": row.get("ngay"),
            "Số hiệu lệnh": row.get("so_hieu_lenh"),
            "Tài khoản lưu ký ": tai_khoan_luu_ky,
            "Số tiểu khoản": key,
            "Tên khách hàng": ten_khach_hang,
            "Mã Chứng khoán": row.get("ma_ck"),
            "Giao dịch (Mua/Bán)": row.get("loai_gd"),
            "Khối lượng đặt": row.get("kl_dat"),
            "Khối lượng khớp": row.get("kl_khop"),
            "Giá khớp": row.get("gia_khop"),
            COL_TRUOC: row.get("tong_gia"),
            "% Phí": row.get("pct_phi"),
            "Phí": row.get("phi"),
            "Thuế": row.get("thue"),
            COL_SAU: row.get("tong_thuan"),
            "Tên người quản lý": ten_quan_ly,
            "Tên Công ty": row.get("ten_cong_ty"),
            "Nhóm khách hàng": nhom_kh,
            "Phân loại khách hàng": phan_loai_kh,
        })

    out_df = pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS)
    mismatch_df = pd.DataFrame(mismatch_rows, columns=MISMATCH_COLUMNS)
    return out_df, mismatch_df, match_stats


# Sai so cho phep khi doi chieu COL_SAU voi cong thuc chuan (VND) - chi de bo qua
# sai so lam tron, KHONG de che giau chenh lech that su (vd 2x Phi+Thue do quy uoc
# netting nguoc chieu cua du lieu nguon).
_VALIDATION_TOLERANCE = 1.0

# So dong loi toi da liet ke chi tiet trong 1 canh bao, tranh spam khi loi hang loat.
_VALIDATION_MAX_LOGGED = 20


def _fill_derived_fields(df: pd.DataFrame, warnings: list[str]) -> pd.DataFrame:
    """Tinh COL_SAU va '% Phi' cho TOAN BO dong, theo dung nghiep vu rieng cho tung chieu
    Mua/Ban (nguoi mua tra them phi; nguoi ban bi tru ca phi lan thue TNCN truoc khi
    nhan tien):

        Mua: COL_SAU ("Tổng thuần") = COL_TRUOC ("Tổng giá ") + Phí
        Bán: COL_SAU ("Tổng thuần") = COL_TRUOC ("Tổng giá ") − Phí − Thuế

    Ap dung THONG NHAT cho MOI nguon (JB/PBSV). Dong hiem gap khong xac dinh duoc
    Mua/Ban (xem parse_jb_input) dung cong thuc du phong: + Phi + Thue.

    Gia tri co san tu nguon (JB: "Tổng thuần_Mua/Bán"; PBSV dinh dang moi: "Thành tiền")
    duoc dung de DOI CHIEU/CANH BAO neu lech so voi cong thuc nghiep vu tren, khong dung
    truc tiep lam ket qua cuoi cung.

    Chuan hoa tat ca cot so thanh kieu du lieu so (float64) thay vi chuoi.
    """
    df = df.copy()

    # Ep tat ca cac cot so thanh kieu so (numeric)
    num_cols = ["Khối lượng đặt", "Khối lượng khớp", "Giá khớp", COL_TRUOC, "% Phí", "Phí", "Thuế", COL_SAU]
    for col in num_cols:
        if col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].apply(_to_number)
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    tong_truoc = df[COL_TRUOC]
    phi = df["Phí"]
    thue = df["Thuế"]
    pct_phi = df["% Phí"]
    sau_nguon = df[COL_SAU]  # gia tri COL_SAU co san tu nguon (chi JB co, PBSV luon = 0.0)
    loai_gd = df["Giao dịch (Mua/Bán)"]

    is_mua = loai_gd == "Mua"
    is_ban = loai_gd == "Bán"
    sau_chuan = np.select(
        [is_mua, is_ban],
        [tong_truoc + phi, tong_truoc - phi - thue],
        default=(tong_truoc + phi + thue),  # du phong cho dong khong xac dinh Mua/Ban
    )
    sau_chuan = pd.Series(sau_chuan, index=df.index)

    # ── Validate: doi chieu gia tri co san tu nguon voi cong thuc nghiep vu tren.
    # Neu lech vuot sai so lam tron -> log lai tung dong de de truy vet, KHONG am
    # tham bo qua (van uu tien cong thuc nghiep vu cho ket qua cuoi cung ben duoi).
    co_gia_tri_nguon = sau_nguon != 0.0
    lech = (sau_nguon - sau_chuan).abs()
    mismatch_mask = co_gia_tri_nguon & (lech > _VALIDATION_TOLERANCE)

    n_mismatch = int(mismatch_mask.sum())
    if n_mismatch:
        msg_header = (
            f"⚠️ Validate '{COL_SAU}': phát hiện {n_mismatch} dòng có giá trị lấy từ file nguồn "
            f"KHÁC với công thức nghiệp vụ (Mua: '{COL_TRUOC}' + Phí; Bán: '{COL_TRUOC}' − Phí − Thuế) "
            "— đã dùng công thức nghiệp vụ cho kết quả cuối cùng, chi tiết các dòng lệch:"
        )
        warnings.append(msg_header)
        mismatch_idx = df.index[mismatch_mask]
        for i, idx in enumerate(mismatch_idx):
            if i >= _VALIDATION_MAX_LOGGED:
                warnings.append(f"  ... và {n_mismatch - _VALIDATION_MAX_LOGGED} dòng khác.")
                break
            row = df.loc[idx]
            warnings.append(
                f"  - SHL {row.get('Số hiệu lệnh') or '(trống)'} | "
                f"Tiểu khoản {row.get('Số tiểu khoản') or '(trống)'} | "
                f"KH {row.get('Tên khách hàng') or '(trống)'} | "
                f"Nguồn {row.get('Tên Công ty') or '(trống)'}: "
                f"giá trị nguồn = {sau_nguon.loc[idx]:,.2f}, "
                f"công thức nghiệp vụ = {sau_chuan.loc[idx]:,.2f}, "
                f"chênh lệch = {(sau_nguon.loc[idx] - sau_chuan.loc[idx]):,.2f}."
            )

    df[COL_SAU] = sau_chuan

    need_pf = (pct_phi == 0.0) | (pct_phi.isna())
    safe_div = tong_truoc.replace(0, np.nan)
    pct_phi = pct_phi.where(~need_pf, (phi / safe_div * 100))
    df["% Phí"] = pct_phi.round(4)
    return df

def build_output(
    df_jb_norm: pd.DataFrame,
    df_pb_norm: pd.DataFrame,
    master_jb: pd.DataFrame,
    master_pb: pd.DataFrame,
    warnings: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Gop du lieu JB + PBSV da chuan hoa, join voi danh sach khach hang, tra ve
    (output_df, mismatch_df, map_stats) theo dung cau truc chuan.

    `warnings` duoc bo sung them cac canh bao validate COL_SAU (xem _fill_derived_fields)
    va 1 dong tom tat ket qua doi chieu tai khoan luu ky (truc tiep / suy luan / khong
    khop) theo tung nguon, de nguoi dung kiem chung (xem _join_source).
    """
    out_jb, mis_jb, stats_jb = _join_source(df_jb_norm, master_jb, "JB")
    out_pb, mis_pb, stats_pb = _join_source(df_pb_norm, master_pb, "PBSV")

    output_df = pd.concat([out_jb, out_pb], ignore_index=True)
    output_df = _fill_derived_fields(output_df, warnings)

    mismatch_df = pd.concat([mis_jb, mis_pb], ignore_index=True)

    map_stats = {"JB": stats_jb, "PBSV": stats_pb}
    for src, st in map_stats.items():
        if st["truc_tiep"] or st["suy_luan_tai_khoan_luu_ky"] or st["khong_khop"]:
            warnings.append(
                f"ℹ️ Đối chiếu tài khoản lưu ký nguồn {src}: {st['truc_tiep']} khớp trực tiếp theo "
                f"số tiểu khoản, {st['suy_luan_tai_khoan_luu_ky']} khớp qua suy luận tài khoản lưu ký "
                f"từ số tiểu khoản, {st['khong_khop']} không khớp được."
            )
        if st["xoa_thieu_ten"]:
            warnings.append(
                f"🗑️ Nguồn {src}: đã xóa {st['xoa_thieu_ten']} dòng khỏi bảng kết quả do không tra "
                "cứu được Tên khách hàng — xem chi tiết ở danh sách dòng không khớp."
            )

    return output_df, mismatch_df, map_stats

# ============================================================
# 5) Ham dieu phoi tong (goi tu Streamlit app) - ho tro nhieu file/o input
# ============================================================

def _concat_master(frames: list[pd.DataFrame], label: str, warnings: list[str]) -> pd.DataFrame:
    """Gop danh sach khach hang tu nhieu file 'DS khach hang' lai lam mot.
    Neu 1 'SỐ TIỂU KHOẢN' xuat hien o nhieu file, giu ban ghi dau tien va canh bao."""
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=False)
    dup = combined.index.duplicated()
    if dup.any():
        warnings.append(
            f"Danh sách khách hàng {label}: có {int(dup.sum())} 'SỐ TIỂU KHOẢN' trùng lặp giữa "
            "các file đã tải lên — chỉ bản ghi xuất hiện đầu tiên được dùng để đối chiếu."
        )
        combined = combined[~dup]
    return combined

def process_files(jb_items, pbsv_items, master_items):
    """Ham tong hop: doc nhieu file cho moi loai dau vao (JB / PBSV / DS khach hang),
    xu ly va tra ve dict ket qua.

    Tham so:
        jb_items, pbsv_items, master_items: list cac tuple (file, filename).
        Moi o input tren giao dien co the nhan 1 hoac nhieu file cung luc.

    Tra ve dict:
        {
            'output_df': DataFrame | None,
            'mismatch_df': DataFrame,
            'errors': list[str]        # loi nghiem trong (fatal) - khong co output
            'file_errors': list[str]   # loi rieng cua tung file JB/PBSV - cac file con lai van duoc xu ly
            'warnings': list[str]      # canh bao (khong chan xu ly)
            'stats': dict
        }
    Ham nay khong bao gio raise - moi loi duoc bat va tra ve trong 'errors'/'file_errors'.
    """
    errors: list[str] = []
    file_errors: list[str] = []
    warnings: list[str] = []
    stats = {}

    # ── Kiểm tra yêu cầu file đầu vào tối thiểu ──
    if not master_items:
        errors.append("Vui lòng tải lên ít nhất 1 file Danh sách khách hàng (Master).")
    if not jb_items and not pbsv_items:
        errors.append("Vui lòng tải lên ít nhất 1 file giao dịch đầu vào (JB hoặc PBSV).")

    if errors:
        return {
            "output_df": None,
            "mismatch_df": pd.DataFrame(columns=MISMATCH_COLUMNS),
            "errors": errors,
            "file_errors": file_errors,
            "warnings": warnings,
            "stats": stats,
        }

    # ── Danh sach khach hang: loi o day la fatal vi anh huong toan bo doi chieu ──
    master_jb_frames: list[pd.DataFrame] = []
    master_pb_frames: list[pd.DataFrame] = []
    for f, fname in master_items:
        try:
            m_jb, m_pb, w = parse_customer_master(f, fname)
            master_jb_frames.append(m_jb)
            master_pb_frames.append(m_pb)
            warnings.extend(w)
        except ValidationError as e:
            errors.append(str(e))
        except Exception as e:  # noqa: BLE001
            errors.append(f"Lỗi không xác định khi đọc file danh sách khách hàng '{fname}': {e}")

    if errors:
        return {
            "output_df": None,
            "mismatch_df": pd.DataFrame(columns=MISMATCH_COLUMNS),
            "errors": errors,
            "file_errors": file_errors,
            "warnings": warnings,
            "stats": stats,
        }

    master_jb = _concat_master(master_jb_frames, "JB", warnings)
    master_pb = _concat_master(master_pb_frames, "PBSV", warnings)

    # ── JB / PBSV input: loi tren 1 file khong chan cac file con lai ──
    df_jb_frames: list[pd.DataFrame] = []
    for f, fname in jb_items:
        try:
            df, w = parse_jb_input(f, fname)
            df_jb_frames.append(df)
            warnings.extend(w)
        except ValidationError as e:
            file_errors.append(str(e))
        except Exception as e:  # noqa: BLE001
            file_errors.append(f"Lỗi không xác định khi đọc file JB input '{fname}': {e}")
    df_jb = pd.concat(df_jb_frames, ignore_index=True) if df_jb_frames else pd.DataFrame()

    df_pb_frames: list[pd.DataFrame] = []
    for f, fname in pbsv_items:
        try:
            df, w = parse_pbsv_input(f, fname)
            df_pb_frames.append(df)
            warnings.extend(w)
        except ValidationError as e:
            file_errors.append(str(e))
        except Exception as e:  # noqa: BLE001
            file_errors.append(f"Lỗi không xác định khi đọc file PBSV input '{fname}': {e}")
    df_pb = pd.concat(df_pb_frames, ignore_index=True) if df_pb_frames else pd.DataFrame()

    try:
        output_df, mismatch_df, map_stats = build_output(df_jb, df_pb, master_jb, master_pb, warnings)
    except Exception as e:  # noqa: BLE001
        errors.append(f"Lỗi không xác định khi gộp/đối chiếu dữ liệu: {e}")
        return {
            "output_df": None,
            "mismatch_df": pd.DataFrame(columns=MISMATCH_COLUMNS),
            "errors": errors,
            "file_errors": file_errors,
            "warnings": warnings,
            "stats": stats,
        }

    stats = {
        "so_file_jb": len(jb_items),
        "so_file_pbsv": len(pbsv_items),
        "so_gd_jb": len(df_jb),
        "so_gd_pbsv": len(df_pb),
        "tong_so_gd": len(output_df),
        "so_dong_khong_khop": len(mismatch_df),
        "map_stats": map_stats,
    }

    return {
        "output_df": output_df,
        "mismatch_df": mismatch_df,
        "errors": errors,
        "file_errors": file_errors,
        "warnings": warnings,
        "stats": stats,
    }
