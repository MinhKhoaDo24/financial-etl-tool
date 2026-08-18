# -*- coding: utf-8 -*-
"""
Test ngan gon cho merge_processor.py:
  - parse_pbsv_input doc dung ca 2 dinh dang PBSV (MOI 'Ket qua khop lenh cua khach hang'
    - header 1 dong, khong co cot Trang thai, xen ke cac dong "Tong theo..." can bo qua;
    va CU 'Lich su dat lenh' - regression, dam bao khong bi vo khi them nhan dien dinh
    dang moi).
  - _join_source suy ra Tai khoan luu ky tu So tieu khoan (_derive_custody_account) lam
    fallback khi khong khop truc tiep, roi khop lai qua Tai khoan luu ky.

Chay: python test_merge_processor.py
"""
from __future__ import annotations

import datetime
import io
import unittest

import openpyxl
import pandas as pd

import merge_processor as mp


def _build_xlsx(rows: list[list]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class TestParsePbsvNewFormat(unittest.TestCase):
    def setUp(self):
        self.rows = [
            ["BÁO CÁO KẾT QUẢ KHỚP LỆNH CỦA KHÁCH HÀNG\nTừ ngày 01/08/2026 đến ngày 17/08/2026"],
            ["Khách hàng: Tất cả\nSố tài khoản: Tất cả"],
            [],
            ["STT", "Ngày đặt lệnh", "Mã CK", "Loại lệnh", "Số hiệu lệnh", "Số tài khoản",
             "SL đặt", "Giá đặt", "SL khớp", "Giá khớp BQ", "Giá trị khớp", "%\nPhí",
             "Giá trị phí", "Thuế TNCN", "Thuế quyền", "Phí CK", "Thành tiền"],
            [1, datetime.datetime(2026, 8, 3), "VIX", "Bán", "1030826000070", "029C000256.00",
             1400, 13200, 1400, 13200, 18480000, 0.15, 27720, 18480, 0, 0, 18433800],
            # dong tong phu xen giua - phai bi bo qua (khong lam dung vong lap doc)
            ["Tổng theo mã CK: VIX", None, None, None, None, None, None, None, 1400, None,
             18480000, None, 27720, 18480, 0, 0, 18433800],
            [2, datetime.datetime(2026, 8, 3), "DPM", "Mua", "8000030826000082", "029C001879.01",
             100, 21950, 100, 21950, 2195000, 0.15, 3293, 0, 0, 0, 2198293],
            # dong tong ket cuoi file - cung phai bi bo qua
            ["Tổng theo ngày:03/08/2026", None, None, None, None, None, None, None, None, None,
             20675000, None, 31013, 18480, 0, 0, 20632093],
            [None, None, None, None, None, None, None, None, None, "Tổng Mua + phí", None, None,
             2198293, None, None, None, None],
        ]

    def test_skips_subtotal_rows_and_keeps_only_real_data(self):
        df, _ = mp.parse_pbsv_input(_build_xlsx(self.rows), "PB input.xlsx")
        self.assertEqual(len(df), 2)

    def test_field_mapping_uses_full_account_key_and_side_convention(self):
        df, _ = mp.parse_pbsv_input(_build_xlsx(self.rows), "PB input.xlsx")
        ban = df.iloc[0]
        self.assertEqual(ban["so_tieu_khoan"], "029C000256.00")
        self.assertEqual(ban["ngay"], "03/08/2026")
        self.assertEqual(ban["loai_gd"], "Bán")
        self.assertIsNone(ban["ten_khach_hang_raw"])
        self.assertEqual(ban["ten_cong_ty"], "PBSV")
        self.assertAlmostEqual(ban["tong_gia"], 18480000)
        self.assertAlmostEqual(ban["phi"], 27720)
        self.assertAlmostEqual(ban["thue"], 18480)


class TestParsePbsvOldFormatRegression(unittest.TestCase):
    """Dinh dang cu 'Lich su dat lenh' phai tiep tuc hoat dong dung sau khi them nhan dien
    dinh dang moi vao parse_pbsv_input."""

    def test_old_format_still_parses_and_filters_non_hoan_thanh(self):
        rows = [
            [None, "LỊCH SỬ ĐẶT LỆNH"],
            [None, "Từ ngày: 01/07/2026"],
            [None, "Đến ngày: 05/08/2026"],
            [None, "Số tài khoản: 029C002214"],
            [None, "Tiểu khoản", "Ngày", "Mua/Bán", "Mã CK", "Loại giá", None,
             "Thông tin giao dịch chứng khoán", None, None, None, None, None, None,
             "Trạng thái", "Phí", "Thuế TNCN", "Thuế TNCN (quyền)"],
            [None, None, None, None, None, None, None, "KL đặt", "Giá đặt", "KL khớp",
             "Giá khớp", "Giá trị khớp"],
            [None, "00", "30/07/2026", "Bán", "KHG", "LO", None, 9500.0, 4740.0, 9500.0,
             4740.0, 45030000.0, None, None, "Hoàn thành", 45030.0, 45030.0, 0.0],
            [None, "00", "30/07/2026", "Bán", "KHG", "LO", None, 100.0, 4740.0, 0.0, 0.0,
             0.0, None, None, "Đã sửa", 0.0, None, 0.0],
        ]
        df, _ = mp.parse_pbsv_input(_build_xlsx(rows), "old.xlsx")
        self.assertEqual(len(df), 1)  # dong "Đã sửa" phai bi loc bo
        self.assertEqual(df.iloc[0]["so_tieu_khoan"], "029C002214.00")


def _build_master_df(records: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(records)
    df["_key"] = df["SỐ TIỂU KHOẢN"].apply(mp._key_str)
    return df.set_index("_key", drop=False)


def _tx_row(**overrides) -> pd.DataFrame:
    base = {
        "ten_khach_hang_raw": None, "ten_quan_ly_raw": None, "ngay": "10/08/2026",
        "so_hieu_lenh": "T1", "ma_ck": "ABC", "loai_gd": "Mua", "kl_dat": 100.0,
        "kl_khop": 100.0, "gia_khop": 10000.0, "tong_gia": 1000000.0, "pct_phi": 0.15,
        "phi": 1500.0, "thue": 0.0, "tong_thuan": 1001500.0,
    }
    base.update(overrides)
    return pd.DataFrame([base])


class TestCustodyAccountFallback(unittest.TestCase):
    """Khi khong khop truc tiep theo So tieu khoan, _join_source phai thu suy ra Tai
    khoan luu ky (_derive_custody_account) va khop lai qua do truoc khi bao that bai."""

    def test_derive_rules(self):
        self.assertEqual(mp._derive_custody_account("123456.00", "PBSV"), "123456")
        self.assertEqual(mp._derive_custody_account("029C123456", "PBSV"), "029C123456")
        self.assertEqual(mp._derive_custody_account("0001986886MG", "JB"), "050C986886")
        self.assertIsNone(mp._derive_custody_account("0099123456XX", "JB"))  # prefix la

    def test_pbsv_fallback_success_via_sibling_subaccount(self):
        master_pb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "029C030987", "SỐ TIỂU KHOẢN": "029C030987.00",
             "TÊN KHÁCH HÀNG": "Bùi Thị Phương Ngân", "TÊN MÔI GIỚI": "Cao Ngọc Toàn"},
        ])
        df_norm = _tx_row(source="PBSV", so_tieu_khoan="029C030987.05", ten_cong_ty="PBSV")
        out_df, mismatch_df, stats = mp._join_source(df_norm, master_pb, "PBSV")

        self.assertEqual(
            stats,
            {"truc_tiep": 0, "suy_luan_tai_khoan_luu_ky": 1, "khong_khop": 0, "xoa_thieu_ten": 0},
        )
        self.assertEqual(out_df.iloc[0]["Tài khoản lưu ký "], "029C030987")
        self.assertEqual(out_df.iloc[0]["Tên khách hàng"], "Bùi Thị Phương Ngân")
        self.assertEqual(out_df.iloc[0]["Tên người quản lý"], "Cao Ngọc Toàn")
        self.assertEqual(len(mismatch_df), 0)

    def test_jb_fallback_success_via_prefix_rule(self):
        master_jb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "050C986886", "SỐ TIỂU KHOẢN": "0001986886MG",
             "TÊN KHÁCH HÀNG": "CÔNG TY TNHH KH INVEST", "TÊN MÔI GIỚI": "Phạm Thị Thu Hằng",
             "TÊN CTV": "Phạm Như Ngọc"},
        ])
        df_norm = _tx_row(source="JB", so_tieu_khoan="0001986886NM", ten_cong_ty="JB")
        out_df, mismatch_df, stats = mp._join_source(df_norm, master_jb, "JB")

        self.assertEqual(stats["suy_luan_tai_khoan_luu_ky"], 1)
        self.assertEqual(out_df.iloc[0]["Tài khoản lưu ký "], "050C986886")
        self.assertEqual(out_df.iloc[0]["Tên người quản lý"], "Phạm Như Ngọc")

    def test_still_unmatched_after_derivation_is_dropped_from_output_and_reports_derived_value(self):
        master_pb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "029C999999", "SỐ TIỂU KHOẢN": "029C999999.00",
             "TÊN KHÁCH HÀNG": "Ai Do", "TÊN MÔI GIỚI": "Ai Khac"},
        ])
        df_norm = _tx_row(source="PBSV", so_tieu_khoan="029C111111.00", ten_cong_ty="PBSV")
        out_df, mismatch_df, stats = mp._join_source(df_norm, master_pb, "PBSV")

        self.assertEqual(stats["khong_khop"], 1)
        self.assertEqual(stats["xoa_thieu_ten"], 1)
        self.assertEqual(len(out_df), 0)  # khong tra cuu duoc ten -> xoa ca dong khoi bang chinh
        self.assertIn("029C111111", mismatch_df.iloc[0]["Lý do"])

    def test_matched_but_missing_broker_name_is_kept_with_blank_cell(self):
        """Khach hang co that (khop tieu khoan) nhung Master khong co Ten Moi gioi ->
        VAN GIU dong trong bang chinh, chi de trong o "Tên người quản lý" (khong xoa
        dong, khac voi truong hop thieu Ten khach hang)."""
        master_pb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "029C222222", "SỐ TIỂU KHOẢN": "029C222222.00",
             "TÊN KHÁCH HÀNG": "Nguyen Van A", "TÊN MÔI GIỚI": None},
        ])
        df_norm = _tx_row(source="PBSV", so_tieu_khoan="029C222222.00", ten_cong_ty="PBSV")
        out_df, mismatch_df, stats = mp._join_source(df_norm, master_pb, "PBSV")

        self.assertEqual(stats["truc_tiep"], 1)
        self.assertEqual(stats["xoa_thieu_ten"], 0)
        self.assertEqual(len(out_df), 1)
        self.assertEqual(out_df.iloc[0]["Tên khách hàng"], "Nguyen Van A")
        self.assertEqual(out_df.iloc[0]["Tên người quản lý"], "")
        self.assertEqual(len(mismatch_df), 0)


class TestCustomerGroupAndClassification(unittest.TestCase):
    """'Nhóm khách hàng' / 'Phân loại khách hàng' phai lay tu danh sach khach hang, va
    dien mac dinh DEFAULT_UNKNOWN ("Không xác định") khi trong/None/NaN/chi co khoang
    trang, hoac khong khop duoc khach hang - khong duoc de trong o."""

    def test_filled_from_master_when_present(self):
        master_pb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "029C333333", "SỐ TIỂU KHOẢN": "029C333333.00",
             "TÊN KHÁCH HÀNG": "Nguyen Van B", "TÊN MÔI GIỚI": "Cao Ngọc Toàn",
             "NHÓM KHÁCH HÀNG": "VIX0", "PHÂN LOẠI KHÁCH HÀNG": "DEAL"},
        ])
        df_norm = _tx_row(source="PBSV", so_tieu_khoan="029C333333.00", ten_cong_ty="PBSV")
        out_df, _, _ = mp._join_source(df_norm, master_pb, "PBSV")

        self.assertEqual(out_df.iloc[0]["Nhóm khách hàng"], "VIX0")
        self.assertEqual(out_df.iloc[0]["Phân loại khách hàng"], "DEAL")

    def test_blank_or_missing_defaults_to_unknown(self):
        master_pb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "029C444444", "SỐ TIỂU KHOẢN": "029C444444.00",
             "TÊN KHÁCH HÀNG": "Nguyen Van C", "TÊN MÔI GIỚI": "Cao Ngọc Toàn",
             "NHÓM KHÁCH HÀNG": "   ", "PHÂN LOẠI KHÁCH HÀNG": None},
        ])
        df_norm = _tx_row(source="PBSV", so_tieu_khoan="029C444444.00", ten_cong_ty="PBSV")
        out_df, _, _ = mp._join_source(df_norm, master_pb, "PBSV")

        self.assertEqual(out_df.iloc[0]["Nhóm khách hàng"], mp.DEFAULT_UNKNOWN)
        self.assertEqual(out_df.iloc[0]["Phân loại khách hàng"], mp.DEFAULT_UNKNOWN)

    def test_missing_columns_in_master_default_to_unknown(self):
        """Sheet Master khong co san 2 cot nay (vd sheet JB cu) -> khong duoc crash,
        van phai dien mac dinh."""
        master_jb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "050C986886", "SỐ TIỂU KHOẢN": "0001986886MG",
             "TÊN KHÁCH HÀNG": "CÔNG TY TNHH KH INVEST", "TÊN MÔI GIỚI": "Phạm Thị Thu Hằng"},
        ])
        df_norm = _tx_row(source="JB", so_tieu_khoan="0001986886MG", ten_cong_ty="JB")
        out_df, _, _ = mp._join_source(df_norm, master_jb, "JB")

        self.assertEqual(out_df.iloc[0]["Nhóm khách hàng"], mp.DEFAULT_UNKNOWN)
        self.assertEqual(out_df.iloc[0]["Phân loại khách hàng"], mp.DEFAULT_UNKNOWN)

    def test_unmatched_customer_defaults_to_unknown(self):
        """Khong khop duoc khach hang trong Master, nhung dong van duoc giu lai (vi JB co
        san Ten khach hang/CTV tren chinh dong giao dich) -> Nhom/Phan loai van phai co
        gia tri mac dinh, khong de trong."""
        master_jb = _build_master_df([
            {"TÀI KHOẢN LƯU KÝ": "050C999999", "SỐ TIỂU KHOẢN": "0001999999MG",
             "TÊN KHÁCH HÀNG": "Ai Do", "TÊN MÔI GIỚI": "Ai Khac",
             "NHÓM KHÁCH HÀNG": "X", "PHÂN LOẠI KHÁCH HÀNG": "Y"},
        ])
        df_norm = _tx_row(
            source="JB", so_tieu_khoan="0001111111MG", ten_cong_ty="JB",
            ten_khach_hang_raw="Khach Le", ten_quan_ly_raw="CTV La",
        )
        out_df, mismatch_df, stats = mp._join_source(df_norm, master_jb, "JB")

        self.assertEqual(stats["khong_khop"], 1)
        self.assertEqual(len(out_df), 1)  # khong bi xoa vi da co Ten khach hang tu raw
        self.assertEqual(out_df.iloc[0]["Nhóm khách hàng"], mp.DEFAULT_UNKNOWN)
        self.assertEqual(out_df.iloc[0]["Phân loại khách hàng"], mp.DEFAULT_UNKNOWN)


if __name__ == "__main__":
    unittest.main()
