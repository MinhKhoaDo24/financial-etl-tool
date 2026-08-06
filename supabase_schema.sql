-- ============================================================================
-- PBSV Data Model — Supabase / PostgreSQL DDL
-- Generated from: Data Model.jpg
-- Entities: Cong_ty_chung_khoan, Nguoi_quan_ly, Chinh_sach, Phan_loai_khach_hang,
--           Nhom_khach_hang, Khach_hang, Co_phieu, Giao_dich, Phi_gia_han,
--           Bao_cao_thu_lai
-- Notes:
--   - Tables are created in FK-dependency order.
--   - "Mã ..." natural-key columns keep their business codes as PRIMARY KEY
--     where the diagram shows them underlined (Mã cổ phiếu, Mã giao dịch,
--     Mã khách hàng, Mã loại khách hàng, Mã nhóm khách hàng, Mã chính sách).
--   - Entities whose PK is drawn as "ID" (Công ty chứng khoán, Người quản lý,
--     Phí gia hạn, Báo cáo thu lãi) use a surrogate BIGINT IDENTITY key.
-- ============================================================================

begin;

-- ----------------------------------------------------------------------------
-- 1. Công ty chứng khoán
-- ----------------------------------------------------------------------------
create table if not exists cong_ty_chung_khoan (
    id                  bigint generated always as identity primary key,
    ma_dinh_danh_cong_ty varchar(50)  not null,
    ten_cong_ty         varchar(255) not null,
    constraint uq_cong_ty_ma_dinh_danh unique (ma_dinh_danh_cong_ty)
);

comment on table cong_ty_chung_khoan is 'Công ty chứng khoán';

-- ----------------------------------------------------------------------------
-- 2. Chính sách
-- ----------------------------------------------------------------------------
create table if not exists chinh_sach (
    ma_chinh_sach   varchar(50)  primary key,
    ten_chinh_sach  varchar(255) not null,
    lai_suat        numeric(10,4)  default 0 check (lai_suat >= 0),
    phi_giao_dich   numeric(10,4)  default 0 check (phi_giao_dich >= 0),
    phi_ung_truoc   numeric(10,4)  default 0 check (phi_ung_truoc >= 0),
    phi_gia_han     numeric(10,4)  default 0 check (phi_gia_han >= 0),
    lai_gia_han     numeric(10,4)  default 0 check (lai_gia_han >= 0),
    thoi_han        integer        check (thoi_han >= 0),
    han_muc_tong    numeric(18,2)  default 0 check (han_muc_tong >= 0),
    ty_le_vay       numeric(10,4)  check (ty_le_vay >= 0 and ty_le_vay <= 100)
);

comment on table chinh_sach is 'Chính sách lãi suất / phí áp dụng cho loại & nhóm khách hàng';

-- ----------------------------------------------------------------------------
-- 3. Người quản lý
-- ----------------------------------------------------------------------------
create table if not exists nguoi_quan_ly (
    id                      bigint generated always as identity primary key,
    ma_nguoi_quan_ly_ctv    varchar(50)  not null,
    ten_nguoi_quan_ly_ctv   varchar(255) not null,
    loai_nguoi_quan_ly      varchar(50)  not null check (loai_nguoi_quan_ly in ('Quản lý', 'CTV')),
    ma_cong_ty_chung_khoan  bigint references cong_ty_chung_khoan(id) on delete restrict,
    tinh_trang_hoat_dong    smallint     not null default 1 check (tinh_trang_hoat_dong in (0, 1)),
    constraint uq_nguoi_quan_ly_ma unique (ma_nguoi_quan_ly_ctv)
);

comment on table nguoi_quan_ly is 'Người quản lý / Cộng tác viên (CTV)';

create index if not exists idx_nguoi_quan_ly_cong_ty on nguoi_quan_ly (ma_cong_ty_chung_khoan);
create index if not exists idx_nguoi_quan_ly_tinh_trang on nguoi_quan_ly (tinh_trang_hoat_dong);

-- ----------------------------------------------------------------------------
-- 4. Phân loại khách hàng
-- ----------------------------------------------------------------------------
create table if not exists phan_loai_khach_hang (
    ma_loai_khach_hang  varchar(50)  primary key,
    ten_loai_khach_hang varchar(255) not null,
    phan_loai           varchar(100),
    mo_ta               text,
    ma_chinh_sach       varchar(50) references chinh_sach(ma_chinh_sach) on delete restrict
);

comment on table phan_loai_khach_hang is 'Phân loại khách hàng';

create index if not exists idx_phan_loai_khach_hang_chinh_sach on phan_loai_khach_hang (ma_chinh_sach);

-- ----------------------------------------------------------------------------
-- 5. Nhóm khách hàng
-- ----------------------------------------------------------------------------
create table if not exists nhom_khach_hang (
    ma_nhom_khach_hang  varchar(50)  primary key,
    ten_nhom_khach_hang varchar(255) not null,
    phan_nhom           varchar(100),
    ma_chinh_sach       varchar(50) references chinh_sach(ma_chinh_sach) on delete restrict,
    mo_ta               text
);

comment on table nhom_khach_hang is 'Nhóm khách hàng';

create index if not exists idx_nhom_khach_hang_chinh_sach on nhom_khach_hang (ma_chinh_sach);

-- ----------------------------------------------------------------------------
-- 6. Cổ phiếu
-- ----------------------------------------------------------------------------
create table if not exists co_phieu (
    ma_co_phieu                             varchar(20)  primary key,
    ten_doanh_nghiep                        varchar(255),
    gia_mo_cua_ngay_giao_dich_gan_nhat      numeric(18,2) check (gia_mo_cua_ngay_giao_dich_gan_nhat >= 0),
    gia_dong_cua_ngay_giao_dich_gan_nhat    numeric(18,2) check (gia_dong_cua_ngay_giao_dich_gan_nhat >= 0)
);

comment on table co_phieu is 'Cổ phiếu';

-- ----------------------------------------------------------------------------
-- 7. Khách hàng
-- ----------------------------------------------------------------------------
create table if not exists khach_hang (
    ma_khach_hang           varchar(50)  primary key,
    ten_khach_hang          varchar(255) not null,
    so_tai_khoan            varchar(50),
    ma_cong_ty_chung_khoan  bigint references cong_ty_chung_khoan(id) on delete restrict,
    ma_loai_khach_hang      varchar(50)  references phan_loai_khach_hang(ma_loai_khach_hang) on delete restrict,
    ma_nhom_khach_hang      varchar(50)  references nhom_khach_hang(ma_nhom_khach_hang) on delete restrict,
    ma_nguoi_quan_ly        varchar(50)  references nguoi_quan_ly(ma_nguoi_quan_ly_ctv) on delete restrict,
    nav                     numeric(18,2) default 0,
    du_no_goc               numeric(18,2) default 0 check (du_no_goc >= 0),
    du_no_lai               numeric(18,2) default 0 check (du_no_lai >= 0),
    ngay_toi_han_gan_nhat   date,
    ghi_chu                 text,
    tinh_trang_hoat_dong    smallint      not null default 1 check (tinh_trang_hoat_dong in (0, 1)),
    tong_du_no              numeric(18,2) default 0 check (tong_du_no >= 0),
    constraint uq_khach_hang_so_tai_khoan unique (so_tai_khoan)
);

comment on table khach_hang is 'Khách hàng';

create index if not exists idx_khach_hang_cong_ty on khach_hang (ma_cong_ty_chung_khoan);
create index if not exists idx_khach_hang_loai on khach_hang (ma_loai_khach_hang);
create index if not exists idx_khach_hang_nhom on khach_hang (ma_nhom_khach_hang);
create index if not exists idx_khach_hang_nguoi_quan_ly on khach_hang (ma_nguoi_quan_ly);
create index if not exists idx_khach_hang_tinh_trang on khach_hang (tinh_trang_hoat_dong);
create index if not exists idx_khach_hang_ngay_toi_han on khach_hang (ngay_toi_han_gan_nhat);

-- ----------------------------------------------------------------------------
-- 8. Giao dịch
-- ----------------------------------------------------------------------------
create table if not exists giao_dich (
    ma_giao_dich          varchar(100) primary key,
    ma_khach_hang         varchar(50) not null references khach_hang(ma_khach_hang) on delete restrict,
    ma_nguoi_quan_ly      varchar(50) references nguoi_quan_ly(ma_nguoi_quan_ly_ctv) on delete restrict,
    gia_tri_giao_dich     numeric(18,2) check (gia_tri_giao_dich >= 0),
    giao_dich_mua_ban     smallint not null check (giao_dich_mua_ban in (1, 2)), -- 1: Mua, 2: Bán
    ma_cp                 varchar(20) references co_phieu(ma_co_phieu) on delete restrict,
    thue_ban              numeric(18,2) default 0 check (thue_ban >= 0),
    ngay_giao_dich        date not null,
    phi_net               numeric(18,2) default 0 check (phi_net >= 0),
    khoi_luong_giao_dich  numeric(18,2) check (khoi_luong_giao_dich >= 0),
    gia_giao_dich         numeric(18,2) check (gia_giao_dich >= 0)
);

comment on table giao_dich is 'Giao dịch (1: Mua, 2: Bán)';

create index if not exists idx_giao_dich_khach_hang on giao_dich (ma_khach_hang);
create index if not exists idx_giao_dich_nguoi_quan_ly on giao_dich (ma_nguoi_quan_ly);
create index if not exists idx_giao_dich_ma_cp on giao_dich (ma_cp);
create index if not exists idx_giao_dich_ngay on giao_dich (ngay_giao_dich);

-- ----------------------------------------------------------------------------
-- 9. Phí gia hạn
-- ----------------------------------------------------------------------------
create table if not exists phi_gia_han (
    id                    bigint generated always as identity primary key,
    ngay                  date not null,
    ma_khach_hang         varchar(50) not null references khach_hang(ma_khach_hang) on delete restrict,
    phi_gia_han_du_thu    numeric(18,2) default 0 check (phi_gia_han_du_thu >= 0),
    phi_gia_han_thuc_thu  numeric(18,2) default 0 check (phi_gia_han_thuc_thu >= 0),
    lai                   numeric(18,2) default 0 check (lai >= 0)
);

comment on table phi_gia_han is 'Phí gia hạn';

create index if not exists idx_phi_gia_han_khach_hang on phi_gia_han (ma_khach_hang);
create index if not exists idx_phi_gia_han_ngay on phi_gia_han (ngay);

-- ----------------------------------------------------------------------------
-- 10. Báo cáo thu lãi
-- ----------------------------------------------------------------------------
create table if not exists bao_cao_thu_lai (
    id              bigint generated always as identity primary key,
    ngay_thu_lai    date not null,
    ma_khach_hang   varchar(50) not null references khach_hang(ma_khach_hang) on delete restrict,
    lai_vay         numeric(18,2) default 0 check (lai_vay >= 0),
    lai_ung_truoc   numeric(18,2) default 0 check (lai_ung_truoc >= 0)
);

comment on table bao_cao_thu_lai is 'Báo cáo thu lãi';

create index if not exists idx_bao_cao_thu_lai_khach_hang on bao_cao_thu_lai (ma_khach_hang);
create index if not exists idx_bao_cao_thu_lai_ngay on bao_cao_thu_lai (ngay_thu_lai);

commit;
