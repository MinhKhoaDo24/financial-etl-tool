-- ============================================================================
-- Migration v2 — Cập nhật schema Supabase theo Data Model.png đã chỉnh sửa
-- Áp dụng cho database ĐÃ chạy supabase_schema.sql (v1) trước đó.
--
-- Thay đổi so với v1:
--   1. Khach_hang: đổi khóa chính  Mã khách hàng -> Số tài khoản
--                  (xóa cột ma_khach_hang)
--   2. Thêm bảng mới Tiểu khoản (Tieu_khoan): PK kép (so_tieu_khoan, so_tai_khoan),
--      FK so_tai_khoan -> Khach_hang(so_tai_khoan)
--   3. Giao_dich: đổi cột FK  ma_khach_hang -> so_tieu_khoan,
--                  tham chiếu sang Tieu_khoan thay vì Khach_hang
--   4. Phi_gia_han, Bao_cao_thu_lai: cột ma_khach_hang giữ nguyên tên nhưng
--      trỏ lại FK vào Khach_hang(so_tai_khoan) (khóa chính mới)
--
-- LƯU Ý TRƯỚC KHI CHẠY:
--   - Backup database trước khi chạy (Supabase Dashboard > Database > Backups).
--   - Script giả định giá trị cột so_tai_khoan hiện có = giá trị ma_khach_hang cũ
--     (đúng với dữ liệu do ETL tool sinh ra). Nếu không đúng, cần map lại dữ liệu
--     trước khi ALTER, nếu không bước "add constraint ... primary key" hoặc FK
--     ở dưới sẽ báo lỗi và toàn bộ transaction sẽ tự rollback (an toàn).
--   - Toàn bộ nằm trong 1 transaction: lỗi ở bất kỳ bước nào -> không thay đổi gì.
-- ============================================================================

begin;

-- ----------------------------------------------------------------------------
-- 1) Gỡ mọi FK đang trỏ vào Khach_hang (không hard-code tên constraint,
--    để không phụ thuộc vào việc Supabase/bạn có thể đã đổi tên chúng)
-- ----------------------------------------------------------------------------
do $$
declare r record;
begin
    for r in
        select conname, conrelid::regclass::text as tbl
        from pg_constraint
        where contype = 'f'
          and confrelid = 'khach_hang'::regclass
    loop
        execute format('alter table %s drop constraint %I', r.tbl, r.conname);
    end loop;
end $$;

drop index if exists idx_khach_hang_giao_dich;
drop index if exists idx_giao_dich_khach_hang;

-- ----------------------------------------------------------------------------
-- 2) Đổi khóa chính của Khách hàng: Mã khách hàng -> Số tài khoản
-- ----------------------------------------------------------------------------
alter table khach_hang drop constraint if exists khach_hang_pkey;
alter table khach_hang drop constraint if exists uq_khach_hang_so_tai_khoan;

alter table khach_hang alter column so_tai_khoan set not null;
alter table khach_hang add constraint khach_hang_pkey primary key (so_tai_khoan);

alter table khach_hang drop column if exists ma_khach_hang;

-- ----------------------------------------------------------------------------
-- 3) Bảng mới: Tiểu khoản — sub-account trực thuộc một Khách hàng
-- ----------------------------------------------------------------------------
create table if not exists tieu_khoan (
    so_tieu_khoan varchar(50) not null,
    so_tai_khoan  varchar(50) not null references khach_hang(so_tai_khoan) on delete restrict,
    constraint tieu_khoan_pkey primary key (so_tieu_khoan, so_tai_khoan),
    constraint uq_tieu_khoan_so_tieu_khoan unique (so_tieu_khoan)
);

comment on table tieu_khoan is 'Tiểu khoản (sub-account), trực thuộc một Khách hàng (Số tài khoản)';

create index if not exists idx_tieu_khoan_so_tai_khoan on tieu_khoan (so_tai_khoan);

-- Backfill: mỗi Khách hàng hiện có ít nhất 1 Tiểu khoản mặc định "00",
-- để các Giao_dich hiện hữu (repoint ở bước 4) có nơi tham chiếu tới.
-- Nếu dữ liệu Tiểu khoản thật đã có sẵn ở nơi khác, hãy thay bước backfill
-- này bằng INSERT dữ liệu thật trước khi chạy bước 4.
insert into tieu_khoan (so_tieu_khoan, so_tai_khoan)
select distinct kh.so_tai_khoan, kh.so_tai_khoan
from khach_hang kh
on conflict do nothing;

-- ----------------------------------------------------------------------------
-- 4) Giao_dich: đổi FK từ Khách hàng -> Tiểu khoản
-- ----------------------------------------------------------------------------
alter table giao_dich rename column ma_khach_hang to so_tieu_khoan;

-- Nhờ bước backfill ở (3) tạo Tiểu khoản mặc định có so_tieu_khoan = so_tai_khoan,
-- giá trị Giao_dich.so_tieu_khoan hiện có (vốn = Mã khách hàng cũ = Số tài khoản)
-- đã khớp sẵn với Tiểu khoản mặc định đó — không cần UPDATE dữ liệu thêm.
alter table giao_dich
    add constraint giao_dich_so_tieu_khoan_fkey
    foreign key (so_tieu_khoan) references tieu_khoan (so_tieu_khoan) on delete restrict;

create index if not exists idx_giao_dich_so_tieu_khoan on giao_dich (so_tieu_khoan);

-- ----------------------------------------------------------------------------
-- 5) Phi_gia_han & Bao_cao_thu_lai: trỏ lại FK vào khóa chính mới (so_tai_khoan)
-- ----------------------------------------------------------------------------
alter table phi_gia_han
    add constraint phi_gia_han_ma_khach_hang_fkey
    foreign key (ma_khach_hang) references khach_hang (so_tai_khoan) on delete restrict;

alter table bao_cao_thu_lai
    add constraint bao_cao_thu_lai_ma_khach_hang_fkey
    foreign key (ma_khach_hang) references khach_hang (so_tai_khoan) on delete restrict;

commit;
