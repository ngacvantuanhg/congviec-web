-- ═══════════════════════════════════════════════════════════════
-- QUẢN LÝ CÔNG TÁC CÁ NHÂN — SCHEMA CHO SUPABASE (Postgres)
-- Ban Tuyên giáo và Dân vận Tỉnh ủy Tuyên Quang
-- Chạy toàn bộ file này trong Supabase Dashboard → SQL Editor
-- ═══════════════════════════════════════════════════════════════

-- Bảng công việc chính (tương ứng bảng "tasks" trong bản desktop,
-- bổ sung 3 cột mới phục vụ Mẫu báo cáo CBCC VP)
create table if not exists tasks (
    id            bigint generated always as identity primary key,
    title         text not null,
    mo_ta         text,
    loai          text default 'Khác',
    uu_tien       text default 'Bình thường',
    ngay_bd       date not null,
    ngay_kt       date not null,
    gio_bd        text,
    gio_kt        text,
    dia_diem      text,
    nguoi_pt      text,
    trang_thai    text default 'Chưa thực hiện',
    ket_qua       text,
    ghi_chu       text,
    -- Cột mới cho Mẫu Tổng hợp BC của CBCC VP:
    lanh_dao_giao       text,   -- Lãnh đạo giao nhiệm vụ
    thoi_han_vb         date,   -- Thời hạn của văn bản / Thời hạn yêu cầu hoàn thành
    lanh_dao_tham_dinh  text,   -- Lãnh đạo thẩm định, duyệt
    -- Cột mới cho Biểu 01: Cá nhân tự đánh giá:
    chu_tri             boolean default true,               -- Chủ trì thực hiện (X)
    phoi_hop            text,                                -- Đơn vị/người phối hợp
    loai_viec_bc        text default 'Kế hoạch/thường xuyên', -- Kế hoạch/thường xuyên | Phát sinh, đột xuất
    tu_danh_gia_cl      text,                                -- Tự đánh giá chất lượng công tác tham mưu
    -- Cột mới cho "Đánh giá cán bộ theo Công văn 1619-CV/TU":
    truc                 text default 'Trục 1',
    nhom_ab              text default 'A - Nhiệm vụ thường xuyên',
    can_cu_nv            text,
    so_luong_ke_hoach    integer default 1,
    nhom_kho             text default 'N2',
    so_luong_thuc_te     integer,
    he_so_chat_luong     numeric default 1.0,
    created_at    timestamptz default now()
);

-- Bảng "Nhiệm vụ tuần tới" — mục 2 của Mẫu Tổng hợp BC CBCC VP
create table if not exists next_week_tasks (
    id                   bigint generated always as identity primary key,
    noi_dung             text not null,
    phoi_hop             text,             -- đơn vị/cá nhân phối hợp (nếu có)
    thoi_gian_hoan_thanh text,             -- mô tả thời gian dự kiến hoàn thành
    tuan                 int  not null,    -- số tuần ISO
    nam                  int  not null,    -- năm
    created_at           timestamptz default now()
);

-- Bảng cấu hình (họ tên, chức vụ người dùng)
create table if not exists config (
    key   text primary key,
    value text
);

insert into config (key, value) values
    ('ten', 'Nguyễn Văn A'),
    ('chuc_vu', 'Chuyên viên')
on conflict (key) do nothing;

-- ─────────────────────────────────────────────────────────────
-- Nếu bạn ĐÃ chạy schema này trước đó (bảng tasks đã tồn tại),
-- chạy riêng khối ALTER TABLE dưới đây để bổ sung 4 cột mới phục
-- vụ Biểu 01 mà không mất dữ liệu đang có. An toàn khi chạy lại
-- nhiều lần (IF NOT EXISTS).
-- ─────────────────────────────────────────────────────────────
alter table tasks add column if not exists chu_tri boolean default true;
alter table tasks add column if not exists phoi_hop text;
alter table tasks add column if not exists loai_viec_bc text default 'Kế hoạch/thường xuyên';
alter table tasks add column if not exists tu_danh_gia_cl text;

-- Bổ sung cho "Đánh giá cán bộ theo Công văn 1619-CV/TU" (Quý III/2026 trở đi):
alter table tasks add column if not exists truc text default 'Trục 1';
alter table tasks add column if not exists nhom_ab text default 'A - Nhiệm vụ thường xuyên';
alter table tasks add column if not exists can_cu_nv text;
alter table tasks add column if not exists so_luong_ke_hoach integer default 1;
alter table tasks add column if not exists nhom_kho text default 'N2';
alter table tasks add column if not exists so_luong_thuc_te integer;
alter table tasks add column if not exists he_so_chat_luong numeric default 1.0;

-- Index phục vụ lọc theo khoảng ngày (dùng nhiều trong báo cáo/lịch)
create index if not exists idx_tasks_ngay on tasks (ngay_bd, ngay_kt);
create index if not exists idx_next_week on next_week_tasks (nam, tuan);

-- ─────────────────────────────────────────────────────────────
-- Bật Row Level Security + cho phép truy cập bằng anon/service key.
-- Vì đây là công cụ cá nhân dùng qua app Streamlit có mật khẩu
-- riêng (không public), ta mở quyền đọc/ghi cho key được cấu hình
-- trong app. Nếu muốn siết chặt hơn, thay policy "true" bằng điều
-- kiện theo auth.uid() khi tích hợp Supabase Auth.
-- ─────────────────────────────────────────────────────────────
alter table tasks enable row level security;
alter table next_week_tasks enable row level security;
alter table config enable row level security;

create policy "allow all - tasks" on tasks for all using (true) with check (true);
create policy "allow all - next_week_tasks" on next_week_tasks for all using (true) with check (true);
create policy "allow all - config" on config for all using (true) with check (true);
