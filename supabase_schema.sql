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
    thoi_han_vb         date,   -- Thời hạn của văn bản
    lanh_dao_tham_dinh  text,   -- Lãnh đạo thẩm định, duyệt
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
