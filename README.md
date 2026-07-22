# Quản lý Công tác Cá nhân — Bản Web

Bản Streamlit + Supabase của phần mềm desktop, truy cập được từ mọi nơi,
có thể khai báo công việc, xem thống kê và tải báo cáo Word — kể cả
**Mẫu Tổng hợp BC của CBCC VP** theo tuần (mục 1: kết quả thực hiện,
mục 2: nhiệm vụ tuần tới).

## 1. Tạo dự án Supabase (miễn phí)

1. Vào https://supabase.com → **New project**. Ghi lại mật khẩu database.
2. Vào **SQL Editor** → dán toàn bộ nội dung file `supabase_schema.sql` → **Run**.
   Việc này tạo 3 bảng: `tasks`, `next_week_tasks`, `config`.
3. Vào **Project Settings → API**, lấy:
   - **Project URL**
   - **anon public key** (dùng cho app web — không dùng service_role)
   - **service_role key** (chỉ dùng 1 lần để di chuyển dữ liệu cũ, không đưa vào app)

## 2. Đưa dữ liệu cũ (cong_tac_v2.db) lên Supabase

Chạy trên máy cá nhân (không chạy trên Streamlit Cloud):

```bash
pip install supabase
python migrate_to_supabase.py --url https://xxxx.supabase.co --key <service_role_key> --db cong_tac_v2.db
```

Toàn bộ công việc và thông tin cấu hình (họ tên, chức vụ) sẽ được chuyển
sang Supabase, giữ nguyên dữ liệu đang theo dõi, không bị gián đoạn.

## 3. Đưa code lên GitHub

```bash
cd webapp
git init
git add .
git commit -m "Khởi tạo web app Quản lý Công tác Cá nhân"
git branch -M main
git remote add origin https://github.com/<ten-tai-khoan>/<ten-repo>.git
git push -u origin main
```

`.gitignore` đã loại trừ `secrets.toml` và file `.db` — không lo lộ khóa
hay dữ liệu cá nhân lên GitHub công khai. Nếu muốn repo private, tạo repo
ở chế độ Private trên GitHub trước khi push.

## 4. Deploy trên Streamlit Community Cloud (miễn phí)

1. Vào https://share.streamlit.io → **New app** → chọn repo vừa push, file chính `app.py`.
2. Vào **Advanced settings → Secrets**, dán nội dung sau (thay giá trị thật):

```toml
[supabase]
url = "https://xxxx.supabase.co"
key = "eyJhbGci...."   # anon public key

[auth]
app_password = "mat-khau-cua-ban"
```

   Mục `[auth]` là tùy chọn — thêm vào nếu muốn có màn hình đăng nhập
   bằng mật khẩu chung trước khi vào app (khuyên dùng vì app sẽ public URL).

3. Bấm **Deploy**. Sau vài phút app có địa chỉ dạng
   `https://<ten-app>.streamlit.app` — truy cập được từ điện thoại, máy tính,
   ở đâu cũng dùng được, dữ liệu đồng bộ realtime qua Supabase.

## 5. Sử dụng hằng ngày

- **Dashboard**: xem nhanh số liệu, biểu đồ tiến độ tháng/tuần.
- **Lịch & Quản lý công việc**: thêm/sửa/xóa công việc, lọc theo trạng thái, loại, khoảng ngày.
- **Nhiệm vụ tuần tới**: khai báo trước các việc dự kiến tuần sau — đây chính
  là nguồn dữ liệu cho mục 2 của Mẫu CBCC VP.
- **Báo cáo & Xuất file** — 3 tab:
  - **Mẫu Tổng hợp BC của CBCC VP**: chọn tuần → xuất đúng bố cục file
    mẫu gốc (mục 1 lấy công việc trong tuần, mục 2 lấy nhiệm vụ đã khai ở
    trang Nhiệm vụ tuần tới của **tuần kế tiếp**).
  - **Biểu 01: Phiếu tự đánh giá**: chọn kỳ (tuần/tháng/quý/6 tháng/9 tháng/
    năm/tùy chọn) → xuất đúng bố cục "Biểu 01: Cá nhân tự đánh giá" —
    mục I và bảng chi tiết 9 cột (Chủ trì, Phối hợp, Thời hạn yêu cầu,
    Thời gian hoàn thành, Đánh giá tiến độ tự động so sánh 2 mốc thời
    gian, Tự đánh giá chất lượng, Loại việc) được tổng hợp tự động từ
    dữ liệu công việc; mục II (đạo đức, lối sống) bạn tự viết trực tiếp
    trên form vì đây là nội dung tường thuật cá nhân, phần mềm không tự
    suy diễn thay bạn.
  - **Báo cáo tổng hợp theo kỳ**: giữ nguyên logic báo cáo tuần/tháng/quý/
    6 tháng/9 tháng/năm như bản desktop, kèm đề xuất xếp loại tự động.

## Ghi chú về các trường mới trong bảng `tasks`

Để khớp đúng các cột trong 2 mẫu báo cáo, bảng `tasks` có thêm:

**Cho Mẫu CBCC VP:**
- `lanh_dao_giao` — tên lãnh đạo giao việc
- `thoi_han_vb` — thời hạn của văn bản / thời hạn yêu cầu hoàn thành (dùng
  chung cho cả 2 mẫu)
- `lanh_dao_tham_dinh` — người thẩm định/duyệt kết quả

**Cho Biểu 01:**
- `chu_tri` — có phải người chủ trì thực hiện việc này không (mặc định: có)
- `phoi_hop` — đơn vị/người phối hợp thực hiện (nếu có)
- `loai_viec_bc` — "Kế hoạch/thường xuyên" hoặc "Phát sinh, đột xuất", dùng
  để tính số liệu tổng hợp ở mục I
- `tu_danh_gia_cl` — tự đánh giá chất lượng công tác tham mưu cho từng việc
  (để trống, phần mềm tự điền "Hoàn thành tốt nhiệm vụ" nếu việc đã hoàn thành)

Cột "Đánh giá tiến độ thực hiện" trong Biểu 01 được **tính tự động**: so
sánh `thoi_han_vb` (hạn yêu cầu) với ngày hoàn thành thực tế — sớm hơn ghi
"Vượt thời gian yêu cầu", đúng ngày ghi "Đúng thời gian yêu cầu", trễ hơn
ghi "Chậm so với yêu cầu".

Tất cả các trường trên không bắt buộc khi nhập việc — để trống vẫn xuất
báo cáo bình thường, chỉ là ô tương ứng sẽ để trống trong file Word.

**Nếu bạn đã tạo Supabase project trước khi có bản cập nhật này**, mở lại
SQL Editor và chạy riêng khối `ALTER TABLE ... ADD COLUMN IF NOT EXISTS`
ở cuối phần đầu file `supabase_schema.sql` để bổ sung 4 cột mới — không
ảnh hưởng dữ liệu đang có.

## Bảo trì

- Muốn "đánh thức" app khỏi trạng thái ngủ của Streamlit Cloud (khi không
  có người dùng lâu), có thể tái sử dụng GitHub Actions kiểu repo
  `nuoi_streamlit` đã có, thêm URL app mới vào danh sách.
- Toàn bộ dữ liệu nằm ở Supabase — có thể backup định kỳ bằng
  **Database → Backups** trong Supabase Dashboard.
