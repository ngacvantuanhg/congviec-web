"""
Di chuyển dữ liệu từ file cong_tac_v2.db (SQLite, bản desktop) lên Supabase.

CÁCH DÙNG:
1. Đã chạy xong supabase_schema.sql trên Supabase (SQL Editor).
2. Lấy "Project URL" và "service_role key" trong Supabase:
   Project Settings → API. (Dùng service_role — không dùng anon — vì
   script này chạy 1 lần từ máy cá nhân và cần quyền ghi đầy đủ.)
3. Đặt file cong_tac_v2.db cùng thư mục với script này.
4. Chạy:
      pip install supabase
      python migrate_to_supabase.py \
          --url https://xxxx.supabase.co \
          --key <service_role_key> \
          --db cong_tac_v2.db
"""
import argparse
import sqlite3
import sys

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True, help="Supabase Project URL")
    ap.add_argument("--key", required=True, help="Supabase service_role key")
    ap.add_argument("--db", default="cong_tac_v2.db", help="Đường dẫn file SQLite")
    args = ap.parse_args()

    try:
        from supabase import create_client
    except ImportError:
        sys.exit("Chưa cài thư viện supabase. Chạy: pip install supabase")

    sb = create_client(args.url, args.key)

    conn = sqlite3.connect(args.db)
    conn.row_factory = sqlite3.Row

    # ── config ──
    rows = conn.execute("SELECT * FROM config").fetchall()
    for r in rows:
        sb.table("config").upsert({"key": r["key"], "value": r["value"]}).execute()
    print(f"✔ Đã đồng bộ {len(rows)} dòng config")

    # ── tasks ──
    rows = conn.execute("SELECT * FROM tasks ORDER BY id").fetchall()
    payload = []
    for r in rows:
        d = dict(r)
        d.pop("id", None)  # để Supabase tự sinh id mới
        payload.append(d)
    if payload:
        # Chèn theo lô 200 dòng để tránh vượt giới hạn payload
        for i in range(0, len(payload), 200):
            sb.table("tasks").insert(payload[i:i + 200]).execute()
    print(f"✔ Đã chuyển {len(payload)} công việc sang bảng 'tasks' trên Supabase")

    conn.close()
    print("\nHoàn tất! Mở lại app Streamlit — dữ liệu cũ sẽ hiển thị đầy đủ.")


if __name__ == "__main__":
    main()
