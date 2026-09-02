import sqlite3
import os
import pandas as pd
import json

DB_PATH = os.path.join("data", "database", "tkb_data.db")

def init_facility_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Bảng Quản lý Cơ sở vật chất (Phòng chức năng)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS rooms (
            ma_phong TEXT PRIMARY KEY,
            loai_phong TEXT,
            suc_chua INTEGER,
            ghi_chu TEXT
        )
    """)
    
    # Bảng Khung TKB Cố định (Hard Pinned Slots)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pinned_slots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loai_tiet TEXT,
            doi_tuong TEXT,
            ngay_hoc TEXT,
            buoi_hoc TEXT,
            tiet_bat_dau INTEGER,
            so_tiet INTEGER,
            ghi_chu TEXT
        )
    """)
    conn.commit()
    conn.close()

def seed_default_pinned_slots():
    """Nạp các nguyên tắc cứng mặc định của trường vào DB nếu chưa có."""
    init_facility_db()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM pinned_slots")
    if cursor.fetchone()[0] == 0:
        default_rules = [
            ("TNHN (GVCN)", "Toàn trường", "Thứ 2", "Sáng", 1, 1, "Cố định tiết 1 Sáng T2"),
            ("Lớp Ghép K10", "Khối 10", "Thứ 4", "Sáng", 1, 4, "Tất cả lớp 10 học Lựa chọn/Chuyên đề"),
            ("Lớp Ghép K11", "Khối 11", "Thứ 5", "Sáng", 1, 4, "Tất cả lớp 11 (trừ ngoại lệ) học Sáng T5"),
            ("Lớp Ghép 11A9,10,13,14", "11A9, 11A10, 11A13, 11A14", "Thứ 5", "Chiều", 1, 2, "Ngoại lệ lớp ghép Chiều T5")
        ]
        cursor.executemany("""
            INSERT INTO pinned_slots (loai_tiet, doi_tuong, ngay_hoc, buoi_hoc, tiet_bat_dau, so_tiet, ghi_chu)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, default_rules)
        conn.commit()
    conn.close()

def load_table(table_name: str) -> pd.DataFrame:
    init_facility_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df