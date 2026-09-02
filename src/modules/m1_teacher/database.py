import sqlite3
import os
import json
import pandas as pd

DB_DIR = os.path.join("data", "database")
DB_PATH = os.path.join(DB_DIR, "tkb_data.db")

def init_db():
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS teachers (
            ma_gv TEXT PRIMARY KEY,
            ho_ten TEXT NOT NULL,
            to_bo_mon TEXT,
            lop_chu_nhiem TEXT,
            ds_lop_day TEXT,
            mon_kiem_nhiem TEXT,
            uu_tien_gv TEXT,
            so_tiet_tkb INTEGER,
            so_tiet_quydinh INTEGER,
            chuc_vu TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_teachers_to_db(df: pd.DataFrame):
    init_db()
    conn = sqlite3.connect(DB_PATH)
    
    df_to_save = df.copy()
    for col in ['ds_lop_day', 'mon_kiem_nhiem']:
        if col in df_to_save.columns:
            df_to_save[col] = df_to_save[col].apply(lambda x: json.dumps(x, ensure_ascii=False) if isinstance(x, (list, dict)) else str(x or ''))

    df_to_save.to_sql("teachers", conn, if_exists="replace", index=False)
    conn.close()

def load_teachers_from_db() -> pd.DataFrame:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT * FROM teachers", conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df