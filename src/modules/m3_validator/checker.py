import sqlite3
import os
import re
import pandas as pd

DB_PATH = os.path.join("data", "database", "tkb_data.db")

def is_valid_class(c): 
    return bool(re.match(r'^(10|11|12)A\d+$', str(c).strip().upper()))

def get_fixed_schedule():
    """
    Truy xuất và tái tạo ma trận Thời Khóa Biểu Cố Định (Module 3).
    ĐÃ VÁ LỖI: Rò rỉ kết nối Database (Database is locked). Đảm bảo conn.close() ở mọi tình huống.
    """
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        
        # 1. Kiểm tra sự tồn tại của các bảng dữ liệu
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='pinned_slots';")
        if not cursor.fetchone():
            conn.close() # VÁ LỖI: Phải đóng kết nối trước khi return sớm
            return {"status": "SUCCESS", "data": pd.DataFrame(columns=["Lớp", "Giáo viên", "Ngày", "Tiết"]), "message": "Hệ thống chưa có tiết cố định nào."}
            
        pinned_df = pd.read_sql("SELECT * FROM pinned_slots", conn)
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='teachers';")
        if cursor.fetchone():
            teachers_raw = pd.read_sql("SELECT * FROM teachers", conn)
        else:
            teachers_raw = pd.DataFrame()
            
        # 2. Thu thập danh sách toàn bộ Lớp học để dựng khung lưới TKB
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='danh_muc_lop';")
        all_classes = []
        if cursor.fetchone():
            classes_df = pd.read_sql("SELECT * FROM danh_muc_lop", conn)
            all_classes = classes_df['ten_lop'].tolist()
            
        if not all_classes and not teachers_raw.empty:
            for ds in teachers_raw['ds_lop_day'].dropna():
                for item in str(ds).split(','):
                    m = re.search(r'(10|11|12)A\d+', str(item))
                    if m: all_classes.append(m.group(0))
            for cn in teachers_raw['lop_chu_nhiem'].dropna():
                m = re.search(r'(10|11|12)A\d+', str(cn))
                if m: all_classes.append(m.group(0))
                
        all_classes = list(set(all_classes))
        
        # Đóng kết nối DB ngay sau khi đã trích xuất xong Pandas DataFrame để giải phóng file DB
        conn.close()
        conn = None

        days = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
        periods = list(range(1, 8))
        classes_set = set(all_classes)
        
        pinned_dict = {}
        
        # 3. Phân tích các tiết cố định (Pinned Slots)
        if not pinned_df.empty:
            for _, row in pinned_df.iterrows():
                d = str(row.get('ngay_hoc', ''))
                try: start_p = int(row.get('tiet_bat_dau', 1))
                except: start_p = 1
                try: so_tiet = int(row.get('so_tiet', 1))
                except: so_tiet = 1
                
                dt = str(row.get('doi_tuong', '')).strip()
                loai_tiet = str(row.get('loai_tiet', '')).strip()
                
                ghi_chu = str(row.get('ghi_chu', '')).strip()
                if ghi_chu == 'nan': ghi_chu = ""
                
                target_classes = [dt] if dt in classes_set else [c for c in classes_set if dt.startswith("Khối") and c.startswith(dt.replace("Khối ", "").strip())]
                if dt == "Toàn trường": target_classes = list(classes_set)
                if not target_classes: target_classes = [dt]
                
                for c in target_classes:
                    if not is_valid_class(c): continue
                    if c not in classes_set: classes_set.add(c)
                    
                    buoi_hoc = str(row.get('buoi_hoc', 'Sáng')).strip()
                    act_p = 1 if (loai_tiet == "TNHN" or "TNHN" in ghi_chu.upper()) else (start_p if buoi_hoc == "Sáng" else start_p + 4)
                    
                    if loai_tiet == "TNHN" or "TNHN" in ghi_chu.upper():
                        d = "Thứ Hai"
                        loai_tiet = "TNHN"
                        
                    t_pin = None
                    if not teachers_raw.empty:
                        for _, t_row in teachers_raw.iterrows():
                            ma_gv = str(t_row.get('ma_gv', '')).strip()
                            if ma_gv == 'nan': ma_gv = ""
                                
                            ho_ten_raw = str(t_row.get('ho_ten', '')).strip()
                            ho_ten_chuan = ho_ten_raw.split()[-1] if ho_ten_raw and ho_ten_raw != 'nan' else ""
                            
                            if (ma_gv and ma_gv in ghi_chu) or (ho_ten_chuan and ho_ten_chuan in ghi_chu):
                                t_pin = ma_gv
                                break
                    
                    for p_offset in range(so_tiet):
                        curr_p = act_p + p_offset
                        if curr_p in periods and d in days:
                            pinned_dict[(c, d, curr_p)] = {"loai": loai_tiet, "gv": t_pin, "label": ghi_chu, "mon": ""}

        # 4. Xuất mảng dữ liệu 2 chiều cho Dataframe 
        tkb_results = []
        for c in sorted(list(classes_set)):
            for d in days:
                for p in periods:
                    if (c, d, p) in pinned_dict:
                        pin = pinned_dict[(c, d, p)]
                        loai = pin["loai"]
                        t_pin = pin["gv"]
                        label = pin["label"]
                        label_lower = label.lower()
                        
                        if loai == "Ngoại vi": 
                            tkb_results.append({"Lớp": c, "Giáo viên": f"🟢 {label}", "Ngày": d, "Tiết": p})
                        elif loai == "TNHN": 
                            display = f"TNHN-{t_pin}" if t_pin else (label if label else "TNHN")
                            tkb_results.append({"Lớp": c, "Giáo viên": display, "Ngày": d, "Tiết": p})
                        elif loai == "Lớp ghép" or "lớp ghép" in label_lower or "lựa chọn" in label_lower: 
                            tkb_results.append({"Lớp": c, "Giáo viên": f"⭐ Nhóm Lựa Chọn", "Ngày": d, "Tiết": p})
                        else:
                            clean_label = label.replace('Dữ liệu quét từ TKB (', '').replace(')', '')
                            tkb_results.append({"Lớp": c, "Giáo viên": f"📌 {clean_label}", "Ngày": d, "Tiết": p})
                    else:
                        tkb_results.append({"Lớp": c, "Giáo viên": "", "Ngày": d, "Tiết": p})

        df_result = pd.DataFrame(tkb_results, columns=["Lớp", "Giáo viên", "Ngày", "Tiết"])
        if not df_result.empty:
            df_result['Tiết_Sort'] = df_result['Tiết']
            df_result['Ngày_Sort'] = df_result['Ngày'].map({"Thứ Hai":1, "Thứ Ba":2, "Thứ Tư":3, "Thứ Năm":4, "Thứ Sáu":5})
            df_result = df_result.sort_values(by=['Lớp', 'Ngày_Sort', 'Tiết_Sort']).drop(columns=['Ngày_Sort', 'Tiết_Sort'])

        return {"status": "SUCCESS", "data": df_result, "message": f"Đã quét và tải {len(pinned_dict)} tiết cố định."}

    except Exception as e:
        # Nếu có lỗi bất ngờ xảy ra, vẫn phải đảm bảo giải phóng Database
        if conn:
            conn.close()
        return {"status": "ERROR", "message": f"Lỗi nạp ma trận cố định: {e}"}