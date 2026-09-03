import sqlite3
import os
import pandas as pd
import re

DB_PATH = os.path.join("data", "database", "tkb_data.db")

def is_valid_class(c): 
    return bool(re.match(r'^(10|11|12)A\d+$', str(c).strip().upper()))

def get_fixed_schedule(active_rules=None):
    """
    Quét DB Lịch Cố Định, xây dựng ma trận hiển thị và tự động
    đối chiếu với Bảng Nguyên Tắc để phát hiện lỗi từ người dùng ghim tay.
    """
    active_rules = active_rules or {}
    conn = sqlite3.connect(DB_PATH)
    try:
        pinned_df = pd.read_sql("SELECT * FROM pinned_slots", conn)
        teachers_df = pd.read_sql("SELECT * FROM teachers", conn)
    except Exception as e:
        return {"status": "ERROR", "message": f"Chưa có dữ liệu: {str(e)}"}
    finally:
        conn.close()

    days = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    periods = list(range(1, 8))
    
    tkb_results = []
    warnings = []
    schedule_dict = {} 
    classes_set = set()
    
    # 1. Quét và giải mã Lịch cố định (Tương đương logic Module 4)
    for _, row in pinned_df.iterrows():
        d = str(row.get('ngay_hoc', ''))
        try: start_p = int(row.get('tiet_bat_dau', 1))
        except: start_p = 1
        try: so_tiet = int(row.get('so_tiet', 1))
        except: so_tiet = 1
        
        dt = str(row.get('doi_tuong', '')).strip()
        loai = str(row.get('loai_tiet', '')).strip()
        ghi_chu = str(row.get('ghi_chu', '')).strip()
        
        target_classes = [dt] if is_valid_class(dt) else []
        
        if dt == "Toàn trường":
            for _, tr in teachers_df.iterrows():
                ds = str(tr.get('ds_lop_day', ''))
                for item in ds.split(','):
                    m = re.search(r'((?:10|11|12)A\d+)', item)
                    if m: target_classes.append(m.group(1))
            target_classes = list(set(target_classes))
            
        elif dt.startswith("Khối"):
            k = dt.replace("Khối ", "").strip()
            for _, tr in teachers_df.iterrows():
                ds = str(tr.get('ds_lop_day', ''))
                for item in ds.split(','):
                    m = re.search(r'((?:10|11|12)A\d+)', item)
                    if m and m.group(1).startswith(k): target_classes.append(m.group(1))
            target_classes = list(set(target_classes))
        
        if not target_classes: target_classes = [dt]
        
        buoi_hoc = str(row.get('buoi_hoc', 'Sáng')).strip()
        act_p = 1 if (loai == "TNHN" or "TNHN" in ghi_chu.upper()) else (start_p if buoi_hoc == "Sáng" else start_p + 4)
        if loai == "TNHN" or "TNHN" in ghi_chu.upper():
            d = "Thứ Hai"
            loai = "TNHN"
            
        t_pin = None
        for _, t_row in teachers_df.iterrows():
            ma_gv = str(t_row.get('ma_gv', '')).strip()
            ho_ten_chuan = str(t_row.get('ho_ten', '')).split()[-1]
            if ma_gv in ghi_chu or ho_ten_chuan in ghi_chu:
                t_pin = ma_gv
                break

        for c in target_classes:
            if not is_valid_class(c): continue
            classes_set.add(c)
            for offset in range(so_tiet):
                curr_p = act_p + offset
                if curr_p in periods and d in days:
                    schedule_dict[(c, d, curr_p)] = {"gv": t_pin, "loai": loai, "label": ghi_chu, "mon": ""}

    # 2. Định dạng TKB để hiển thị UI
    for (c, d, p), info in schedule_dict.items():
        val = info["label"]
        if info["loai"] == "Ngoại vi": val = f"🟢 {val}"
        elif info["loai"] == "TNHN": val = f"TNHN - {info['gv']}" if info['gv'] else "TNHN"
        elif info["loai"] == "Lớp ghép": val = "⭐ Nhóm Lựa Chọn"
        else: val = f"📌 {val}"
        tkb_results.append({"Lớp": c, "Giáo viên": val, "Ngày": d, "Tiết": p})

    df_result = pd.DataFrame(tkb_results, columns=["Lớp", "Giáo viên", "Ngày", "Tiết"])

    # =========================================================================
    # 3. MÁY QUÉT ĐỐI CHIẾU LUẬT (KIỂM TRA CÁC LỖI TỪ LỊCH GHIM TAY)
    # =========================================================================
    
    # LUẬT 1: Quét ngày nghỉ cố định của GV
    if active_rules.get("TEACHER_PREFS", True):
        days_off = {}
        for _, row in teachers_df.iterrows():
            t = str(row.get('ma_gv', '')).strip()
            to_bm = str(row.get('to_bo_mon', '')).lower()
            ho_ten = str(row.get('ho_ten', '')).lower()
            mon_chinh = t.split('-')[0].lower() if '-' in t else t.lower()
            offs = []
            if ("toán" in mon_chinh or "toán" in to_bm) and "lợi" not in ho_ten: offs.append("Thứ Ba")
            if "tin" in mon_chinh or "tin" in to_bm: offs.append("Thứ Ba")
            if "văn" in mon_chinh or "văn" in to_bm: offs.append("Thứ Tư")
            days_off[t] = offs
            
        for (c, d, p), info in schedule_dict.items():
            gv = info.get("gv")
            if gv and gv in days_off and d in days_off[gv]:
                warnings.append(f"Lớp {c}: Tiết cố định đang ép buộc giáo viên {gv} đi dạy vào ngày nghỉ ({d}).")

    # LUẬT 2: Quét cách tiết, tối đa 2 tiết, thủng lỗ 1-0-1...
    for c in classes_set:
        for d in days:
            morn_slots = {p: schedule_dict.get((c, d, p)) for p in [1,2,3,4] if (c, d, p) in schedule_dict}
            aft_slots = {p: schedule_dict.get((c, d, p)) for p in [5,6,7] if (c, d, p) in schedule_dict}
            
            for slots, session_name in [(morn_slots, "Sáng"), (aft_slots, "Chiều")]:
                if not slots: continue
                
                gv_counts = {}
                for p, info in slots.items():
                    gv = info.get("gv")
                    if gv: gv_counts[gv] = gv_counts.get(gv, 0) + 1
                
                if active_rules.get("MAX_2_TIET", True):
                    for gv, count in gv_counts.items():
                        if count > 2:
                            warnings.append(f"Lớp {c} ({d} {session_name}): Lịch cố định bắt buộc GV {gv} dạy quá 2 tiết/buổi.")
                
                if active_rules.get("NO_CACH_TIET", True):
                    for gv in gv_counts.keys():
                        p_list = sorted([p for p, info in slots.items() if info.get("gv") == gv])
                        if len(p_list) == 2:
                            if p_list in [[1,3], [2,4], [1,4], [5,7]]:
                                warnings.append(f"Lớp {c} ({d} {session_name}): Lịch cố định ép GV {gv} bị cách tiết ({p_list[0]} và {p_list[1]}).")

            if active_rules.get("NO_AFTERNOON_GAP", True):
                if 5 in aft_slots and 7 in aft_slots and 6 not in aft_slots:
                    warnings.append(f"Lớp {c} ({d} Chiều): Lịch cố định gây hổng rỗng Tiết 6 (Mô hình 1-0-1).")

    return {
        "status": "SUCCESS", 
        "data": df_result, 
        "message": "Đã tổng hợp hoàn thiện ma trận lịch cố định.", 
        "warnings": list(set(warnings))
    }