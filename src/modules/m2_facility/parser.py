import pandas as pd
import re

def extract_lc_cd_matrix(file_bytes):
    """Bóc tách Ma trận môn Lựa chọn - Chuyên đề (Đã thêm cơ chế phòng vệ lỗi thiếu Sheet)"""
    empty_cd = pd.DataFrame(columns=["ma_lop", "mon_chuyen_de", "tong_mon_cd"])
    empty_lc = pd.DataFrame(columns=["mon_lua_chon", "so_nhom_can_thiet"])
    
    try:
        xls = pd.ExcelFile(file_bytes)
        if "Môn LC-CĐ" not in xls.sheet_names:
            return empty_cd, empty_lc
            
        df_raw = pd.read_excel(xls, sheet_name="Môn LC-CĐ", header=None)
    except Exception:
        return empty_cd, empty_lc
        
    start_row = 0
    for idx, row in df_raw.iterrows():
        if str(row[0]).strip().lower() == "stt":
            start_row = idx
            break
            
    classes_data = []
    nhom_lc_dict = {}
    mon_lc_headers = ["Vật lý", "Hóa học", "Sinh học", "Tin học", "Địa lý", "GDKT&PL", "Công nghệ"]
    mon_cd_headers = ["Toán", "Vật lý", "Hóa học", "Sinh học", "Tin học", "Ngữ văn", "Lịch sử", "Địa lý"]
    
    for idx in range(start_row, len(df_raw)):
        row = df_raw.iloc[idx]
        if pd.isna(row[0]) or not str(row[0]).isdigit(): continue
            
        ten_lop = str(row[1]).strip()
        ds_cd = []
        for i, col_idx in enumerate(range(9, min(17, len(row)))):
            val = str(row[col_idx]).strip()
            if val in ["1", "1.0", "x", "X"]: ds_cd.append(mon_cd_headers[i])
                
        classes_data.append({"ma_lop": ten_lop, "mon_chuyen_de": ", ".join(ds_cd), "tong_mon_cd": len(ds_cd)})
        
        if ten_lop == "10A1":
            for i, col_idx in enumerate(range(2, min(9, len(row)))):
                val = row[col_idx]
                try:
                    so_nhom = int(float(val)) if pd.notna(val) else 0
                    if so_nhom > 0: nhom_lc_dict[mon_lc_headers[i]] = so_nhom
                except ValueError: pass
                
    df_cd = pd.DataFrame(classes_data) if classes_data else empty_cd
    df_lc = pd.DataFrame([{"mon_lua_chon": k, "so_nhom_can_thiet": v} for k, v in nhom_lc_dict.items()]) if nhom_lc_dict else empty_lc
                    
    return df_cd, df_lc

def extract_all_pinned_from_excel(file_bytes):
    """
    Quét động toàn bộ lưới TKB trong 2 sheet tkb sang / tkb chieu.
    Đã bỏ hoàn toàn bộ lọc từ khóa. Thu thập 100% các tiết khác rỗng (Bao gồm cả tiết "Nghỉ").
    """
    empty_df = pd.DataFrame(columns=["loai_tiet", "doi_tuong", "ngay_hoc", "buoi_hoc", "tiet_bat_dau", "so_tiet", "ghi_chu"])
    
    try:
        xls = pd.ExcelFile(file_bytes)
    except Exception:
        return empty_df
        
    results = []
    sheet_candidates = [s for s in xls.sheet_names if "tkb" in s.lower() and ("sang" in s.lower() or "chieu" in s.lower())]
    
    for sheet_name in sheet_candidates:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        except Exception:
            continue
            
        if len(df) < 3: continue
        classes_row = df.iloc[2]
        current_day = ""
        period = 0
        buoi = "Sáng" if "sang" in sheet_name.lower() else "Chiều"
        
        for idx, row in df.iterrows():
            if idx <= 2: continue
            
            cell_0 = str(row[0]).strip().upper()
            if cell_0.startswith("THỨ"):
                current_day = cell_0.title()
                period = 1 
                
            has_valid_data_on_this_row = False
            
            if current_day != "":
                if period > 5:
                    continue
                
                for c_idx, cell in enumerate(row):
                    if c_idx == 0 or c_idx >= len(classes_row): continue
                    
                    val = str(cell).strip()
                    class_name = str(classes_row[c_idx]).strip()
                    
                    if class_name != 'nan' and ('A' in class_name or '10' in class_name or '11' in class_name or '12' in class_name):
                        # CẬP NHẬT: Quét lấy tất cả các tiết khác rỗng, không lọc bỏ "nghỉ"
                        if val not in ['nan', '', '0', '0.0'] and len(val) > 1:
                            has_valid_data_on_this_row = True
                            
                            val_up = val.upper()
                            loai_tiet = "Cố định"
                            
                            if "TH_MOS" in val_up or "T.ANH NN" in val_up or "T.ANHNN" in val_up:
                                loai_tiet = "Ngoại vi"
                            elif "TNHN" in val_up or "TNHH" in val_up:
                                loai_tiet = "TNHN"
                            elif "-" in val or "LỚP GHÉP" in val_up or "LỰA CHỌN" in val_up: 
                                loai_tiet = "Lớp ghép"

                            results.append({
                                "loai_tiet": loai_tiet,
                                "doi_tuong": class_name,
                                "ngay_hoc": current_day,
                                "buoi_hoc": buoi,
                                "tiet_bat_dau": period,
                                "so_tiet": 1,
                                "ghi_chu": val
                            })
                
                if has_valid_data_on_this_row:
                    period += 1
                
    df_results = pd.DataFrame(results)
    
    if not df_results.empty:
        df_results = df_results.sort_values(by=["doi_tuong", "ngay_hoc", "buoi_hoc", "tiet_bat_dau"])
        merged = []
        curr = None
        for _, row in df_results.iterrows():
            if curr is None: curr = row.to_dict()
            else:
                if (curr["doi_tuong"] == row["doi_tuong"] and 
                    curr["ngay_hoc"] == row["ngay_hoc"] and 
                    curr["buoi_hoc"] == row["buoi_hoc"] and 
                    curr["ghi_chu"] == row["ghi_chu"] and 
                    curr["tiet_bat_dau"] + curr["so_tiet"] == row["tiet_bat_dau"]):
                    curr["so_tiet"] += 1
                else:
                    merged.append(curr)
                    curr = row.to_dict()
        if curr: merged.append(curr)
        return pd.DataFrame(merged)
        
    return empty_df

def extract_assignments_matrix(file_bytes):
    """
    Dùng thuật toán Regex thuần túy để bóc tách 100% lớp ghép.
    Đã thêm cơ chế phòng vệ khi sheet không tồn tại.
    """
    empty_df = pd.DataFrame(columns=["ma_lop", "mon", "gv_viet_tat", "raw_str"])
    
    try:
        xls = pd.ExcelFile(file_bytes)
        if "Mã lớp-Tên môn.GV" not in xls.sheet_names: 
            return empty_df
        
        df = pd.read_excel(xls, sheet_name="Mã lớp-Tên môn.GV", header=None)
    except Exception:
        return empty_df
        
    start_row = 0
    for idx, row in df.iterrows():
        if str(row[0]).strip().upper() == "LỚP":
            start_row = idx
            break

    assignments = []
    for idx in range(start_row + 2, len(df)):
        row = df.iloc[idx]
        lop = str(row[0]).strip()
        if not lop or lop == 'nan' or not ('A' in lop or '10' in lop or '11' in lop or '12' in lop): continue
            
        for c_idx in range(2, len(row)):
            val = str(row[c_idx]).strip()
            if val and val != 'nan' and '-' in val:
                if "Toán-Lý-Hóa" in val or "Toán-Văn-Địa" in val:
                    continue
                    
                parts = val.split('-')
                mon_lop = parts[0].strip() 
                gv_vt = parts[-1].strip()  
                
                match = re.search(r'\s*\d+\.\d+\s*', mon_lop)
                if match:
                    mon = mon_lop[:match.start()].strip()
                else:
                    mon = re.sub(r'\s*\d+A\d+\s*', '', mon_lop).strip()
                    if mon == mon_lop: 
                        mon = mon_lop.split(' ')[0]
                
                if mon in ['Li', 'Lí', 'Lý']: mon = 'Lí'
                elif mon in ['Thể', 'Thể dục']: mon = 'Thể dục'
                elif mon in ['Anh', 'T.Anh', 'Tiếng Anh']: mon = 'Anh'
                elif mon in ['Hoá', 'Hóa', 'Hóa học']: mon = 'Hoá'
                elif mon in ['GDCD', 'KTPL']: mon = 'KTPL'
                
                assignments.append({
                    "ma_lop": lop, 
                    "mon": mon, 
                    "gv_viet_tat": gv_vt, 
                    "raw_str": val
                })
                
    if not assignments:
        return empty_df
        
    return pd.DataFrame(assignments)