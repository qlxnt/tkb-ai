import pandas as pd
import re
import unicodedata

def std_sub(s):
    s = str(s).strip()
    if s in ['Li', 'Lí', 'Vật lí', 'Vật lý']: return 'Lí'
    if s in ['Thể', 'Thể dục']: return 'Thể dục'
    if s in ['Anh', 'T.Anh', 'Tiếng Anh', 'Ngoại ngữ']: return 'Anh'
    if s in ['Hoá', 'Hóa', 'Hóa học']: return 'Hoá'
    if s in ['Sinh', 'Sinh học']: return 'Sinh'
    return s

def normalize_key(s):
    """
    Làm sạch chuỗi tuyệt đối để bắt khớp chính xác (Xóa khoảng trắng, đưa về chữ thường).
    Quy đổi uỷ/ủy về một chuẩn chung.
    """
    s = str(s).strip()
    if not s or s == 'nan': return ""
    s = unicodedata.normalize('NFC', s).replace('uỷ', 'ủy').replace('Uỷ', 'Ủy')
    return s.replace(' ', '').lower()

def generate_id_lop_mon(prefix, name, mapping_dict):
    """Tạo ID chuẩn hóa không phụ thuộc chuỗi gốc"""
    if name not in mapping_dict:
        clean_name = re.sub(r'[^\w\s]', '', str(name))
        clean_name = unicodedata.normalize('NFKD', clean_name).encode('ASCII', 'ignore').decode('utf-8')
        clean_name = clean_name.upper().strip().replace(' ', '_')
        mapping_dict[name] = f"{prefix}_{clean_name}"
    return mapping_dict[name]

def extract_raw_from_excel(file_path):
    xls = pd.ExcelFile(file_path)
    
    # =========================================================================
    # BƯỚC 1: ĐỌC SHEET 2 LẤY ĐỦ DANH SÁCH GIÁO VIÊN VÀ MÃ GV
    # =========================================================================
    try:
        df_sheet2 = pd.read_excel(xls, sheet_name=1, header=None)
    except Exception:
        df_sheet2 = pd.DataFrame()
        
    teachers_list = []
    gv_id_counter = 1
    
    if not df_sheet2.empty:
        for idx, row in df_sheet2.iterrows():
            if pd.notna(row[1]) and pd.notna(row[2]):
                full_name = str(row[1]).strip()
                code = str(row[2]).strip()
                
                # Bóc tách mã GV (VD: "Toán-Th. Mai" -> Tổ Môn: "Toán", Viết tắt: "Th. Mai")
                if '-' in code:
                    parts = code.rsplit('-', 1)
                    mon_goc = std_sub(parts[0].strip())
                    short_name = parts[1].strip()
                else:
                    mon_goc = ""
                    short_name = code
                    
                teachers_list.append({
                    'id_gv': f"GV_{gv_id_counter:03d}",
                    'ma_gv_cu': code, # Dùng đúng mã gốc làm chuẩn
                    'ho_ten': full_name,
                    'to_bo_mon': mon_goc,
                    'short_name': short_name,
                    'norm_code': normalize_key(code),
                    'norm_short': normalize_key(short_name),
                    'norm_full': normalize_key(full_name)
                })
                gv_id_counter += 1

    def find_teacher(raw_gv):
        """Hàm dò tìm GV chỉ dựa vào tên người (không quan tâm môn học trên cột)"""
        norm_gv = normalize_key(raw_gv)
        
        # 1. Khớp Tên viết tắt (short_name) hoặc Tên đầy đủ
        for t in teachers_list:
            if norm_gv == t['norm_short'] or norm_gv == t['norm_full']: return t
                
        # 2. Khớp chuỗi con
        for t in teachers_list:
            if norm_gv in t['norm_short'] or norm_gv in t['norm_full']: return t
                
        return None

    # =========================================================================
    # BƯỚC 2: QUÉT SHEET 1 LẤY DS_LOP_DAY, SỐ TIẾT LỚP & TỔNG TIẾT GV
    # =========================================================================
    df_sheet1 = pd.read_excel(xls, sheet_name=0)
    headers = df_sheet1.iloc[0].fillna('').astype(str)
    
    assignments = []
    lop_map, mon_map = {}, {}
    
    classes_totals_dict = {}
    grade_standard_periods = {'10': 0, '11': 0, '12': 0}
    current_periods = pd.Series([1]*len(headers))
    
    for idx, row in df_sheet1.iterrows():
        val0 = str(row.iloc[0]).strip()
        
        # Bắt "Số tiết QĐ" riêng biệt cho từng khối 10, 11, 12
        if val0.startswith("Số tiết QĐ"):
            current_periods = row
            match_grade = re.search(r'(10|11|12)', val0)
            if match_grade:
                grade = match_grade.group(1)
                for c in range(len(headers)-1, 0, -1):
                    if 'Tổng số' in str(headers.iloc[c]):
                        try: grade_standard_periods[grade] = int(round(float(row.iloc[c])))
                        except: pass
                        break
            continue
            
        lop = val0
        if not re.match(r'^(10|11|12)A\d+$', lop): continue
        
        id_lop = generate_id_lop_mon("LOP", lop, lop_map)
        if lop not in classes_totals_dict:
            classes_totals_dict[lop] = {'so_mon': 0, 'so_tiet': 0}
        
        # 2.1 Quét Cột GVCN (Chủ nhiệm)
        # CẬP NHẬT: Vẫn quét để biết GVCN, nhưng TUYỆT ĐỐI KHÔNG ĐẾM TIẾT
        gvcn = str(row.iloc[1]).strip()
        if gvcn and gvcn != 'nan':
            matched_gv = find_teacher(gvcn)
            assignments.append({
                'id_lop': id_lop, 'ten_lop': lop,
                'id_mon': "", 'ten_mon': '',
                'id_gv': matched_gv['id_gv'] if matched_gv else "",
                'ten_gv': matched_gv['ho_ten'] if matched_gv else gvcn,
                'ma_gv_cu': matched_gv['ma_gv_cu'] if matched_gv else gvcn,
                'so_tiet': 0, 'loai': 'GVCN'
            })
            # Không cộng vào classes_totals_dict để loại bỏ hoàn toàn việc tính toán tiết này
            
        # 2.2 Quét Hàng của Lớp (Các bộ môn)
        for col_idx in range(2, len(headers)):
            mon_header = str(headers.iloc[col_idx]).strip()
            if not mon_header or mon_header == 'nan' or 'Tổng số' in mon_header: continue
            
            cell_val = str(row.iloc[col_idx]).strip()
            
            # CHỈ XÉT CÁC Ô CÓ DỮ LIỆU MÃ GV (Hoặc chữ khác rỗng)
            if cell_val and cell_val != 'nan':
                try: p_count = int(round(float(current_periods.iloc[col_idx])))
                except: p_count = 1
                
                # Tách lấy tên viết tắt của GV (Bỏ đi môn học ghi dính liền nếu có)
                if '-' in cell_val:
                    raw_gv = cell_val.rsplit('-', 1)[1].strip()
                else:
                    raw_gv = cell_val
                
                raw_mon = std_sub(mon_header)
                # CẬP NHẬT: Loại bỏ hoàn toàn môn TNHN thuần túy (chỉ lấy TNHN 2 và TNHN 3)
                if 'TNHN' in mon_header.upper():
                    if '2' in mon_header: raw_mon = 'TNHN 2'
                    elif '3' in mon_header: raw_mon = 'TNHN 3'
                    else: continue # Tuyệt đối bỏ qua việc tạo tiết TNHN (Sinh hoạt)
                
                # Dựa vào mã GV để dò Sheet 2
                matched_gv = find_teacher(raw_gv)
                id_mon = generate_id_lop_mon("MH", raw_mon, mon_map)
                
                # Đóng gói thông tin quét được cho giáo viên đó
                assignments.append({
                    'id_lop': id_lop, 'ten_lop': lop,
                    'id_mon': id_mon, 'ten_mon': raw_mon,
                    'id_gv': matched_gv['id_gv'] if matched_gv else "",
                    'ten_gv': matched_gv['ho_ten'] if matched_gv else raw_gv,
                    'ma_gv_cu': matched_gv['ma_gv_cu'] if matched_gv else f"{raw_mon.split()[0]}-{raw_gv}",
                    'so_tiet': p_count, 'loai': 'Bộ Môn'
                })
                
                classes_totals_dict[lop]['so_mon'] += 1
                classes_totals_dict[lop]['so_tiet'] += p_count

    # Tổng kết tiết cho Lớp
    classes_totals = []
    for k, v in sorted(classes_totals_dict.items()):
        grade_match = re.match(r'^(10|11|12)', k)
        grade = grade_match.group(1) if grade_match else '10'
        classes_totals.append({
            'Lớp': k, 
            'Tổng số môn học': v['so_mon'],
            'Tổng tiết phân công': v['so_tiet'],
            'Số tiết QĐ (Khối)': grade_standard_periods.get(grade, 0)
        })
        
    return pd.DataFrame(assignments), pd.DataFrame(classes_totals)

def process_hybrid_ai(df_gv_raw):
    """
    Gom nhóm ds_lop_day và tổng số tiết TKB (tổng tiết dạy của từng GV) dựa trên dữ liệu quét.
    """
    teachers = {}
    
    for _, r in df_gv_raw.iterrows():
        id_gv = r.get('id_gv', '')
        if not id_gv or pd.isna(id_gv): continue
        
        ten_gv = r['ten_gv']
        mon = r['ten_mon']
        lop = r['ten_lop']
        so_tiet = r['so_tiet']
        loai = r['loai']
        
        if id_gv not in teachers:
            # Lấy tổ bộ môn gốc từ mã GV của Sheet 2
            to_bm_goc = ""
            if '-' in r['ma_gv_cu']: to_bm_goc = r['ma_gv_cu'].split('-')[0].strip()
                
            teachers[id_gv] = {
                'id_gv': id_gv,             
                'ma_gv': r['ma_gv_cu'], 
                'ho_ten': ten_gv,
                'to_bo_mon': to_bm_goc if to_bm_goc else mon,
                'lop_chu_nhiem': '',
                'ds_lop_day': {},
                'so_tiet_tkb': 0,
                'uu_tien_gv': ''
            }
        
        t = teachers[id_gv]
        
        if loai == 'GVCN':
            t['lop_chu_nhiem'] = lop
            # CẬP NHẬT: Tuyệt đối không tạo tiết TNHN và không đếm tiết cho GVCN ở danh sách dạy
            continue
            
        # CẬP NHẬT MỚI: Lấy toàn bộ tên môn học ghép vào phía trước mã giáo viên
        # Ví dụ: mon="TNHN 2", ma_gv_cu="TNHN -Thái" -> "TNHN 2 - TNHN -Thái"
        ma_gv_cu = r.get('ma_gv_cu', '')
        mon_hien_thi = f"{mon} - {ma_gv_cu}"

        # Cộng dồn số tiết nếu cùng dạy chung 1 môn cho 1 lớp
        key_lop = f"{mon_hien_thi} {lop}"
        if key_lop not in t['ds_lop_day']:
            t['ds_lop_day'][key_lop] = 0
        t['ds_lop_day'][key_lop] += so_tiet
        
        # Đếm tất cả các ô chứa mã GV -> Tổng số tiết dạy của từng GV
        t['so_tiet_tkb'] += so_tiet

    results = []
    for v in teachers.values():
        v['ds_lop_day'] = ', '.join([f"{ml} ({t}t)" for ml, t in v['ds_lop_day'].items()])
        results.append(v)
        
    return results