import sqlite3
import os
import re
import pandas as pd
from ortools.sat.python import cp_model

DB_PATH = os.path.join("data", "database", "tkb_data.db")

def load_data():
    """
    Đọc dữ liệu trực tiếp từ bảng `teachers` (ds_lop_day, lop_chu_nhiem, ma_gv) và `pinned_slots`.
    """
    conn = sqlite3.connect(DB_PATH)
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='teachers';")
        if not cursor.fetchone():
            conn.close()
            raise Exception("Chưa có dữ liệu giáo viên trong CSDL. Vui lòng sang Module 1, nạp và bấm 'Bóc tách & Chuẩn hóa' file Excel trước khi chạy AI.")
            
        teachers_raw = pd.read_sql("SELECT * FROM teachers", conn)
        pinned = pd.read_sql("SELECT * FROM pinned_slots", conn)
    except Exception as e:
        conn.close()
        if "Chưa có dữ liệu" in str(e):
            raise e
        raise Exception(f"Lỗi đọc CSDL: {e}")
    conn.close()
    return teachers_raw, pinned

def is_valid_class(c): 
    return bool(re.match(r'^(10|11|12)A\d+$', str(c).strip().upper()))

def std_sub(s):
    s = str(s).strip()
    if s in ['Li', 'Lí', 'Vật lí', 'Vật lý']: return 'Lí'
    if s in ['Thể', 'Thể dục']: return 'Thể dục'
    if s in ['Anh', 'T.Anh', 'Tiếng Anh', 'Ngoại ngữ']: return 'Anh'
    if s in ['Hoá', 'Hóa', 'Hóa học']: return 'Hoá'
    if s in ['Sinh', 'Sinh học']: return 'Sinh'
    return s

def get_valid_pairs_and_mapping(teachers_raw):
    mapping = {} 
    subject_periods = {}
    
    for _, t in teachers_raw.iterrows():
        gv_code = str(t.get('ma_gv', '')).strip()
        if not gv_code or gv_code == 'nan': continue
        
        lop_cn = str(t.get('lop_chu_nhiem', '')).strip().upper()
        if is_valid_class(lop_cn):
            mapping[(lop_cn, "TNHN", gv_code)] = True
            subject_periods[(lop_cn, "TNHN", gv_code)] = 1
            
        ds_lop = str(t.get('ds_lop_day', ''))
        if ds_lop and ds_lop != 'nan':
            items = [x.strip() for x in ds_lop.split(',')]
            for item in items:
                if not item: continue
                m = re.search(r'^(.*?)\s+((?:10|11|12)A\d+)\s*\((\d+)[tT]\)$', item)
                if m:
                    raw_mon = m.group(1).strip()
                    
                    if ' - ' in raw_mon:
                        raw_mon = raw_mon.split(' - ')[0].strip()
                    elif '-' in raw_mon and ('TNHN' in raw_mon or 'CH ĐỀ' in raw_mon):
                        raw_mon = raw_mon.split('-')[0].strip()

                    lop = m.group(2).strip().upper()
                    so_tiet = int(m.group(3))
                    
                    mon = std_sub(raw_mon)
                    if is_valid_class(lop) and mon:
                        if (lop, mon, gv_code) not in subject_periods:
                            subject_periods[(lop, mon, gv_code)] = 0
                        subject_periods[(lop, mon, gv_code)] += so_tiet
                        mapping[(lop, mon, gv_code)] = True

    classes = list(set([k[0] for k in mapping.keys()]))
    class_subjects = {c: [] for c in classes} 
    
    teacher_codes = list(set([k[2] for k in mapping.keys()]))
    teacher_subjects = {t: [] for t in teacher_codes}
    
    for (c, mon, gv) in mapping.keys():
        if (mon, gv) not in class_subjects[c]: 
            class_subjects[c].append((mon, gv))
        if (c, mon) not in teacher_subjects[gv]: 
            teacher_subjects[gv].append((c, mon))
        
    return mapping, classes, class_subjects, teacher_subjects, subject_periods

def parse_teacher_prefs(teachers_raw):
    prefs = {
        'days_off': {}, 'session_off': {}, 'forbidden_periods': {},
        'allowed_days': {}, 'max_periods_per_session': {}, 
        'max_sessions_per_week': {}, 'must_teach_session': {}
    }
    if teachers_raw.empty: return prefs
    
    for _, row in teachers_raw.iterrows():
        t = str(row.get('ma_gv', '')).strip()
        if not t or t == 'nan': continue
            
        ho_ten = str(row.get('ho_ten', '')).lower()
        to_bm = str(row.get('to_bo_mon', '')).lower()
        mon_chinh = t.split('-')[0].lower() if '-' in t else t.lower()
        
        if t not in prefs['days_off']:
            prefs['days_off'][t] = []
            prefs['session_off'][t] = []
            prefs['forbidden_periods'][t] = []
        
        if ("toán" in mon_chinh or "toán" in to_bm) and "lợi" not in ho_ten:
            prefs['days_off'][t].append("Thứ Ba")
        if "tin" in mon_chinh or "tin" in to_bm:
            prefs['days_off'][t].append("Thứ Ba")
        if "văn" in mon_chinh or "văn" in to_bm:
            prefs['days_off'][t].append("Thứ Tư")
            
    return prefs

def get_teacher_days_off(teachers_raw):
    return parse_teacher_prefs(teachers_raw)['days_off']

def build_pinned_dict(pinned_df, classes, teachers_raw):
    days = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    periods = list(range(1, 8))
    classes_set = set(classes)
    pinned_dict = {}
    
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
                for _, t_row in teachers_raw.iterrows():
                    ma_gv = str(t_row.get('ma_gv', '')).strip()
                    ho_ten_chuan = str(t_row.get('ho_ten', '')).split()[-1]
                    if ma_gv in ghi_chu or ho_ten_chuan in ghi_chu:
                        t_pin = ma_gv
                        break
                
                for p_offset in range(so_tiet):
                    curr_p = act_p + p_offset
                    if curr_p in periods and d in days:
                        pinned_dict[(c, d, curr_p)] = {"loai": loai_tiet, "gv": t_pin, "label": ghi_chu, "mon": ""}
    return pinned_dict, list(classes_set)

def get_pinned_count(c, mon, gv, pinned_dict):
    cnt = 0
    for (pc, pd, pp), pin in pinned_dict.items():
        if pc == c and pin.get("gv") == gv:
            p_mon = pin.get("mon", "")
            p_label = pin.get("label", "")
            if p_mon == mon or mon in p_label or (pin.get("loai") == "TNHN" and mon == "TNHN"):
                cnt += 1
    return cnt

def format_display_name(subject, teacher_code):
    if not teacher_code: return f"{subject}-Trống"
    
    if 'TNHN' in subject or 'CH ĐỀ' in subject:
        return f"{subject} - {teacher_code}"
        
    clean_code = str(teacher_code).split('-')[-1].strip() if '-' in str(teacher_code) else str(teacher_code).strip()
    return f"{subject}-{clean_code}"

def cleanup_display_name(val):
    """
    HÀM DỌN DẸP CHO VÒNG 3:
    Cắt bỏ phần số và chuẩn hóa dấu phân cách cho các môn TNHN 2/3, CH ĐỀ 1/2/3.
    Ví dụ: TNHN 2 - KTPL-Duy -> TNHN-KTPL-Duy
           CH ĐỀ 1 - Toán-Yến -> CH ĐỀ-Toán-Yến
    """
    val_str = str(val)
    prefix = ""
    
    # Giữ lại các icon trạng thái
    if val_str.startswith('🔴 '):
        prefix = "🔴 "
        val_str = val_str[2:].strip()
    elif val_str.startswith('📌 '):
        prefix = "📌 "
        val_str = val_str[2:].strip()
    elif val_str.startswith('🟢 '):
        prefix = "🟢 "
        val_str = val_str[2:].strip()
        
    # Áp dụng biểu thức chính quy (regex) để nhận diện và thu gọn tên
    val_str = re.sub(r'^TNHN\s+[23]\s*-\s*', 'TNHN-', val_str)
    val_str = re.sub(r'^CH ĐỀ\s+[123]\s*-\s*', 'CH ĐỀ-', val_str)
        
    return prefix + val_str

def generate_df_from_vars(classes, days, periods, pinned_dict, class_subjects, X, solver, days_off):
    tkb_results = []
    for c in classes:
        for d in days:
            for p in periods:
                if (c, d, p) in pinned_dict:
                    pin = pinned_dict[(c, d, p)]
                    loai = pin["loai"]
                    t_pin = pin["gv"]
                    label = pin["label"]
                    is_off_day = (t_pin and d in days_off.get(t_pin, []))
                    
                    if loai == "Ngoại vi": 
                        prefix = "🔴" if is_off_day else "🟢"
                        tkb_results.append({"Lớp": c, "Giáo viên": f"{prefix} {label}", "Ngày": d, "Tiết": p})
                    elif loai == "TNHN": 
                        display = format_display_name("TNHN", t_pin) if t_pin else (label if label else "TNHN")
                        if is_off_day: display = f"🔴 {display}"
                        tkb_results.append({"Lớp": c, "Giáo viên": display, "Ngày": d, "Tiết": p})
                    elif loai == "Lớp ghép": 
                        tkb_results.append({"Lớp": c, "Giáo viên": f"⭐ Nhóm Lựa Chọn", "Ngày": d, "Tiết": p})
                    else:
                        found_mon = pin.get("mon")
                        if not found_mon:
                            for (mon, gv) in class_subjects.get(c, []):
                                if mon in label: found_mon = mon; break
                        if found_mon and t_pin:
                            display = format_display_name(found_mon, t_pin)
                            display = f"🔴 {display}" if is_off_day else f"📌 {display}"
                            tkb_results.append({"Lớp": c, "Giáo viên": display, "Ngày": d, "Tiết": p})
                        else:
                            clean_label = label.replace('Dữ liệu quét từ TKB (', '').replace(')', '')
                            prefix = "🔴" if is_off_day else "📌"
                            tkb_results.append({"Lớp": c, "Giáo viên": f"{prefix} {clean_label}", "Ngày": d, "Tiết": p})
                elif X is not None and solver is not None:
                    for (mon, gv) in class_subjects.get(c, []):
                        var_key = (c, mon, gv, d, p)
                        if var_key in X:
                            try:
                                val = solver.Value(X[var_key]) if hasattr(solver, 'Value') else X[var_key]
                                if val == 1:
                                    display = format_display_name(mon, gv)
                                    if d in days_off.get(gv, []):
                                        display = f"🔴 {display}"
                                    tkb_results.append({"Lớp": c, "Giáo viên": display, "Ngày": d, "Tiết": p})
                                    break
                            except: pass
                            
    df_result = pd.DataFrame(tkb_results, columns=["Lớp", "Giáo viên", "Ngày", "Tiết"])
    if not df_result.empty:
        df_result['Tiết_Sort'] = df_result['Tiết']
        df_result['Ngày_Sort'] = df_result['Ngày'].map({"Thứ Hai":1, "Thứ Ba":2, "Thứ Tư":3, "Thứ Năm":4, "Thứ Sáu":5})
        df_result = df_result.sort_values(by=['Lớp', 'Ngày_Sort', 'Tiết_Sort']).drop(columns=['Ngày_Sort', 'Tiết_Sort'])
    return df_result

def apply_safe_core_constraints(model, classes, days, periods, class_subjects, teacher_subjects, teacher_codes, pinned_dict, teachers_raw, objective_terms, is_round_3=False, rules=None):
    rules = rules or {}
    prefs = parse_teacher_prefs(teachers_raw)

    for c in classes:
        for d in days:
            for p in periods:
                c_vars = [X[(c, mon, gv, d, p)] for (mon, gv) in class_subjects.get(c, []) if (c, mon, gv, d, p) in X]
                c_pinned = 1 if (c, d, p) in pinned_dict else 0
                model.Add(sum(c_vars) <= max(0, 1 - c_pinned))

    for t in teacher_codes:
        for d in days:
            for p in periods:
                t_vars = [X[(c, mon, t, d, p)] for (c, mon) in teacher_subjects.get(t, []) if (c, mon, t, d, p) in X]
                t_pinned = sum(1 for (pc, pd, pp), pin in pinned_dict.items() if pd == d and pp == p and pin.get("gv") == t)
                model.Add(sum(t_vars) <= max(0, 1 - t_pinned))

    if rules.get("TEACHER_PREFS", True) and not is_round_3:
        for t, off_days in prefs['days_off'].items():
            for d in off_days:
                t_vars = [X[(c, mon, t, d, p)] for (c, mon) in teacher_subjects.get(t, []) for p in periods if (c, mon, t, d, p) in X]
                for var in t_vars:
                    objective_terms.append(var * -500)

    if rules.get("MAX_2_TIET", True) and not is_round_3:
        for c in classes:
            for (mon, gv) in class_subjects.get(c, []):
                for d in days:
                    for buoi_periods in [[1,2,3,4], [5,6,7]]:
                        m_vars = [X[(c, mon, gv, d, p)] for p in buoi_periods if (c, mon, gv, d, p) in X]
                        if m_vars:
                            pinned_m = sum(1 for (pc, pd, pp), pin in pinned_dict.items() if pc == c and pd == d and pp in buoi_periods and pin.get("gv") == gv and mon in pin.get("label", ""))
                            model.Add(sum(m_vars) <= max(0, 2 - pinned_m))

    if rules.get("NO_CACH_TIET", True) and not is_round_3:
        for c in classes:
            for (mon, gv) in class_subjects.get(c, []):
                for d in days:
                    for (p1, p2) in [(1, 3), (1, 4), (2, 4), (5, 7)]:
                        if (c, mon, gv, d, p1) in X and (c, mon, gv, d, p2) in X:
                            pin1 = any(pc == c and pd == d and pp == p1 and pin.get("gv") == gv for (pc, pd, pp), pin in pinned_dict.items())
                            pin2 = any(pc == c and pd == d and pp == p2 and pin.get("gv") == gv for (pc, pd, pp), pin in pinned_dict.items())
                            if not (pin1 and pin2):
                                model.Add(X[(c, mon, gv, d, p1)] + X[(c, mon, gv, d, p2)] <= 1)

def apply_morning_session_absolute_fill(model, X, classes, days, class_subjects, pinned_dict, objective_terms):
    morning_periods = [1, 2, 3, 4]
    for c in classes:
        for d in days:
            for p in morning_periods:
                c_vars = [X[(c, mon, gv, d, p)] for (mon, gv) in class_subjects.get(c, []) if (c, mon, gv, d, p) in X]
                c_pinned = 1 if (c, d, p) in pinned_dict else 0
                
                slack_m = model.NewBoolVar(f'slack_morning_{c}_{d}_{p}')
                model.Add(sum(c_vars) + c_pinned + slack_m == 1)
                objective_terms.append(slack_m * -100000)

def apply_no_afternoon_gap_priority(model, X, classes, days, class_subjects, pinned_dict, objective_terms):
    for c in classes:
        for d in days:
            y_vars = {}
            for p in [5, 6, 7]:
                c_vars = [X[(c, mon, gv, d, p)] for (mon, gv) in class_subjects.get(c, []) if (c, mon, gv, d, p) in X]
                c_pinned = 1 if (c, d, p) in pinned_dict else 0
                y_vars[p] = sum(c_vars) + c_pinned
            
            gap_var = model.NewBoolVar(f'gap_aft_{c}_{d}')
            model.Add(gap_var >= y_vars[5] + y_vars[7] - y_vars[6] - 1)
            objective_terms.append(gap_var * -50000)

def apply_teacher_session_optimization(model, X, days, teacher_subjects, teacher_codes, pinned_dict, objective_terms):
    for t in teacher_codes:
        for d in days:
            for session_periods in [[1, 2, 3, 4], [5, 6, 7]]:
                t_vars = [X[(c, mon, t, d, p)] for (c, mon) in teacher_subjects.get(t, []) for p in session_periods if (c, mon, t, d, p) in X]
                t_pinned = sum(1 for (pc, pd, pp), pin in pinned_dict.items() if pd == d and pp in session_periods and pin.get("gv") == t)
                
                if not t_vars and t_pinned == 0:
                    continue
                    
                sum_vars = sum(t_vars) + t_pinned
                
                b_0 = model.NewBoolVar(f'b0_{t}_{d}_{session_periods[0]}')
                b_1 = model.NewBoolVar(f'b1_{t}_{d}_{session_periods[0]}')
                b_gt_1 = model.NewBoolVar(f'bgt1_{t}_{d}_{session_periods[0]}')
                
                model.Add(sum_vars == 0).OnlyEnforceIf(b_0)
                model.Add(sum_vars == 1).OnlyEnforceIf(b_1)
                model.Add(sum_vars >= 2).OnlyEnforceIf(b_gt_1)
                model.Add(b_0 + b_1 + b_gt_1 == 1)
                
                objective_terms.append(b_1 * -2000)

def apply_core_double_period_logic(model, X, classes, days, class_subjects, subject_periods, pinned_dict, objective_terms):
    hard_core = ["toán", "văn", "anh", "thể dục"]
    soft_core = ["lí", "sinh", "hoá", "hóa"]
    
    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            mon_lower = mon.lower()
            is_hard = any(prio in mon_lower for prio in hard_core)
            is_soft = any(prio in mon_lower for prio in soft_core)
            
            if is_hard or is_soft:
                std_p = subject_periods.get((c, mon, gv), 2)
                if std_p >= 2:
                    pinned_pairs = 0
                    for d in days:
                        for p1, p2 in [(1, 2), (3, 4), (5, 6)]:
                            if (c, d, p1) in pinned_dict and (c, d, p2) in pinned_dict:
                                pin1 = pinned_dict[(c, d, p1)]
                                pin2 = pinned_dict[(c, d, p2)]
                                if pin1.get("gv") == gv and pin2.get("gv") == gv and mon in pin1.get("label", "") and mon in pin2.get("label", ""):
                                    pinned_pairs += 1
                    
                    pairs_vars = []
                    for d in days:
                        for p1, p2 in [(1, 2), (3, 4), (5, 6)]:
                            if (c, mon, gv, d, p1) in X and (c, mon, gv, d, p2) in X:
                                b_pair = model.NewBoolVar(f'core_pair_{c}_{mon}_{gv}_{d}_{p1}')
                                model.Add(X[(c, mon, gv, d, p1)] + X[(c, mon, gv, d, p2)] == 2).OnlyEnforceIf(b_pair)
                                model.Add(X[(c, mon, gv, d, p1)] + X[(c, mon, gv, d, p2)] <= 1).OnlyEnforceIf(b_pair.Not())
                                pairs_vars.append(b_pair)
                                
                    if pairs_vars:
                        slack_pair = model.NewBoolVar(f'slack_core_pair_{c}_{mon}_{gv}')
                        model.Add(sum(pairs_vars) + slack_pair * 10 >= max(0, 1 - pinned_pairs))
                        
                        if is_hard:
                            objective_terms.append(slack_pair * -50000)
                        else:
                            objective_terms.append(slack_pair * -1000)
                            for bv in pairs_vars:
                                objective_terms.append(bv * 5000)

def get_subj_weight(mon):
    m = mon.lower()
    if any(x in m for x in ["toán", "văn", "anh"]): return 50
    if any(x in m for x in ["lí", "hoá", "hóa", "sinh", "thể dục"]): return 40
    return 10

X = {}

def run_round_1(rules=None):
    global X
    rules = rules or {}
    teachers_raw, pinned_df = load_data()
    if teachers_raw.empty: return {"status": "ERROR", "message": "Dữ liệu CSDL Trống."}
    
    mapping, classes, class_subjects, teacher_subjects, subject_periods = get_valid_pairs_and_mapping(teachers_raw)
    pinned_dict, classes = build_pinned_dict(pinned_df, classes, teachers_raw)
    days_off = get_teacher_days_off(teachers_raw)

    days = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    periods = list(range(1, 8))
    teacher_codes = list(teacher_subjects.keys())
    
    core_subjects_r1 = ["toán", "văn", "anh", "thể dục"]

    model = cp_model.CpModel()
    X.clear()
    
    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            for d in days:
                if d in days_off.get(gv, []):
                    continue
                for p in periods:
                    X[(c, mon, gv, d, p)] = model.NewBoolVar(f'X_{c}_{mon}_{gv}_{d}_{p}')

    objective_terms = []
    apply_safe_core_constraints(model, classes, days, periods, class_subjects, teacher_subjects, teacher_codes, pinned_dict, teachers_raw, objective_terms, is_round_3=False, rules=rules)
    
    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            std_p = subject_periods.get((c, mon, gv), 2)
            pinned_p = get_pinned_count(c, mon, gv, pinned_dict)
            
            if any(core in mon.lower() for core in core_subjects_r1):
                target_X = max(0, std_p - pinned_p)
            else:
                target_X = 0 
            
            model.Add(sum(X[(c, mon, gv, d, p)] for d in days for p in periods if (c, mon, gv, d, p) in X) == target_X)
            
    apply_core_double_period_logic(model, X, classes, days, class_subjects, subject_periods, pinned_dict, objective_terms)
    apply_morning_session_absolute_fill(model, X, classes, days, class_subjects, pinned_dict, objective_terms)

    model.Maximize(sum(objective_terms))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20.0 
    solver.parameters.num_search_workers = 8 
    status = solver.Solve(model)
    
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        r1_vars = []
        for (c, mon, gv, d, p), var in X.items():
            if solver.Value(var) == 1:
                r1_vars.append((c, mon, gv, d, p))
                            
        df_result = generate_df_from_vars(classes, days, periods, pinned_dict, class_subjects, X, solver, days_off)
        return {"status": "SUCCESS", "data": df_result, "r1_vars": r1_vars}
        
    return {"status": "INFEASIBLE", "message": "Thuật toán bế tắc ở Vòng 1.", "diagnosis": []}

def run_round_2(r1_vars, rules=None):
    global X
    teachers_raw, pinned_df = load_data()
    mapping, classes, class_subjects, teacher_subjects, subject_periods = get_valid_pairs_and_mapping(teachers_raw)
    pinned_dict, classes = build_pinned_dict(pinned_df, classes, teachers_raw)
    days_off = get_teacher_days_off(teachers_raw)

    days = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    periods = list(range(1, 8))
    teacher_codes = list(teacher_subjects.keys())

    model = cp_model.CpModel()
    X.clear()
    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            for d in days:
                if d in days_off.get(gv, []):
                    continue
                for p in periods:
                    X[(c, mon, gv, d, p)] = model.NewBoolVar(f'X_{c}_{mon}_{gv}_{d}_{p}')

    objective_terms = []
    apply_safe_core_constraints(model, classes, days, periods, class_subjects, teacher_subjects, teacher_codes, pinned_dict, teachers_raw, objective_terms, is_round_3=False, rules=rules)
    apply_morning_session_absolute_fill(model, X, classes, days, class_subjects, pinned_dict, objective_terms)
    
    apply_no_afternoon_gap_priority(model, X, classes, days, class_subjects, pinned_dict, objective_terms)
    apply_teacher_session_optimization(model, X, days, teacher_subjects, teacher_codes, pinned_dict, objective_terms)

    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            std_p = subject_periods.get((c, mon, gv), 2)
            pinned_p = get_pinned_count(c, mon, gv, pinned_dict)
            
            if str(mon).strip().upper() == "TNHN":
                target_X = 0
            else:
                target_X = max(0, std_p - pinned_p)
                
            model.Add(sum(X[(c, mon, gv, d, p)] for d in days for p in periods if (c, mon, gv, d, p) in X) == target_X)
            
    apply_core_double_period_logic(model, X, classes, days, class_subjects, subject_periods, pinned_dict, objective_terms)

    r1_set = set(r1_vars)
    r1_bonus = []
    for (c, mon, gv, d, p), var in X.items():
        if (c, mon, gv, d, p) in r1_set:
            r1_bonus.append(var * 20000)
            model.AddHint(var, 1)
        else:
            model.AddHint(var, 0)

    model.Maximize(sum(objective_terms) + sum(r1_bonus))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 25.0 
    solver.parameters.num_search_workers = 8 
    status = solver.Solve(model)
    
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        all_vars = []
        for (c, mon, gv, d, p), var in X.items():
            if solver.Value(var) == 1:
                all_vars.append((c, mon, gv, d, p))
                            
        df_result = generate_df_from_vars(classes, days, periods, pinned_dict, class_subjects, X, solver, days_off)
        return {"status": "SUCCESS", "data": df_result, "r2_vars": all_vars}
        
    return {"status": "INFEASIBLE", "message": "Thuật toán bế tắc ở Vòng 2.", "diagnosis": []}

def run_round_3(r2_vars, rules=None):
    global X
    teachers_raw, pinned_df = load_data()
    mapping, classes, class_subjects, teacher_subjects, subject_periods = get_valid_pairs_and_mapping(teachers_raw)
    pinned_dict, classes = build_pinned_dict(pinned_df, classes, teachers_raw)
    days_off = get_teacher_days_off(teachers_raw)

    days = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]
    periods = list(range(1, 8))
    teacher_codes = list(teacher_subjects.keys())
    
    model = cp_model.CpModel()
    X.clear()
    
    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            for d in days:
                if d in days_off.get(gv, []):
                    continue
                for p in periods:
                    X[(c, mon, gv, d, p)] = model.NewBoolVar(f'X_{c}_{mon}_{gv}_{d}_{p}')

    objective_terms = []
    apply_safe_core_constraints(model, classes, days, periods, class_subjects, teacher_subjects, teacher_codes, pinned_dict, teachers_raw, objective_terms, is_round_3=True, rules=rules)
    apply_morning_session_absolute_fill(model, X, classes, days, class_subjects, pinned_dict, objective_terms)
    
    apply_no_afternoon_gap_priority(model, X, classes, days, class_subjects, pinned_dict, objective_terms)
    apply_teacher_session_optimization(model, X, days, teacher_subjects, teacher_codes, pinned_dict, objective_terms)
    
    for c in classes:
        for (mon, gv) in class_subjects.get(c, []):
            std_p = subject_periods.get((c, mon, gv), 2)
            pinned_p = get_pinned_count(c, mon, gv, pinned_dict)
            target = max(0, std_p - pinned_p)
            model.Add(sum(X[(c, mon, gv, d, p)] for d in days for p in periods if (c, mon, gv, d, p) in X) == target)

    apply_core_double_period_logic(model, X, classes, days, class_subjects, subject_periods, pinned_dict, objective_terms)

    r2_set = set(r2_vars)
    r2_bonus = []
    for (c, mon, gv, d, p), var in X.items():
        if (c, mon, gv, d, p) in r2_set:
            r2_bonus.append(var * 20000)
            model.AddHint(var, 1)
        else:
            model.AddHint(var, 0)
                        
    model.Maximize(sum(objective_terms) + sum(r2_bonus))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 25.0 
    solver.parameters.num_search_workers = 8 
    status = solver.Solve(model)
    
    diagnosis = []
    
    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        r3_vars = []
        for (c, mon, gv, d, p), var in X.items():
            if solver.Value(var) == 1:
                r3_vars.append((c, mon, gv, d, p))
                            
        actual_counts = {c: {(m, g): 0 for m, g in class_subjects.get(c, [])} for c in classes}
        
        for c in classes:
            for (mon, gv) in class_subjects.get(c, []):
                actual_counts[c][(mon, gv)] += get_pinned_count(c, mon, gv, pinned_dict)

        for (c, mon, gv, d, p) in r3_vars:
            if (mon, gv) in actual_counts.get(c, {}):
                actual_counts[c][(mon, gv)] += 1

        for c in classes:
            for (mon, gv) in class_subjects.get(c, []):
                target = subject_periods.get((c, mon, gv), 0)
                actual = actual_counts[c].get((mon, gv), 0)
                if actual < target:
                    diagnosis.append(f"❌ KHÔNG THỂ XẾP: Lớp {c} bị thiếu {target - actual} tiết môn {mon} (Mã GV: {gv}). Nguyên nhân: Xung đột với lịch cố định hoặc ép kín buổi sáng.")
            
            for d in ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu"]:
                for p in [1, 2, 3, 4]:
                    has_p = any((vc == c and vd == d and vp == p) for (vc, vmon, vgv, vd, vp) in r3_vars)
                    has_pinned_p = any((pc == c and pd == d and pp == p) for (pc, pd, pp), _ in pinned_dict.items())
                    if not has_p and not has_pinned_p:
                        diagnosis.append(f"⚠️ BỎ TRỐNG TIẾT SÁNG: Lớp {c} bị trống tiết {p} vào {d}. Nguyên nhân: Không còn môn nào phù hợp để lấp.")
                        
                has_5 = any((vc == c and vd == d and vp == 5) for (vc, vmon, vgv, vd, vp) in r3_vars) or any((pc == c and pd == d and pp == 5) for (pc, pd, pp), _ in pinned_dict.items())
                has_6 = any((vc == c and vd == d and vp == 6) for (vc, vmon, vgv, vd, vp) in r3_vars) or any((pc == c and pd == d and pp == 6) for (pc, pd, pp), _ in pinned_dict.items())
                has_7 = any((vc == c and vd == d and vp == 7) for (vc, vmon, vgv, vd, vp) in r3_vars) or any((pc == c and pd == d and pp == 7) for (pc, pd, pp), _ in pinned_dict.items())
                
                if has_5 and has_7 and not has_6:
                    diagnosis.append(f"⚠️ RĂNG LƯỢC BUỔI CHIỀU: Lớp {c} bị trống Tiết 6 vào {d} (Đang học Tiết 5 và 7). Nguyên nhân: AI phải né vi phạm các luật cứng khác.")

        audit_class_list = []
        for c in sorted(classes):
            target_c = sum(subject_periods.get((c, m, g), 0) for m, g in class_subjects.get(c, []))
            actual_c = sum(actual_counts[c].values())
            audit_class_list.append({
                "Lớp": c,
                "Yêu cầu (M1)": target_c,
                "Đã xếp (M4)": actual_c,
                "Trạng thái": "✅ Đủ" if actual_c >= target_c else f"❌ Thiếu {target_c - actual_c}"
            })
            
        audit_gv_list = []
        for gv in sorted(teacher_codes):
            ho_ten = gv
            target_gv = 0
            for _, row in teachers_raw.iterrows():
                if str(row.get('ma_gv', '')).strip() == gv:
                    ho_ten = str(row.get('ho_ten', '')).strip()
                    try: 
                        target_gv = int(row.get('so_tiet_tkb', 0))
                    except: 
                        target_gv = sum(subject_periods.get((c, m, gv), 0) for c in classes for m, g in class_subjects.get(c, []) if g == gv)
                    break
                    
            actual_gv = sum(actual_counts[c].get((m, gv), 0) for c in classes for m, g in class_subjects.get(c, []) if g == gv)
            
            audit_gv_list.append({
                "Mã GV": gv,
                "Họ Tên": ho_ten,
                "Phân công (M1)": target_gv,
                "Thực tế TKB (M4)": actual_gv,
                "Trạng thái": "✅ Đủ" if actual_gv >= target_gv else f"❌ Thiếu {target_gv - actual_gv}"
            })
            
        df_audit_class = pd.DataFrame(audit_class_list)
        df_audit_gv = pd.DataFrame(audit_gv_list)

        evaluated_X = {k: 1 for k in r3_vars}
        class DummySolver:
            def Value(self, val): return val
            
        df_result = generate_df_from_vars(classes, days, periods, pinned_dict, class_subjects, evaluated_X, DummySolver(), days_off)
        
        # BƯỚC HOÀN CHỈNH: Cắt bỏ phần số và dấu phân cách đã ghép ở Module 1
        df_result['Giáo viên'] = df_result['Giáo viên'].apply(cleanup_display_name)
        
        return {
            "status": "SUCCESS", 
            "data": df_result, 
            "diagnosis": list(set(diagnosis)),
            "audit_class": df_audit_class,
            "audit_gv": df_audit_gv
        }
        
    return {"status": "INFEASIBLE", "message": "Thuật toán bế tắc ở Vòng 3 do xung đột cứng.", "diagnosis": ["⚠️ Không tìm được phương án lấp đầy hoàn chỉnh."]}