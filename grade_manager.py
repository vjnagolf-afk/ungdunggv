import streamlit as st
import sqlite3
import pandas as pd
import os
import re
import io

DB_PATH = "teacher_assistant.db"

def setup_database_structure():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS students (
        student_code TEXT PRIMARY KEY, fullname TEXT, classroom TEXT
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS grades (
        student_code TEXT PRIMARY KEY,
        kttx1 REAL, kttx2 REAL, kttx3 REAL, kttx4 REAL,
        ktgk REAL, ktck REAL, tb REAL, comment_hk TEXT
    )
    """)
    conn.commit()
    conn.close()

def parse_score_smart(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "None"]: 
        return None
    try:
        val_str = str(val).replace(',', '.').strip()
        if val_str in ["10", "100", "10.0"]: return 10.0
        if val_str.isdigit() and len(val_str) == 2:
            return float(val_str) / 10.0
        val_float = float(val_str)
        if 0 <= val_float <= 10:
            return round(val_float, 1)
        return None
    except:
        return None
def render_grade_manager_section():
    setup_database_structure()
    st.header("📊 HỆ THỐNG QUẢN LÝ ĐIỂM HỌC SINH TẬP TRUNG (SMAS)")
    
    st.markdown("""
    <style>
    .tip-box { background-color: #FEF3C7; color: #92400E; padding: 10px; border-radius: 5px; font-weight: bold; border: 1px solid #F59E0B; margin-bottom: 15px;}
    div[data-testid="stForm"] button[kind="primary"] {
        background-color: #0284C7 !important; border-color: #0369A1 !important; color: white !important; font-weight: 900 !important; font-size: 16px !important; border-radius: 6px !important;
    }
    </style>
    <div class="tip-box">
    🚀 TÍNH NĂNG NÂNG CAO: QUẢN LÝ HÀNG LOẠT ĐA LỚP / ĐA SHEET<br>
    - <b>Nạp dữ liệu hàng loạt:</b> File SMAS tải lên có thể chứa nhiều Sheet (mỗi lớp 1 sheet). Hệ thống sẽ tự động quét và hút điểm đồng bộ toàn bộ.<br>
    - <b>Xuất dữ liệu hàng loạt:</b> Thầy tích chọn các lớp ở khung bên dưới để gom tất cả vào một file Excel duy nhất, tự động phân chia mỗi lớp một Sheet.
    </div>
    """, unsafe_allow_html=True)
    
    col_filter1, col_filter2, col_import = st.columns(3)
    
    with col_filter1:
        selected_grade = st.selectbox("Khối hiển thị:", ["Tất cả khối", "Khối 6", "Khối 7", "Khối 8", "Khối 9"])
        class_config = {
            "6": ["6A", "6B", "6C", "6D", "6E", "6F"],
            "7": ["7A", "7B", "7C", "7D", "7E", "7F"],
            "8": ["8A", "8B", "8C", "8D", "8E", "8F"],
            "9": ["9A", "9B", "9C", "9D", "9E", "9F", "9G"]
        }
        
    available_classes = []
    # --- KHẮC PHỤC LỖI: QUÉT ĐẦY ĐỦ LỚP KHI CHỌN "TẤT CẢ KHỐI" ---
    if selected_grade == "Tất cả khối":
        for classes in class_config.values():
            available_classes.extend(classes)
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT classroom FROM students WHERE classroom IS NOT NULL")
            db_classes = [str(row[0]) for row in cursor.fetchall() if row[0] is not None]
            for dc in db_classes:
                if dc not in available_classes: available_classes.append(dc)
            conn.close()
        except: pass
    else:
        grade_num = "".join([c for c in selected_grade if c.isdigit()])
        if grade_num in class_config:
            available_classes = list(class_config[grade_num])
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT classroom FROM students WHERE classroom LIKE ?", (f"{grade_num}%",))
            db_classes = [str(row[0]) for row in cursor.fetchall() if row[0] is not None]
            for dc in db_classes:
                if dc not in available_classes: available_classes.append(dc)
            conn.close()
        except: pass
        
    clean_classes = sorted(list(set([str(c).strip() for c in available_classes if c and str(c).strip() != ""])))
    final_class_list = ["Tất cả lớp"] + clean_classes if selected_grade == "Tất cả khối" else clean_classes
    
    with col_filter2:
        selected_class = st.selectbox("Lớp hiển thị nhập liệu:", final_class_list)
        
    with col_import:
        uploaded_smas = st.file_uploader("📥 Nhập file SMAS đa lớp (.xlsx)", type=["xlsx", "xls"], label_visibility="collapsed")
        
    if uploaded_smas:
        if st.button("🚀 Bắt đầu đồng bộ SMAS Hàng loạt", type="secondary", use_container_width=True):
            with st.spinner("Đang bóc tách dữ liệu đa Sheet..."):
                try:
                    excel_file = pd.ExcelFile(uploaded_smas)
                    all_sheets = excel_file.sheet_names
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()
                    imported_classes = []
                    
                    for sheet_name in all_sheets:
                        if str(sheet_name).lower() == "nan": continue
                        detected_class = str(sheet_name).strip()
                        df_head = pd.read_excel(uploaded_smas, sheet_name=sheet_name, nrows=12, header=None)
                        for r_idx, row in df_head.iterrows():
                            for val in row:
                                val_str = str(val).strip()
                                match = re.search(r'lớp\s*:?\s*(\s*[A-ZđĐtT\d]+)', val_str, re.IGNORECASE)
                                if match:
                                    detected_class = match.group(1).replace(" ", "").strip()
                                    break
                        df = pd.read_excel(uploaded_smas, sheet_name=sheet_name, skiprows=11)
                        df.columns = [str(c).strip() for c in df.columns]
                        
                        col_indices = {str(c): idx for idx, c in enumerate(df.columns)}
                        idx_ma = col_indices.get("Mã học sinh")
                        idx_ten = col_indices.get("Họ và tên")
                        
                        if idx_ma is None or idx_ten is None: continue
                        
                        idx_tx1 = idx_ten + 1
                        idx_tx2 = idx_ten + 2
                        idx_tx3 = idx_ten + 3
                        idx_tx4 = idx_ten + 4
                        idx_gk = col_indices.get("ĐĐG GK")
                        idx_ck = col_indices.get("ĐĐG CK")
                        idx_nx = col_indices.get("Nhận xét HKII") or col_indices.get("Nhận xét HKI") or col_indices.get("Nhận xét cả năm")
                        
                        has_data = False
                        for _, row in df.iterrows():
                            ma_hs = str(row.iloc[idx_ma]).strip() if not pd.isna(row.iloc[idx_ma]) else ""
                            ho_ten = str(row.iloc[idx_ten]).strip() if not pd.isna(row.iloc[idx_ten]) else ""
                            if not ma_hs or "STT" in ma_hs or ma_hs == "nan": continue
                            has_data = True
                            
                            tx1 = parse_score_smart(row.iloc[idx_tx1]) if idx_tx1 < len(row) else None
                            tx2 = parse_score_smart(row.iloc[idx_tx2]) if idx_tx2 < len(row) else None
                            tx3 = parse_score_smart(row.iloc[idx_tx3]) if idx_tx3 < len(row) else None
                            tx4 = parse_score_smart(row.iloc[idx_tx4]) if idx_tx4 < len(row) else None
                            gk = parse_score_smart(row.iloc[idx_gk]) if idx_gk is not None else None
                            ck = parse_score_smart(row.iloc[idx_ck]) if idx_ck is not None else None
                            nx = str(row.iloc[idx_nx]).strip() if idx_nx is not None and not pd.isna(row.iloc[idx_nx]) else None
                            
                            tx_scores = [x for x in [tx1, tx2, tx3, tx4] if x is not None]
                            tbm = None
                            if gk is not None and ck is not None:
                                tbm = round((sum(tx_scores) + (gk * 2) + (ck * 3)) / (len(tx_scores) + 5), 1)
                                
                            cursor.execute("INSERT OR REPLACE INTO students (student_code, fullname, classroom) VALUES (?, ?, ?)", (ma_hs, ho_ten, detected_class))
                            cursor.execute("""
                            INSERT OR REPLACE INTO grades (student_code, kttx1, kttx2, kttx3, kttx4, ktgk, ktck, tb, comment_hk) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (ma_hs, tx1, tx2, tx3, tx4, gk, ck, tbm, nx))
                            
                        if has_data: imported_classes.append(detected_class)
                    conn.commit()
                    conn.close()
                    st.success(f"🎉 Hệ thống đã đồng bộ hàng loạt {len(set(imported_classes))} lớp: {', '.join(set(imported_classes))}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi nhập liệu tệp đa sheet: {e}")
    conn = sqlite3.connect(DB_PATH)
    query = """
    SELECT s.student_code as [Mã HS], s.fullname as [Họ và tên], s.classroom as [Lớp],
           g.kttx1 as [TX1], g.kttx2 as [TX2], g.kttx3 as [TX3], g.kttx4 as [TX4],
           g.ktgk as [Điểm GK], g.ktck as [Điểm CK], g.tb as [TBM HK], g.comment_hk as [Nhận xét]
    FROM students s LEFT JOIN grades g ON s.student_code = g.student_code
    """
    params = []
    conditions = []
    if selected_class != "Tất cả lớp":
        conditions.append("s.classroom = ?")
        params.append(selected_class)
    elif selected_grade != "Tất cả khối":
        grade_num = "".join([c for c in selected_grade if c.isdigit()])
        conditions.append("s.classroom LIKE ?")
        params.append(f"%{grade_num}%")
        
    if conditions: query += " WHERE " + " AND ".join(conditions)
    query += " ORDER BY s.classroom, s.student_code ASC"
    df_display = pd.read_sql_query(query, conn, params=params)
    
    # --- QUÉT VÀ ÉP MẢNG CHỮ CHUẨN XUẤT TÙY Ý NHIỀU LỚP ---
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT classroom FROM students WHERE classroom IS NOT NULL ORDER BY classroom ASC")
    # Ép dữ liệu từ Tuple ('6A',) thành chuỗi chữ thô '6A' để st.multiselect cho phép tích chọn tùy ý
    all_db_classes = [str(row[0]).strip() for row in cursor.fetchall()]
    conn.close()
    
    if not df_display.empty:
        st.markdown(f"##### 📝 BẢNG VÀO ĐIỂM CHI TIẾT LỚP {selected_class.upper()}")
        df_display.insert(0, "STT", range(1, len(df_display) + 1))
        
        score_cols = ["TX1", "TX2", "TX3", "TX4", "Điểm GK", "Điểm CK", "TBM HK"]
        for c in score_cols:
            df_display[c] = df_display[c].apply(lambda x: f"{float(x):.1f}" if pd.notna(x) and str(x).strip() != "" else "")
            
        col_config = {
            "STT": st.column_config.NumberColumn("STT", width="small", disabled=True),
            "Mã HS": None,
            "Lớp": st.column_config.TextColumn("Lớp", width="small", disabled=True),
            "Họ và tên": st.column_config.TextColumn("Họ và tên", width="medium", disabled=True),
            "TX1": st.column_config.TextColumn("TX1", width="small"),
            "TX2": st.column_config.TextColumn("TX2", width="small"),
            "TX3": st.column_config.TextColumn("TX3", width="small"),
            "TX4": st.column_config.TextColumn("TX4", width="small"),
            "Điểm GK": st.column_config.TextColumn("GK", width="small"),
            "Điểm CK": st.column_config.TextColumn("CK", width="small"),
            "TBM HK": st.column_config.TextColumn("TBM", width="small", disabled=True),
            "Nhận xét": st.column_config.TextColumn("Nhận xét HK", width="medium")
        }
        
        with st.form("grade_entry_form", border=False):
            col_save_top, col_empty = st.columns(2)
            with col_save_top:
                submitted_top = st.form_submit_button("💾 LƯU ĐIỂM & TÍNH TBM (Nút đầu bảng)", type="primary", use_container_width=True)
            
            edited_df = st.data_editor(
                df_display,
                column_order=["STT", "Họ và tên", "Lớp", "TX1", "TX2", "TX3", "TX4", "Điểm GK", "Điểm CK", "TBM HK", "Nhận xét"],
                use_container_width=True,
                num_rows="fixed",
                column_config=col_config,
                hide_index=True,
                height=450,
                key="stable_grade_editor" 
            )
            st.markdown("<br>", unsafe_allow_html=True)
            submitted_bottom = st.form_submit_button("💾 LƯU ĐIỂM & TÍNH TBM (Nút cuối bảng)", type="primary", use_container_width=True)
            
        if submitted_top or submitted_bottom:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            for idx, row in edited_df.iterrows():
                ma_hs = df_display.iloc[idx]["Mã HS"]
                tx1 = parse_score_smart(row["TX1"])
                tx2 = parse_score_smart(row["TX2"])
                tx3 = parse_score_smart(row["TX3"])
                tx4 = parse_score_smart(row["TX4"])
                gk = parse_score_smart(row["Điểm GK"])
                ck = parse_score_smart(row["Điểm CK"])
                nx = row["Nhận xét"] if pd.notna(row["Nhận xét"]) else None
                
                tx_scores = [x for x in [tx1, tx2, tx3, tx4] if x is not None]
                tbm = None
                if gk is not None and ck is not None:
                    tbm = round((sum(tx_scores) + (gk * 2) + (ck * 3)) / (len(tx_scores) + 5), 1)
                    
                cursor.execute("""
                INSERT OR REPLACE INTO grades (student_code, kttx1, kttx2, kttx3, kttx4, ktgk, ktck, tb, comment_hk)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (ma_hs, tx1, tx2, tx3, tx4, gk, ck, tbm, nx))
            conn.commit()
            conn.close()
            st.success("🎉 Đã cập nhật và lưu trữ điểm số thành công!")
            st.rerun()
            
        # ==================== CHỨC NĂNG TÍCH CHỌN TÙY Ý ĐA LỚP XUẤT FILE TỔNG HỢP ====================
        st.markdown("---")
        st.markdown("📦 **TRÌNH XUẤT FILE EXCEL TÙY CHỌN HÀNG LOẠT (MỖI LỚP 1 SHEET):**")
        
        # Hộp multiselect cho phép thầy gõ tìm kiếm hoặc click chọn TÙY Ý nhiều lớp cùng lúc
        cac_lop_can_xuat = st.multiselect(
            "Chọn/Bỏ chọn tùy ý các lớp muốn kết xuất dữ liệu vào chung 1 file Excel:", 
            options=all_db_classes, 
            default=[selected_class] if selected_class in all_db_classes else all_db_classes[:1]
        )
        
        if cac_lop_can_xuat:
            output_bulk = io.BytesIO()
            with pd.ExcelWriter(output_bulk, engine='openpyxl') as writer:
                conn = sqlite3.connect(DB_PATH)
                for class_item in cac_lop_can_xuat:
                    q_bulk = """
                    SELECT s.student_code as [Mã học sinh], s.fullname as [Họ và tên],
                           g.kttx1 as [TX1], g.kttx2 as [TX2], g.kttx3 as [TX3], g.kttx4 as [TX4],
                           g.ktgk as [ĐĐG GK], g.ktck as [ĐĐG CK], g.tb as [TBM HK], g.comment_hk as [Nhận xét]
                    FROM students s LEFT JOIN grades g ON s.student_code = g.student_code WHERE s.classroom = ?
                    """
                    df_bulk = pd.read_sql_query(q_bulk, conn, params=[class_item])
                    if not df_bulk.empty:
                        # Đẩy dữ liệu lớp đã chọn vào một Sheet độc lập
                        df_bulk.to_excel(writer, index=False, sheet_name=f"Lớp {class_item}")
                conn.close()
                
            st.download_button(
                label=f"📥 Tải xuống File Excel Tổng hợp ({len(cac_lop_can_xuat)} Lớp Đã Chọn)", 
                data=output_bulk.getvalue(), 
                file_name=f"SMAS_Ket_Xuat_Hang_Loat.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("⚠️ Vui lòng chọn ít nhất một lớp để xuất dữ liệu Excel.")
    else:
        st.info("ℹ️ Chưa có dữ liệu học sinh trong hệ thống. Vui lòng nạp tệp Excel SMAS (.xlsx) ở đầu trang.")

    else:
        st.info("ℹ️ Chưa có dữ liệu học sinh trong hệ thống. Vui lòng nạp tệp Excel SMAS (.xlsx) ở đầu trang.")
