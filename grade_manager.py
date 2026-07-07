# grade_manager.py
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

# BỘ LỌC ĐIỂM THÔNG MINH (Xử lý thời gian thực)
def parse_score_smart(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "None"]: 
        return None
    try:
        val_str = str(val).replace(',', '.').strip()
        
        if val_str in ["10", "100", "10.0"]: return 10.0
        
        # Nhập 95 tự thành 9.5, 55 tự thành 5.5
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
    st.header("📈 HỆ THỐNG QUẢN LÝ ĐIỂM HỌC SINH (SMAS)")
    
    st.markdown("""
    <style>
    .tip-box { background-color: #FEF3C7; color: #92400E; padding: 10px; border-radius: 5px; font-weight: bold; border: 1px solid #F59E0B; margin-bottom: 15px;}
    </style>
    <div class="tip-box">
    💡 TÍNH NĂNG THỜI GIAN THỰC:<br>
    - Thầy gõ trực tiếp điểm (VD: 95, 10, 85). Bấm Enter hệ thống lập tức chuyển thành 9.5, 10.0, 8.5, <b>TỰ TÍNH TBM</b> và <b>TỰ LƯU</b>.<br>
    - Cột STT, Mã HS, Họ và tên, TBM HK đã được khóa an toàn (Không thể sửa tay).
    </div>
    """, unsafe_allow_html=True)

    col_filter1, col_filter2, col_import = st.columns([2, 2, 6])
    
    with col_filter1:
        selected_grade = st.selectbox("Khối:", ["Tất cả khối", "Khối 6", "Khối 7", "Khối 8", "Khối 9"])

    class_config = {
        "6": ["6A", "6B", "6C", "6D", "6E", "6F"],
        "7": ["7A", "7B", "7C", "7D", "7E", "7F"],
        "8": ["8A", "8B", "8C", "8D", "8E", "8F"],
        "9": ["9A", "9B", "9C", "9D", "9E", "9F", "9G"]
    }
    
    available_classes = []
    if selected_grade == "Tất cả khối":
        for classes in class_config.values():
            available_classes.extend(classes)
    else:
        grade_num = "".join([c for c in selected_grade if c.isdigit()])
        if grade_num in class_config:
            available_classes = class_config[grade_num]

    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT classroom FROM students WHERE classroom IS NOT NULL")
        db_classes = [row[0] for row in cursor.fetchall()]
        for dc in db_classes:
            if dc not in available_classes:
                available_classes.append(dc)
        conn.close()
    except:
        pass

    final_class_list = ["Tất cả lớp"] + sorted(list(set(available_classes)))

    with col_filter2:
        selected_class = st.selectbox("Lớp:", final_class_list)

    with col_import:
        uploaded_smas = st.file_uploader("📥 Nhập dữ liệu SMAS (.xlsx)", type=["xlsx", "xls"], label_visibility="collapsed")

    # XỬ LÝ IMPORT FILE SMAS CÓ CHỨA SẴN ĐIỂM
    if uploaded_smas:
        if st.button("🚀 Bắt đầu đồng bộ SMAS", type="primary"):
            with st.spinner("Đang bóc tách dữ liệu và hút điểm..."):
                try:
                    excel_file = pd.ExcelFile(uploaded_smas)
                    all_sheets = excel_file.sheet_names
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()
                    imported_classes = []
                    
                    for sheet_name in all_sheets:
                        if str(sheet_name).lower() == "nan": continue
                        
                        detected_class = str(sheet_name)
                        df_head = pd.read_excel(uploaded_smas, sheet_name=sheet_name, nrows=12, header=None)
                        for r_idx, row in df_head.iterrows():
                            for val in row:
                                val_str = str(val).strip()
                                match = re.search(r'lớp\s*:?\s*([6789]\s*[A-ZđĐtT\d]+)', val_str, re.IGNORECASE)
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
                                total_sum = sum(tx_scores) + (gk * 2) + (ck * 3)
                                total_coef = len(tx_scores) + 2 + 3
                                tbm = round(total_sum / total_coef, 1)

                            cursor.execute("INSERT OR REPLACE INTO students (student_code, fullname, classroom) VALUES (?, ?, ?)", (ma_hs, ho_ten, detected_class))
                            cursor.execute("""
                                INSERT OR REPLACE INTO grades (student_code, kttx1, kttx2, kttx3, kttx4, ktgk, ktck, tb, comment_hk) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (ma_hs, tx1, tx2, tx3, tx4, gk, ck, tbm, nx))
                            
                        if has_data: imported_classes.append(detected_class)

                    conn.commit()
                    conn.close()
                    if "grade_editor" in st.session_state: del st.session_state["grade_editor"]
                    st.success(f"✅ Đã đồng bộ thành công dữ liệu và điểm của các lớp: {', '.join(set(imported_classes))}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi nhập liệu: {e}")

    # LẤY DỮ LIỆU ĐỂ HIỂN THỊ
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT s.student_code as [Mã HS], s.fullname as [Họ và tên], 
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
    conn.close()

    # HIỂN THỊ & XỬ LÝ THỜI GIAN THỰC
    if not df_display.empty:
        st.markdown(f"##### 📝 BẢNG VÀO ĐIỂM LỚP {selected_class.upper()}")
        
        # Thêm Cột STT
        df_display.insert(0, "STT", range(1, len(df_display) + 1))
        
        # Format cột điểm hiển thị đẹp 10.0
        score_cols = ["TX1", "TX2", "TX3", "TX4", "Điểm GK", "Điểm CK", "TBM HK"]
        for c in score_cols:
            df_display[c] = df_display[c].apply(lambda x: f"{float(x):.1f}" if pd.notna(x) and str(x).strip() != "" else "")

        # Cấu hình chiều rộng cột và khóa các cột quan trọng
        col_config = {
            "STT": st.column_config.NumberColumn("STT", width="small", disabled=True),
            "Mã HS": st.column_config.TextColumn("Mã HS", width="medium", disabled=True),
            "Họ và tên": st.column_config.TextColumn("Họ và tên", width="large", disabled=True),
            "TX1": st.column_config.TextColumn("TX1", width="small"),
            "TX2": st.column_config.TextColumn("TX2", width="small"),
            "TX3": st.column_config.TextColumn("TX3", width="small"),
            "TX4": st.column_config.TextColumn("TX4", width="small"),
            "Điểm GK": st.column_config.TextColumn("Điểm GK", width="small"),
            "Điểm CK": st.column_config.TextColumn("Điểm CK", width="small"),
            "TBM HK": st.column_config.TextColumn("TBM HK", width="small", disabled=True),
            "Nhận xét": st.column_config.TextColumn("Nhận xét", width="medium")
        }

        # Bảng Data Editor Cao 700px
        edited_df = st.data_editor(
            df_display,
            use_container_width=True,
            num_rows="fixed",
            column_config=col_config,
            disabled=["STT", "Mã HS", "Họ và tên", "TBM HK"],
            hide_index=True,
            height=700,
            key="grade_editor"
        )
        
        # LẮNG NGHE SỰ KIỆN ENTER TỪ BÀN PHÍM VÀ LƯU NGAY LẬP TỨC
        editor_state = st.session_state.get("grade_editor", {})
        edits = editor_state.get("edited_rows", {})
        
        if edits:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            for row_idx_str, col_edits in edits.items():
                row_idx = int(row_idx_str)
                ma_hs = df_display.at[row_idx, "Mã HS"]
                
                row_data = edited_df.iloc[row_idx]
                tx1 = parse_score_smart(row_data["TX1"])
                tx2 = parse_score_smart(row_data["TX2"])
                tx3 = parse_score_smart(row_data["TX3"])
                tx4 = parse_score_smart(row_data["TX4"])
                gk = parse_score_smart(row_data["Điểm GK"])
                ck = parse_score_smart(row_data["Điểm CK"])
                nx = row_data["Nhận xét"] if pd.notna(row_data["Nhận xét"]) else None
                
                tx_scores = [x for x in [tx1, tx2, tx3, tx4] if x is not None]
                tbm = None
                # SMAS chỉ tính TBM nếu đã có điểm Giữa kỳ và Cuối kỳ
                if gk is not None and ck is not None:
                    total_sum = sum(tx_scores) + (gk * 2) + (ck * 3)
                    total_coef = len(tx_scores) + 2 + 3
                    tbm = round(total_sum / total_coef, 1)

                cursor.execute("""
                    UPDATE grades SET kttx1=?, kttx2=?, kttx3=?, kttx4=?, ktgk=?, ktck=?, tb=?, comment_hk=? 
                    WHERE student_code=?
                """, (tx1, tx2, tx3, tx4, gk, ck, tbm, nx, ma_hs))
            
            conn.commit()
            conn.close()
            
            # Xóa cache trên phiên làm việc và Rerun lại giao diện
            del st.session_state["grade_editor"]
            st.rerun()

        # Nút Xuất File bên dưới
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_df = edited_df.drop(columns=["STT"]) # Bỏ cột STT để cấu trúc giống hệt SMAS
            export_df.to_excel(writer, index=False, sheet_name=f"{selected_class}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📥 Xuất File Điểm SMAS Hoàn Chỉnh", 
            data=output.getvalue(), 
            file_name=f"Diem_{selected_class}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("💡 Chưa có dữ liệu học sinh. Vui lòng tải file SMAS (.xlsx) lên để đồng bộ.")
