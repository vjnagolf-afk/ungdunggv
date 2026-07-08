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
        
        # Tự động quy đổi điểm gõ liền ngay lập tức (88 thành 8.8)
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
    st.header("📊 HỆ THỐNG QUẢN LÝ ĐIỂM HỌC SINH (SMAS)")
    
    st.markdown("""
    <style>
    .tip-box { background-color: #FEF3C7; color: #92400E; padding: 10px; border-radius: 5px; font-weight: bold; border: 1px solid #F59E0B; margin-bottom: 15px;}
    </style>
    <div class="tip-box">
    🚀 CẨM NĂNG VÀO ĐIỂM THỜI GIAN THỰC (REAL-TIME):<br>
    - <b>Mẹo gõ siêu tốc:</b> Thầy gõ điểm (Ví dụ: gõ 85 hoặc 8.5) rồi nhấn <b>Enter</b>. Hệ thống sẽ tự động chuyển thành 8.5 và <b>tự động tính ngay điểm trung bình (TBM)</b> trực tiếp hiển thị lên màn hình mà không cần bấm nút lưu.<br>
    - <b>Phím điều hướng:</b> Nhấn Enter để kết thúc gõ 1 ô, sau đó sử dụng các phím mũi tên di chuyển cực kỳ mượt mà.
    </div>
    """, unsafe_allow_html=True)
    
    col_filter1, col_filter2, col_import = st.columns(3)
    
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
        
    if uploaded_smas:
        if st.button("🚀 Bắt đầu đồng bộ SMAS", type="secondary"):
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
                    st.success(f"🎉 Đã đồng bộ thành công dữ liệu và điểm của các lớp: {', '.join(set(imported_classes))}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi nhập liệu: {e}")
    # --- THUẬT TOÁN XỬ LÝ SỰ KIỆN TÍNH TOÁN REAL-TIME NGAY KHI GÕ XONG ---
    def handle_grade_change():
        if "realtime_editor" in st.session_state and st.session_state["realtime_editor"]:
            changes = st.session_state["realtime_editor"].get("edited_rows", {})
            if changes:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                for row_idx, updated_cols in changes.items():
                    # Lấy Mã HS ngầm từ DataFrame hiện tại dựa vào chỉ số dòng thay đổi
                    ma_hs = df_display.iloc[row_idx]["Mã HS"]
                    
                    # Đọc dữ liệu cũ có sẵn trong Database làm nền
                    cursor.execute("SELECT kttx1, kttx2, kttx3, kttx4, ktgk, ktck, comment_hk FROM grades WHERE student_code=?", (ma_hs,))
                    old_data = cursor.fetchone()
                    if old_data:
                        tx1, tx2, tx3, tx4, gk, ck, nx = old_data
                    else:
                        tx1 = tx2 = tx3 = tx4 = gk = ck = nx = None
                        
                    # Cập nhật đè các cột vừa được gõ mới và chuyển đổi 88 -> 8.8 lập tức
                    if "TX1" in updated_cols: tx1 = parse_score_smart(updated_cols["TX1"])
                    if "TX2" in updated_cols: tx2 = parse_score_smart(updated_cols["TX2"])
                    if "TX3" in updated_cols: tx3 = parse_score_smart(updated_cols["TX3"])
                    if "TX4" in updated_cols: tx4 = parse_score_smart(updated_cols["TX4"])
                    if "Điểm GK" in updated_cols: gk = parse_score_smart(updated_cols["Điểm GK"])
                    if "Điểm CK" in updated_cols: ck = parse_score_smart(updated_cols["Điểm CK"])
                    if "Nhận xét" in updated_cols: nx = updated_cols["Nhận xét"]
                    
                    # Tự động tính toán lại TBM HK ngay lập tức
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
    
    if not df_display.empty:
        st.markdown(f"##### 📝 BẢNG VÀO ĐIỂM LỚP {selected_class.upper()}")
        df_display.insert(0, "STT", range(1, len(df_display) + 1))
        
        score_cols = ["TX1", "TX2", "TX3", "TX4", "Điểm GK", "Điểm CK", "TBM HK"]
        for c in score_cols:
            df_display[c] = df_display[c].apply(lambda x: f"{float(x):.1f}" if pd.notna(x) and str(x).strip() != "" else "")
            
        # CẤU HÌNH CO NHỎ CỘT ĐIỂM TX VÀ ẨN HOÀN TOÀN MÃ HỌC SINH TĂNG DIỆN TÍCH CHO HỌ TÊN
        col_config = {
            "STT": st.column_config.NumberColumn("STT", width="small", disabled=True),
            "Mã HS": None, # ẨN HOÀN TOÀN CỘT MÃ HỌC SINH TRÊN GIAO DIỆN
            "Họ và tên": st.column_config.TextColumn("Họ và tên", width="large", disabled=True), # Mở rộng không gian hiển thị tên
            "TX1": st.column_config.TextColumn("TX1", width="small"), # Co nhỏ cột điểm
            "TX2": st.column_config.TextColumn("TX2", width="small"),
            "TX3": st.column_config.TextColumn("TX3", width="small"),
            "TX4": st.column_config.TextColumn("TX4", width="small"),
            "Điểm GK": st.column_config.TextColumn("GK", width="small"),
            "Điểm CK": st.column_config.TextColumn("CK", width="small"),
            "TBM HK": st.column_config.TextColumn("TBM", width="small", disabled=True),
            "Nhận xét": st.column_config.TextColumn("Nhận xét HK", width="medium")
        }
        
        # Gọi trình editor tương tác kích hoạt hàm lắng nghe sự kiện thay đổi Real-time
        edited_df = st.data_editor(
            df_display,
            column_order=["STT", "Họ và tên", "TX1", "TX2", "TX3", "TX4", "Điểm GK", "Điểm CK", "TBM HK", "Nhận xét"],
            use_container_width=True,
            num_rows="fixed",
            column_config=col_config,
            hide_index=True,
            height=500,
            key="realtime_editor",
            on_change=handle_grade_change # Tự động chạy tính toán ngay khi thầy gõ xong 1 ô
        )
        
        # Luồng xuất file Excel chuẩn mực không lỗi dữ liệu ngầm
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_df = edited_df.drop(columns=["STT"]) 
            export_df.to_excel(writer, index=False, sheet_name=f"{selected_class}")
            
        st.download_button(
            label="📥 Xuất File Điểm SMAS Hoàn Chỉnh", 
            data=output.getvalue(), 
            file_name=f"Diem_{selected_class}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ Chưa có dữ liệu học sinh. Vui lòng tải file SMAS (.xlsx) lên để đồng bộ.")
