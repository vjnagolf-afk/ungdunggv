# grade_manager.py
import streamlit as st
import sqlite3
import pandas as pd
import os
import re
import io

# Khởi tạo CSDL SQLite như cấu trúc gốc
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

def render_grade_manager_section():
    setup_database_structure()
    st.header("📈 HỆ THỐNG QUẢN LÝ ĐIỂM HỌC SINH (SMAS)")
    
    st.markdown("""
    <style>
    .tip-box { background-color: #FEF3C7; color: #92400E; padding: 10px; border-radius: 5px; font-weight: bold; border: 1px solid #F59E0B; margin-bottom: 15px;}
    </style>
    <div class="tip-box">💡 CƠ CHẾ CHẤM ĐIỂM: Thầy nhập trực tiếp điểm vào bảng phía dưới. Điểm Trung Bình Môn (TBM) sẽ được tự động tính và lưu vào cơ sở dữ liệu ngay lập tức.</div>
    """, unsafe_allow_html=True)

    # Đọc danh sách lớp từ Database
    conn = sqlite3.connect(DB_PATH)
    df_classes = pd.read_sql_query("SELECT DISTINCT classroom FROM students WHERE classroom IS NOT NULL", conn)
    class_list = ["Tất cả lớp"] + sorted(df_classes['classroom'].tolist())
    conn.close()

    # Thanh điều khiển phía trên
    col_filter1, col_filter2, col_import = st.columns([2, 2, 6])
    with col_filter1:
        selected_grade = st.selectbox("Khối:", ["Tất cả khối", "Khối 6", "Khối 7", "Khối 8", "Khối 9"])
    with col_filter2:
        selected_class = st.selectbox("Lớp:", class_list)
    with col_import:
        uploaded_smas = st.file_uploader("📥 Nhập dữ liệu SMAS (.xlsx)", type=["xlsx", "xls"], label_visibility="collapsed")

    # XỬ LÝ: Nhập dữ liệu SMAS từ Excel
    if uploaded_smas:
        if st.button("🚀 Bắt đầu đồng bộ SMAS", type="primary"):
            with st.spinner("Đang bóc tách dữ liệu..."):
                try:
                    excel_file = pd.ExcelFile(uploaded_smas)
                    all_sheets = excel_file.sheet_names
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()

                    imported_classes = []
                    
                    for sheet_name in all_sheets:
                        if str(sheet_name).lower() == "nan": continue
                        
                        # Quét tìm tên lớp từ các dòng đầu của Excel
                        detected_class = str(sheet_name)
                        df_head = pd.read_excel(uploaded_smas, sheet_name=sheet_name, nrows=12, header=None)
                        for r_idx, row in df_head.iterrows():
                            for val in row:
                                val_str = str(val).strip()
                                match = re.search(r'lớp\s*:?\s*([6789]\s*[A-ZđĐtT\d]+)', val_str, re.IGNORECASE)
                                if match:
                                    detected_class = match.group(1).replace(" ", "").strip()
                                    break

                        # Đọc dữ liệu điểm, bỏ qua 11 dòng tiêu đề của SMAS
                        df = pd.read_excel(uploaded_smas, sheet_name=sheet_name, skiprows=11)
                        df.columns = [str(c).strip() for c in df.columns]
                        
                        col_indices = {str(c): idx for idx, c in enumerate(df.columns)}
                        idx_ma = col_indices.get("Mã học sinh")
                        idx_ten = col_indices.get("Họ và tên")
                        
                        if idx_ma is None or idx_ten is None: continue
                            
                        has_data = False
                        for _, row in df.iterrows():
                            ma_hs = str(row.iloc[idx_ma]).strip() if not pd.isna(row.iloc[idx_ma]) else ""
                            ho_ten = str(row.iloc[idx_ten]).strip() if not pd.isna(row.iloc[idx_ten]) else ""
                            
                            if not ma_hs or "STT" in ma_hs or ma_hs == "nan": continue
                            has_data = True
                            
                            # Lưu vào DB
                            cursor.execute("INSERT OR REPLACE INTO students (student_code, fullname, classroom) VALUES (?, ?, ?)", (ma_hs, ho_ten, detected_class))
                            cursor.execute("INSERT OR IGNORE INTO grades (student_code) VALUES (?)", (ma_hs,))
                            
                        if has_data: imported_classes.append(detected_class)

                    conn.commit()
                    conn.close()
                    st.success(f"✅ Đã đồng bộ thành công danh sách các lớp: {', '.join(set(imported_classes))}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi nhập liệu: {e}")

    # Lấy dữ liệu hiển thị
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

    # Hiển thị và Chỉnh sửa trực tiếp
    if not df_display.empty:
        st.markdown("##### 📝 BẢNG VÀO ĐIỂM (TỰ ĐỘNG LƯU)")
        
        # Streamlit Data Editor thay thế cho Tkinter Matrix
        edited_df = st.data_editor(
            df_display,
            use_container_width=True,
            num_rows="dynamic",
            disabled=["Mã HS", "Họ và tên", "TBM HK"],
            key="grade_editor"
        )
        
        # Nút Lưu và Tính toán TBM
        col_btn1, col_btn2 = st.columns([2, 8])
        with col_btn1:
            if st.button("💾 Lưu thay đổi & Tính TBM", type="primary", use_container_width=True):
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                for _, row in edited_df.iterrows():
                    ma_hs = row["Mã HS"]
                    tx1, tx2, tx3, tx4 = row["TX1"], row["TX2"], row["TX3"], row["TX4"]
                    gk, ck, nx = row["Điểm GK"], row["Điểm CK"], row["Nhận xét"]
                    
                    # Thuật toán tính TBM y hệt mã nguồn gốc
                    tx_scores = [float(x) for x in [tx1, tx2, tx3, tx4] if pd.notna(x) and str(x).strip() != ""]
                    val_gk = float(gk) if pd.notna(gk) else None
                    val_ck = float(ck) if pd.notna(ck) else None
                    
                    tbm = None
                    if val_gk is not None and val_ck is not None:
                        total_sum = sum(tx_scores) + (val_gk * 2) + (val_ck * 3)
                        total_coef = len(tx_scores) + 2 + 3
                        tbm = round(total_sum / total_coef, 1)

                    cursor.execute("""
                        UPDATE grades SET kttx1=?, kttx2=?, kttx3=?, kttx4=?, ktgk=?, ktck=?, tb=?, comment_hk=? 
                        WHERE student_code=?
                    """, (tx1, tx2, tx3, tx4, gk, ck, tbm, nx, ma_hs))
                
                conn.commit()
                conn.close()
                st.success("✅ Đã lưu và cập nhật Điểm Trung Bình thành công!")
                st.rerun()
                
        with col_btn2:
            # Chức năng xuất file theo lớp
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, index=False, sheet_name=f"{selected_class}")
            st.download_button(label="📤 Xuất File Điểm Hiện Tại", data=output.getvalue(), file_name=f"Diem_{selected_class}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("💡 Chưa có dữ liệu học sinh. Vui lòng tải file SMAS (.xlsx) lên để đồng bộ.")