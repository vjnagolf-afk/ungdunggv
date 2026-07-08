import streamlit as st
import pandas as pd
import sqlite3
import re
import io

DB_PATH = "teacher_assistant.db"

# --- 1. KHỞI TẠO CẤU TRÚC BẢNG LƯU TRỮ TKB VĨNH VIỄN ---
def setup_tkb_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # Bảng lưu trữ thông tin tên các đợt TKB
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tkb_versions (
        version_id INTEGER PRIMARY KEY AUTOINCREMENT,
        version_name TEXT UNIQUE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    # Bảng lưu trữ chi tiết các ô dữ liệu của từng đợt
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tkb_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        version_name TEXT,
        thu TEXT,
        tiet TEXT,
        lop TEXT,
        noi_dung TEXT
    )
    """)
    conn.commit()
    conn.close()

def render_tkb_manager():
    setup_tkb_database()
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU THEO ĐỢT")
    
    # --- KHU VỰC 1: TẢI LÊN VÀ QUẢN LÝ CÁC ĐỢT TKB ---
    st.markdown("##### 📥 Nạp đợt Thời khóa biểu mới")
    col_file, col_name = st.columns([2, 1])
    with col_file:
        uploaded_tkb = st.file_uploader("Tải lên file TKB (.xlsx)", type=["xlsx"], key="tkb_uploader_main")
    with col_name:
        ten_dot_tkb = st.text_input("Đặt tên cho đợt TKB này:", placeholder="Ví dụ: Áp dụng từ 05/10")
        
    if uploaded_tkb and ten_dot_tkb:
        if st.button("🚀 Lưu và kích hoạt đợt TKB này", type="primary", use_container_width=True):
            with st.spinner("Hệ thống đang phân tích cấu trúc file và lưu vào cơ sở dữ liệu..."):
                try:
                    df_raw = pd.read_excel(uploaded_tkb, header=None)
                    header_idx = 4  
                    for idx, row in df_raw.iterrows():
                        row_str = [str(x).upper() for x in row.values if pd.notna(x)]
                        if "THỨ" in row_str and "TIẾT" in row_str:
                            header_idx = idx
                            break
                    
                    df = pd.read_excel(uploaded_tkb, header=header_idx)
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    if "THỨ" in df.columns:
                        df["THỨ"] = df["THỨ"].ffill()
                        df["THỨ"] = df["THỨ"].apply(lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x).strip())
                        
                    df = df.fillna("").astype(str).replace(["None", "nan", "NaN"], "")
                    class_columns = [c for c in df.columns if "Unnamed" not in c and c.upper() not in ["THỨ", "TIẾT", "STT", "CỘT 1", "CỘT 2"]]
                    
                    # Tiến hành lưu vào Database SQLite3
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()
                    
                    # Ghi nhận tên đợt
                    cursor.execute("INSERT OR REPLACE INTO tkb_versions (version_name) VALUES (?)", (ten_dot_tkb.strip(),))
                    
                    # Xóa dữ liệu cũ của đợt này nếu trùng tên để ghi đè
                    cursor.execute("DELETE FROM tkb_data WHERE version_name = ?", (ten_dot_tkb.strip(),))
                    
                    # Duyệt lưu từng ô dữ liệu
                    for _, row in df.iterrows():
                        thu = row.get("THỨ", "")
                        tiet = str(row.get("TIẾT", "")).strip()
                        for col_class in class_columns:
                            noi_dung_o = str(row[col_class]).strip()
                            if noi_dung_o:
                                cursor.execute("""
                                INSERT INTO tkb_data (version_name, thu, tiet, lop, noi_dung)
                                VALUES (?, ?, ?, ?, ?)
                                """, (ten_dot_tkb.strip(), thu, tiet, col_class, noi_dung_o))
                    conn.commit()
                    conn.close()
                    st.success(f"🎉 Đã lưu trữ thành công đợt TKB: **{ten_dot_tkb}** vào hệ thống!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi khi đọc file TKB: {e}")

    # --- KHU VỰC 2: MENU XỔ XUỐNG GỌI ĐỢT TKB VÀ NÚT XÓA ---
    st.markdown("---")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT version_name FROM tkb_versions ORDER BY created_at DESC")
    list_versions = [row[0] for row in cursor.fetchall()]
    conn.close()
    
    if not list_versions:
        st.info("ℹ️ Hệ thống hiện tại chưa có đợt dữ liệu TKB nào. Vui lòng đặt tên đợt và tải file Excel lên ở phía trên.")
        return

    st.markdown("##### 📁 Quản lý và Chọn đợt Thời khóa biểu tác nghiệp")
    col_select, col_delete = st.columns([3, 1])
    
    with col_select:
        # MENU XỔ XUỐNG CHỌN ĐỢT TKB THEO YÊU CẦU CỦA THẦY
        dot_duoc_chon = st.selectbox("Chọn đợt TKB muốn tra cứu:", list_versions)
        
    with col_delete:
        st.write("") # Cân lề dòng
        # CHỨC NĂNG XÓA ĐỢT TKB LỖI HOẶC ĐỢT CŨ KHÔNG DÙNG ĐẾN
        if st.button("🗑️ Xóa đợt TKB này", type="secondary", use_container_width=True):
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM tkb_versions WHERE version_name = ?", (dot_duoc_chon,))
            cursor.execute("DELETE FROM tkb_data WHERE version_name = ?", (dot_duoc_chon,))
            conn.commit()
            conn.close()
            st.warning(f"❌ Đã xóa hoàn toàn dữ liệu của đợt TKB: **{dot_duoc_chon}**!")
            st.rerun()

    # --- KHU VỰC 3: TRÍCH XUẤT VÀ DỰNG LẠI GIAO DIỆN THEO ĐỢT ĐÃ CHỌN ---
    conn = sqlite3.connect(DB_PATH)
    df_db = pd.read_sql_query("SELECT thu, tiet, lop, noi_dung FROM tkb_data WHERE version_name = ?", conn, params=[dot_duoc_chon])
    conn.close()
    
    if df_db.empty:
        st.warning("Đợt TKB này không có dữ liệu chi tiết.")
        return

    # Tự động nhận diện danh sách lớp học và giáo viên ĐỘNG dựa trên đợt được chọn (GV đi và đến tự cập nhật)
    class_columns = sorted(list(df_db["lop"].unique()))
    all_teachers = set()
    for nd in df_db["noi_dung"].values:
        if "-" in str(nd):
            parts = str(nd).split("-")
            if len(parts) >= 2: all_teachers.add(parts[1].strip())

    # Khởi tạo bảng TKB tổng thể từ dữ liệu SQLite3 để hiển thị ở Tab 1
    slots_list = ["1", "2", "3", "4", "5"]
    days_list = ["2", "3", "4", "5", "6", "7"]
    
    rows_master = []
    for d in days_list:
        for s in slots_list:
            row_dict = {"THỨ": d, "TIẾT": s}
            for c in class_columns:
                val_o = df_db[(df_db["thu"]==d) & (df_db["tiet"]==s) & (df_db["lop"]==c)]["noi_dung"].values
                row_dict[c] = val_o[0] if len(val_o) > 0 else ""
            rows_master.append(row_dict)
    df_master_view = pd.DataFrame(rows_master)

    # --- ĐIỀU HƯỚNG GIAO DIỆN TABS NGANG ---
    tab1, tab2 = st.tabs(["📊 Thời khóa biểu chung", "👤 TKB theo giáo viên"])
    
    with tab1:
        st.markdown(f"🗓️ **Bảng tổng thể đợt TKB: {dot_duoc_chon}**")
        st.dataframe(df_master_view, use_container_width=True, hide_index=True)
        
    with tab2:
        if not all_teachers:
            st.error("Không tìm thấy danh sách giáo viên bộ môn trong dữ liệu đợt này.")
        else:
            selected_teacher = st.selectbox("👤 Chọn tên Giáo viên bộ môn:", sorted(list(all_teachers)), key="gv_select_box")
            
            # Khung ma trận lưới TKB cầm tay 5 tiết x 6 thứ cho giáo viên
            days_grid = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"]
            matrix_data = {day: ["" for _ in slots_list] for day in days_grid}
            
            for _, row in df_master_view.iterrows():
                thu_raw = str(row.get("THỨ", "")).strip()
                tiet_raw = str(row.get("TIẾT", "")).strip()
                thu_clean = f"Thứ {thu_raw}"
                
                if thu_clean in matrix_data and tiet_raw in slots_list:
                    tiet_idx = slots_list.index(tiet_raw)
                    cell_lessons = []
                    
                    for col_class in class_columns:
                        cell_content = str(row[col_class]).strip()
                        if "-" in cell_content:
                            parts = cell_content.split("-")
                            if len(parts) >= 2 and parts[1].strip() == selected_teacher:
                                cell_lessons.append(f"{parts[0].strip()} - {col_class.split('(')[0].strip()}")
                                
                    if cell_lessons:
                        matrix_data[thu_clean][tiet_idx] = " / ".join(cell_lessons)
                        
            df_matrix = pd.DataFrame(matrix_data)
            df_matrix.insert(0, "TIẾT", slots_list)
            
            st.markdown(f"### 🗂️ Lịch dạy đợt **[{dot_duoc_chon}]** của Thầy/Cô: **{selected_teacher}**")
            st.dataframe(df_matrix, use_container_width=True, hide_index=True)
            
            # Xuất lịch dạy cá nhân của đợt tương ứng ra file Excel thu nhỏ
            output_personal = io.BytesIO()
            with pd.ExcelWriter(output_personal, engine='openpyxl') as writer:
                df_matrix.to_excel(writer, index=False, sheet_name=f"TKB_{selected_teacher}")
            st.download_button(
                label=f"📥 Tải lịch dạy đợt [{dot_duoc_chon}] về máy (.xlsx)",
                data=output_personal.getvalue(),
                file_name=f"TKB_{selected_teacher}_{dot_duoc_chon.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
