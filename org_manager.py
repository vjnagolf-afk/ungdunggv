import streamlit as st
import sqlite3
import pandas as pd
import io
import re

DB_PATH = "teacher_assistant.db"

# --- 1. KHỞI TẠO CẤU TRÚC DỮ LIỆU THÀNH VIÊN VÀ THI ĐUA ĐA NĂM HỌC ---
def setup_org_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tạo bảng danh sách thành viên cố định của tổ (Thẻ 1)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS org_members (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fullname TEXT UNIQUE, position TEXT, main_subject TEXT, email TEXT, phone TEXT, note TEXT
    )
    """)
    
    # Tạo bảng phân công giảng dạy (Thẻ 2)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS org_assignments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fullname TEXT UNIQUE, subject_class TEXT, homeroom TEXT, concurrent TEXT, total_periods INTEGER DEFAULT 0
    )
    """)
    
    # Tạo bảng lưu thi đua thành tích phân tách theo năm học (Thẻ 3)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS org_emulation_years (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nam_hoc TEXT, fullname TEXT, dg_vien_chuc TEXT DEFAULT '', bd_hsg TEXT DEFAULT '',
        nckh TEXT DEFAULT '', stem TEXT DEFAULT '', sang_kien TEXT DEFAULT '', gvdg TEXT DEFAULT '',
        gvcng TEXT DEFAULT '', tdtt_nv TEXT DEFAULT '', hien_mau TEXT DEFAULT '', khac TEXT DEFAULT '',
        UNIQUE(nam_hoc, fullname)
    )
    """)
    
    # Nạp mẫu danh sách thành viên gốc trường THCS Nguyễn Chí Thanh nếu hệ thống trống
    cursor.execute("SELECT COUNT(*) FROM org_members")
    if cursor.fetchone() == 0:
        danh_sach_goc = [
            ("Lê Hồng Dưỡng", "Tổ trưởng", "KHTN (Vật lý) - CN", "vjnagolf@gmail.com", "0984331778", ""),
            ("Nguyễn Thị Huyền Trang", "Tổ viên", "KHTN (Vật lý) - CN", "nthtrangnct@gmail.com", "", ""),
            ("Lý Nguyễn Thu Nhi", "Tổ viên", "KHTN (Vật lý) - CN", "nthtrangnct@gmail.com", "", ""),
            ("Lê Hùng Cường", "Tổ viên", "KHTN (Vật lý) - CN", "nthtrangnct@gmail.com", "", ""),
            ("Khương Thị Thúy Vân", "Tổ viên", "KHTN (Sinh)", "nthtrangnct@gmail.com", "", ""),
            ("Trần Xuân Hạnh", "Tổ viên", "GDTC", "nthtrangnct@gmail.com", "", ""),
            ("Trương Vĩnh Văn", "Tổ viên", "KHTN (Sinh) - GDTC", "nthtrangnct@gmail.com", "", ""),
            ("Phạm Xuân Thọ", "Tổ viên", "KHTN (Sinh) - GDTC", "nthtrangnct@gmail.com", "", ""),
            ("Phạm Thùy Ngoan", "Tổ viên", "KHTN (Hóa)", "", "", ""),
            ("Huỳnh Thị Kim Lý", "Tổ viên", "KHTN", "", "", ""),
            ("Nguyễn Thanh Mai", "Tổ viên", "KHTN", "", "", ""),
            ("Phạm Thị Minh Anh", "Tổ viên", "KHTN", "", "", "")
        ]
        for row in danh_sach_goc:
            cursor.execute("INSERT OR IGNORE INTO org_members (fullname, position, main_subject, email, phone, note) VALUES (?, ?, ?, ?, ?, ?)", row)
            
    conn.commit()
    conn.close()
def render_org_section():
    setup_org_database()
    
    # --- SIDEBAR ĐĂNG NHẬP PHÂN QUYỀN MÃ PIN ADMIN ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
    vai_tro = st.sidebar.selectbox("Vai trò", ["Giáo viên bộ môn", "Tổ trưởng chuyên môn (Admin)", "Ban giám hiệu"], label_visibility="collapsed")
    
    is_admin = False
    if vai_tro == "Tổ trưởng chuyên môn (Admin)":
        st.sidebar.caption("Nhập mã pin quản lý Admin:")
        ma_pin = st.sidebar.text_input("Mã PIN", type="password", value="", label_visibility="collapsed")
        if ma_pin == "123456":
            st.sidebar.success("✅ Quyền Admin đã mở.")
            is_admin = True
        elif ma_pin != "":
            st.sidebar.error("❌ Mã PIN không chính xác.")

    st.subheader("📋 HỆ THỐNG QUẢN LÝ VÀ PHÂN CÔNG CHUYÊN MÔN GIẢNG DẠY")
    tab1, tab2, tab3 = st.tabs(["👥 Danh sách thành viên", "📊 Phân công giảng dạy", "🏅 Thành tích & Thi đua"])
    
    # Đọc nhanh danh sách họ tên giáo viên trong tổ để làm form chọn dữ liệu
    conn = sqlite3.connect(DB_PATH)
    all_names = [row[0] for row in conn.execute("SELECT fullname FROM org_members ORDER BY fullname ASC").fetchall()]
    conn.close()

    # ==================== THÈ 1: DANH SÁCH THÀNH VIÊN ====================
    with tab1:
        st.markdown("#### 👥 Danh sách thành viên tổ chuyên môn")
        
        # CHỨC NĂNG BỔ SUNG CHẤM ĐIỂM: TẢI FILE EXCEL LÊN (CHỈ DÀNH CHO ADMIN)
        if is_admin:
            col_f1, col_btn_f1 = st.columns([3, 1])
            with col_f1:
                excel_m = st.file_uploader("📥 Tải file Excel danh sách Giáo viên lên hệ thống:", type=["xlsx", "xls"], key="up_m")
            with col_btn_f1:
                st.write(""); st.write("")
                if st.button("🚀 Nạp Excel GV", use_container_width=True, type="secondary") and excel_m:
                    try:
                        df_up_m = pd.read_excel(excel_m)
                        df_up_m.columns = [str(c).strip() for c in df_up_m.columns]
                        conn = sqlite3.connect(DB_PATH)
                        cursor = conn.cursor()
                        for _, r in df_up_m.iterrows():
                            cursor.execute("""
                            INSERT OR REPLACE INTO org_members (fullname, position, main_subject, email, phone, note)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """, (str(r.get("Họ và tên", "")), str(r.get("Chức vụ", "Tổ viên")), str(r.get("Phân môn chính", "")), str(r.get("Email", "")), str(r.get("Số điện thoại", "")), str(r.get("Ghi chú", ""))))
                        conn.commit()
                        conn.close()
                        st.success("🎉 Đã đồng bộ danh sách giáo viên từ file Excel thành công!")
                        st.rerun()
                    except Exception as e: st.error(f"Lỗi đọc file: {e}")

        conn = sqlite3.connect(DB_PATH)
        df_members = pd.read_sql_query("SELECT fullname as [Họ và tên], position as [Chức vụ], main_subject as [Phân môn chính], email as [Email], phone as [Số điện thoại], note as [Ghi chú] FROM org_members", conn)
        conn.close()

        if not df_members.empty:
            df_members.insert(0, "STT", range(1, len(df_members) + 1))
            st.dataframe(df_members, use_container_width=True, hide_index=True)
            st.session_state["db_thanh_vien"] = df_members.to_dict(orient="records")
            
            # CHỨC NĂNG 2: XUẤT DANH SÁCH GIÁO VIÊN RA FILE EXCEL TẢI VỀ MÁY
            out_m = io.BytesIO()
            with pd.ExcelWriter(out_m, engine='openpyxl') as w: df_members.to_excel(w, index=False, sheet_name="Danh_Sach_GV")
            st.download_button(label="📥 Tải file Excel danh sách giáo viên về máy", data=out_m.getvalue(), file_name="Danh_Sach_Thanh_Vien_To.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.caption("Chưa có dữ liệu thành viên tổ.")
    # ==================== THÈ 2: PHÂN CÔNG GIẢNG DẠY ====================
    with tab2:
        st.markdown("#### 📊 Sơ đồ Phân công giảng dạy của Tổ")
        
        # CHỨC NĂNG BỔ SUNG CHẤM ĐIỂM: TẢI FILE EXCEL LÊN (CHỈ DÀNH CHO ADMIN)
        if is_admin:
            col_f2, col_btn_f2 = st.columns([3, 1])
            with col_f2:
                excel_a = st.file_uploader("📥 Tải file Excel phân công giảng dạy lên hệ thống:", type=["xlsx", "xls"], key="up_a")
            with col_btn_f2:
                st.write(""); st.write("")
                if st.button("🚀 Nạp Excel PC", use_container_width=True, type="secondary") and excel_a:
                    try:
                        df_up_a = pd.read_excel(excel_a)
                        df_up_a.columns = [str(c).strip() for c in df_up_a.columns]
                        conn = sqlite3.connect(DB_PATH)
                        cursor = conn.cursor()
                        for _, r in df_up_a.iterrows():
                            # Ép kiểu số tiết về số nguyên an toàn
                            t_periods = int(float(r.get("Tổng số tiết", 0))) if str(r.get("Tổng số tiết", 0)).replace('.','').isdigit() else 0
                            cursor.execute("""
                            INSERT OR REPLACE INTO org_assignments (fullname, subject_class, homeroom, concurrent, total_periods)
                            VALUES (?, ?, ?, ?, ?)
                            """, (str(r.get("Họ tên GV", "")), str(r.get("Môn-Lớp", "")), str(r.get("Chủ nhiệm", "")), str(r.get("Kiêm nhiệm", "")), t_periods))
                        conn.commit()
                        conn.close()
                        st.success("🎉 Đã lưu sơ đồ phân công giảng dạy từ file Excel thành công!")
                        st.rerun()
                    except Exception as e: st.error(f"Lỗi đọc file phân công: {e}")

        # Đọc dữ liệu từ hệ thống để dựng bảng 7 cột quy chuẩn
        conn = sqlite3.connect(DB_PATH)
        df_assign = pd.read_sql_query("""
            SELECT m.fullname as [Họ tên GV], 
                   ifnull(a.subject_class, '') as [Môn-Lớp], 
                   ifnull(a.homeroom, '') as [Chủ nhiệm], 
                   ifnull(a.concurrent, '') as [Kiêm nhiệm], 
                   ifnull(a.total_periods, 0) as [Tổng số tiết]
            FROM org_members m LEFT JOIN org_assignments a ON m.fullname = a.fullname
        """, conn)
        conn.close()
        
        if not df_assign.empty:
            df_assign.insert(0, "STT", range(1, len(df_assign) + 1))
            st.dataframe(df_assign, use_container_width=True, hide_index=True)
            
            # CHỨC NĂNG 2: XUẤT FILE EXCEL PHÂN CÔNG GIẢNG DẠY TẢI VỀ MÁY
            out_a = io.BytesIO()
            with pd.ExcelWriter(out_a, engine='openpyxl') as w: df_assign.to_excel(w, index=False, sheet_name="Phan_Cong")
            st.download_button(label="📥 Tải file Excel phân công giảng dạy về máy", data=out_a.getvalue(), file_name="So_Do_Phan_Cong_Giang_Day.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    # ==================== THÈ 3: THÀNH TÍCH VÀ THI ĐUA ====================
    with tab3:
        st.markdown("#### 🏆 Sổ theo dõi Thành tích & Thi đua qua các năm học")
        
        years_options = [f"{y} - {y+1}" for y in range(2020, 2036)]
        selected_year = st.selectbox("📅 CHỌN NĂM HỌC TRA CỨU THI ĐUA:", years_options, index=5) # Mặc định năm học 2025-2026
        
        # CHỨC NĂNG BỔ SUNG CHẤM ĐIỂM: TẢI FILE EXCEL LÊN (CHỈ DÀNH CHO ADMIN)
        if is_admin:
            col_f3, col_btn_f3 = st.columns([3, 1])
            with col_f3:
                excel_e = st.file_uploader(f"📥 Tải file Excel thi đua năm học {selected_year} lên hệ thống:", type=["xlsx", "xls"], key="up_e")
            with col_btn_f3:
                st.write(""); st.write("")
                if st.button("🚀 Nạp Excel TĐ", use_container_width=True, type="secondary") and excel_e:
                    try:
                        df_up_e = pd.read_excel(excel_e)
                        df_up_e.columns = [str(c).strip() for c in df_up_e.columns]
                        conn = sqlite3.connect(DB_PATH)
                        cursor = conn.cursor()
                        for _, r in df_up_e.iterrows():
                            cursor.execute("""
                            INSERT OR REPLACE INTO org_emulation_years 
                            (nam_hoc, fullname, dg_vien_chuc, bd_hsg, nckh, stem, sang_kien, gvdg, gvcng, tdtt_nv, hien_mau, khac)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (selected_year, str(r.get("Họ và tên", "")), str(r.get("Đánh giá viên chức", "")), str(r.get("BD HSG", "")), str(r.get("NCKH", "")), str(r.get("STEM", "")), str(r.get("Sáng kiến", "")), str(r.get("GVDG", "")), str(r.get("GVCNG", "")), str(r.get("TDTT-NV", "")), str(r.get("Hiển máu nhân đạo", "")), str(r.get("Khác", ""))))
                        conn.commit()
                        conn.close()
                        st.success(f"🎉 Đã nạp thành công sổ điểm thi đua từ file Excel cho năm học {selected_year}!")
                        st.rerun()
                    except Exception as e: st.error(f"Lỗi đọc file thi đua: {e}")

        # Đọc dữ liệu từ SQLite3 dựng lưới ma trận 12 cột chuẩn xác theo ảnh mẫu của thầy
        conn = sqlite3.connect(DB_PATH)
        df_emulation = pd.read_sql_query("""
            SELECT m.fullname as [Họ và tên],
                   ifnull(e.dg_vien_chuc, '') as [Đánh giá viên chức],
                   ifnull(e.bd_hsg, '') as [BD HSG],
                   ifnull(e.nckh, '') as [NCKH],
                   ifnull(e.stem, '') as [STEM],
                   ifnull(e.sang_kien, '') as [Sáng kiến],
                   ifnull(e.gvdg, '') as [GVDG],
                   ifnull(e.gvcng, '') as [GVCNG],
                   ifnull(e.tdtt_nv, '') as [TDTT-NV],
                   ifnull(e.hien_mau, '') as [Hiển máu nhân đạo],
                   ifnull(e.khac, '') as [Khác]
            FROM org_members m LEFT JOIN org_emulation_years e ON m.fullname = e.fullname AND e.nam_hoc = ?
        """, conn, params=[selected_year])
        conn.close()

        df_emulation.insert(0, "Năm học", selected_year)
        df_emulation.insert(0, "STT", range(1, len(df_emulation) + 1))
        
        st.markdown(f"📊 **Bảng chi tiết chỉ tiêu thi đua danh hiệu năm học: {selected_year}**")
        st.dataframe(df_emulation, use_container_width=True, hide_index=True)
        
        # CHỨC NĂNG 2: XUẤT FILE EXCEL THÀNH TÍCH THI ĐUA CỦA NĂM ĐANG TRA CỨU
        out_e = io.BytesIO()
        with pd.ExcelWriter(out_e, engine='openpyxl') as w: df_emulation.to_excel(w, index=False, sheet_name=f"Thi_Dua")
        st.download_button(label=f"📥 Kết xuất Báo cáo Thi đua năm học {selected_year} (.xlsx)", data=out_e.getvalue(), file_name=f"Bao_Cao_Thi_Dua_Nam_Hoc_{selected_year.replace(' ', '')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- CÁC HÀM QUẢN LÝ VỆ TINH PHỤ TRỢ (Giữ nguyên cấu trúc gọi từ app.py) ---
def render_meeting_minutes():
    st.header("📝 BIÊN BẢN SINH HOẠT TỔ")
    st.write("Giao diện nhập biên bản họp chuyên môn định kỳ.")

def render_personal_plan():
    st.header("📅 KẾ HOẠCH GIÁO DỤC CÁ NHÂN")
    st.write("Giao diện lập kế hoạch cá nhân (Phụ lục III - 5512).")
