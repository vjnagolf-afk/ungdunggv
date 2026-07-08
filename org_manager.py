import streamlit as st
import sqlite3
import pandas as pd
import io
import re

DB_PATH = "teacher_assistant.db"

def setup_org_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS org_members (id INTEGER PRIMARY KEY AUTOINCREMENT, fullname TEXT UNIQUE, position TEXT, main_subject TEXT, email TEXT, phone TEXT, note TEXT)")
    cursor.execute("CREATE TABLE IF NOT EXISTS org_assignments (id INTEGER PRIMARY KEY AUTOINCREMENT, fullname TEXT UNIQUE, subject_class TEXT, homeroom TEXT, concurrent TEXT, total_periods TEXT DEFAULT '0')")
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS org_emulation_years (
        id INTEGER PRIMARY KEY AUTOINCREMENT, nam_hoc TEXT, fullname TEXT, dg_vien_chuc TEXT DEFAULT '', bd_hsg TEXT DEFAULT '',
        nckh TEXT DEFAULT '', stem TEXT DEFAULT '', sang_kien TEXT DEFAULT '', gvdg TEXT DEFAULT '', gvcng TEXT DEFAULT '',
        tdtt_nv TEXT DEFAULT '', hien_mau TEXT DEFAULT '', khac TEXT DEFAULT '', UNIQUE(nam_hoc, fullname)
    )
    """)
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
            cursor.execute("INSERT OR IGNORE INTO org_assignments (fullname, subject_class, homeroom, concurrent, total_periods) VALUES (?, '-', '-', '-', '0')", (row,))
    conn.commit()
    conn.close()

def clean_dataframe_nan(df):
    df = df.fillna("-")
    for col in df.columns:
        df[col] = df[col].apply(lambda x: "-" if str(x).strip().lower() in ["nan", "none", ""] else str(x).strip())
    return df
def render_org_section():
    setup_org_database()
    
    # --- SIDEBAR ĐĂNG NHẬP PHÂN QUYỀN MÃ PIN ADMIN ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
    
    # 💥 KHẮC PHỤC TRIỆT ĐỂ LỖI: Bổ sung tham số key tĩnh cố định ID cho widget radio khi gọi AI ngầm
    vai_tro = st.sidebar.radio(
        "Vai trò", 
        ["Giáo viên bộ môn", "Tổ trưởng chuyên môn (Admin)", "Ban giám hiệu"], 
        label_visibility="collapsed",
        key="vai_tro_sidebar_co_dinh" 
    )
    
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
    all_names = [row for row in conn.execute("SELECT fullname FROM org_members ORDER BY fullname ASC").fetchall()]
    conn.close()

    # ==================== THÈ 1: DANH SÁCH THÀNH VIÊN ====================
    with tab1:
        st.markdown("#### 👥 Danh sách thành viên tổ chuyên môn")
        
        excel_m = st.file_uploader("📥 Tải file Excel danh sách Giáo viên lên hệ thống:", type=["xlsx", "xls"], key="up_m")
        if excel_m:
            if st.button("🚀 Xác nhận nạp dữ liệu danh sách Giáo viên", type="secondary", use_container_width=True):
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
                    st.success("🎉 Đã cập nhật danh sách giáo viên từ file Excel thành công!")
                    st.rerun()
                except Exception as e: st.error(f"Lỗi đọc file: {e}")

        conn = sqlite3.connect(DB_PATH)
        df_members = pd.read_sql_query("SELECT fullname as [Họ và tên], position as [Chức vụ], main_subject as [Phân môn chính], email as [Email], phone as [Số điện thoại], note as [Ghi chú] FROM org_members", conn)
        conn.close()

        df_members = df_members.fillna("-")
        for col in df_members.columns:
            df_members[col] = df_members[col].apply(lambda x: "-" if str(x).strip().lower() in ["nan", "none", ""] else str(x).strip())

        if not df_members.empty:
            df_members.insert(0, "STT", range(1, len(df_members) + 1))
            
            col_cfg_members = {
                "STT": st.column_config.NumberColumn("STT", width=50, disabled=True),
                "Họ và tên": st.column_config.TextColumn("Họ và tên", width=200, disabled=True),
                "Chức vụ": st.column_config.TextColumn("Chức vụ", width=100, disabled=not is_admin),
                "Phân môn chính": st.column_config.TextColumn("Phân môn chính", width=180, disabled=not is_admin),
                "Email": st.column_config.TextColumn("Email", width=200, disabled=not is_admin),
                "Số điện thoại": st.column_config.TextColumn("Số điện thoại", width=120, disabled=not is_admin),
                "Ghi chú": st.column_config.TextColumn("Ghi chú", width=100, disabled=not is_admin)
            }
            
            with st.form("form_realtime_members", border=False):
                edited_m_df = st.data_editor(df_members, use_container_width=True, hide_index=True, column_config=col_cfg_members, key="stable_member_editor")
                save_m = st.form_submit_button("💾 XÁC NHẬN LƯU THAY ĐỔI DANH SÁCH GIÁO VIÊN", type="primary", use_container_width=True, disabled=not is_admin)
                
            if save_m and is_admin:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                for _, r in edited_m_df.iterrows():
                    cursor.execute("""
                    INSERT OR REPLACE INTO org_members (fullname, position, main_subject, email, phone, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """, (str(r["Họ và tên"]), str(r["Chức vụ"]), str(r["Phân môn chính"]), str(r["Email"]), str(r["Số điện thoại"]), str(r["Ghi chú"])))
                conn.commit()
                conn.close()
                st.success("🎉 Đã lưu trữ trực tiếp các thay đổi danh sách nhân sự!")
                st.rerun()
            
            out_m = io.BytesIO()
            with pd.ExcelWriter(out_m, engine='openpyxl') as w: edited_m_df.to_excel(w, index=False, sheet_name="Danh_Sach_GV")
            st.download_button(label="📥 Tải file Excel danh sách giáo viên về máy", data=out_m.getvalue(), file_name="Danh_Sach_Thanh_Vien_To.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.caption("Chưa có dữ liệu thành viên tổ.")
    # ==================== THÈ 2: KHO LƯU TRỮ BIÊN BẢN ====================
    with tab_kho:
        st.markdown("#### 🗄️ Thư viện lưu trữ biên bản sinh hoạt tổ")
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT id, meeting_date, session_number, present_members, absent_members, content_text, resolution FROM org_minutes ORDER BY meeting_date DESC")
        all_minutes = cursor.fetchall()
        conn.close()
        
        if not all_minutes:
            st.info("Hiện tại chưa có biên bản cuộc họp nào được lưu trữ trong phiên này.")
        else:
            for idx, row in enumerate(all_minutes):
                col_exp, col_del = st.columns([0.88, 0.12])
                with col_exp:
                    with st.expander(f"📝 {idx+1}. Biên bản số: {row[2]} - Ngày họp: {row[1]}"):
                        st.markdown(f"**🗓️ Ngày họp:** {row[1]}")
                        st.markdown(f"**👥 Thành phần tham gia:** {row[3]}")
                        st.markdown(f"**❌ Vắng mặt:** {row[4]}")
                        st.markdown("**📄 Diễn biến cuộc họp:**")
                        st.caption(row[5])
                        st.markdown("**🏅 Quyết nghị cuộc họp:**")
                        st.write(row[6])
                        
                        from org_manager import export_minutes_to_docx
                        word_data = export_minutes_to_docx(row[1], row[2], row[3], row[4], row[5], row[6])
                        st.download_button(
                            label="📥 Tải file Word (.docx) biên bản này",
                            data=word_data,
                            file_name=f"Bien_Ban_So_{row[2].replace('/', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_word_min_{row[0]}"
                        )
                with col_del:
                    st.write("") 
                    if st.button("🗑️ Xóa", key=f"del_min_{row[0]}", use_container_width=True):
                        conn = sqlite3.connect(DB_PATH)
                        conn.execute("DELETE FROM org_minutes WHERE id=?", (row[0],))
                        conn.commit()
                        conn.close()
                        st.success("Đã xóa biên bản!")
                        st.rerun()

# --- HÀM GIỮ NGUYÊN GỌI THEO HỆ THỐNG ---
def render_personal_plan():
    st.header("📅 KẾ HOẠCH GIÁO DỤC CÁ NHÂN")

# --- HÀM GIỮ NGUYÊN GỌI TỪ APP.PY ---
def render_personal_plan():
    st.header("📅 KẾ HOẠCH GIÁO DỤC CÁ NHÂN")

def render_personal_plan():
    st.header("📅 KẾ HOẠCH GIÁO DỤC CÁ NHÂN")
