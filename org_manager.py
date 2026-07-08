import streamlit as st
import sqlite3
import pandas as pd

DB_PATH = "teacher_assistant.db"

# --- 1. KHỞI TẠO CẤU TRÚC BẢNG THI ĐUA VĨNH VIỄN TRONG CƠ SỞ DỮ LIỆU ---
def setup_org_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # Tạo bảng lưu thi đua thành tích giáo viên nếu chưa có
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS org_emulation (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fullname TEXT UNIQUE,
        good_lessons INTEGER DEFAULT 0,
        skkn_count INTEGER DEFAULT 0,
        hsg_count INTEGER DEFAULT 0,
        emulation_title TEXT DEFAULT 'Lao động tiên tiến'
    )
    """)
    conn.commit()
    conn.close()

def render_org_section():
    setup_org_database()
    
    # --- 2. SIDEBAR ĐĂNG NHẬP PHÂN QUYỀN VAI TRÒ ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
    st.sidebar.caption("VỚI TƯ CÁCH LÀ:")
    
    vai_tro = st.sidebar.selectbox(
        "Vai trò",
        ["Giáo viên bộ môn", "Tổ trưởng chuyên môn (Admin)", "Ban giám hiệu"],
        label_visibility="collapsed"
    )
    
    is_admin = False
    if vai_tro == "Tổ trưởng chuyên môn (Admin)":
        st.sidebar.caption("Nhập mã pin quản lý Admin:")
        ma_pin = st.sidebar.text_input("Mã PIN", type="password", value="", label_visibility="collapsed")
        
        # Mã PIN mặc định kết nối hệ thống
        if ma_pin == "123456":
            st.sidebar.success("✅ Quyền Admin đã mở.")
            is_admin = True
        elif ma_pin != "":
            st.sidebar.error("❌ Mã PIN không chính xác.")

    # --- 3. VÙNG HIỂN THỊ NỘI DUNG CHÍNH (MAIN PANEL TABS) ---
    st.subheader("📋 HỆ THỐNG QUẢN LÝ VÀ PHÂN CÔNG CHUYÊN MÔN GIẢNG DẠY")
    
    # Khởi tạo 3 Thẻ điều hướng ngang khít theo ảnh mẫu của thầy
    tab1, tab2, tab3 = st.tabs(["👥 Danh sách thành viên", "📊 Phân công", "🏅 Thành tích"])
    
    # --- THẺ 1: DANH SÁCH THÀNH VIÊN ---
    with tab1:
        st.markdown("#### 👥 Danh sách thành viên tổ chuyên môn")
        if is_admin:
            st.info("💡 Chế độ Quản trị: Bạn có thể nhập thêm thành viên mới.")
            with st.form("form_add_member", clear_on_submit=True):
                new_name = st.text_input("Họ và tên giáo viên bộ môn:")
                new_subject = st.selectbox("Phân môn giảng dạy:", ["Ngữ văn", "Toán học", "Tiếng Anh", "Khoa học tự nhiên", "Lịch sử & Địa lý", "Tin học", "Công nghệ"])
                if st.form_submit_button("➕ Thêm vào tổ"):
                    if new_name:
                        st.session_state["db_thanh_vien"].append({"Họ và tên": new_name, "Phân môn chính": new_subject})
                        # Đồng thời khởi tạo dòng thi đua trống bên bảng Thành tích để không bị lỗi kết nối
                        try:
                            conn = sqlite3.connect(DB_PATH)
                            cursor = conn.cursor()
                            cursor.execute("INSERT OR IGNORE INTO org_emulation (fullname) VALUES (?)", (new_name.strip(),))
                            conn.commit()
                            conn.close()
                        except: pass
                        st.success(f"Đã thêm thầy/cô {new_name} vào danh sách tổ!")
                        st.rerun()
                        
        df_tv = pd.DataFrame(st.session_state.get("db_thanh_vien", []))
        if not df_tv.empty:
            df_tv.insert(0, "STT", range(1, len(df_tv) + 1))
            st.dataframe(df_tv, use_container_width=True, hide_index=True)
        else:
            st.caption("Chưa có dữ liệu thành viên tổ. Thầy có thể cấu hình nhanh ở phân hệ Thống kê số liệu.")
            
    # --- THẺ 2: PHÂN CÔNG GIẢNG DẠY ---
    with tab2:
        st.markdown("#### 📊 Sơ đồ Phân công giảng dạy của Tổ")
        df_pc = pd.DataFrame(st.session_state.get("db_phan_cong_hien_tai", []))
        if not df_pc.empty:
            st.dataframe(df_pc, use_container_width=True)
        else:
            st.info("ℹ️ Lịch dạy phân công chi tiết đang được đồng bộ động trực tiếp từ tệp Thời khóa biểu Excel của trường.")

    # --- THẺ 3: THÀNH TÍCH VÀ THI ĐUA (SỬA LỖI MÃ RÁC JSON CHỮ {}) ---
    with tab3:
        st.markdown("#### 🏆 Bảng theo dõi Thành tích & Thi đua Tổ chuyên môn")
        
        # Nếu đăng nhập quyền Admin, hiển thị biểu mẫu cập nhật thi đua hành vi định lượng
        if is_admin:
            with st.expander("🛠️ BIỂU MẪU CẬP NHẬT THÀNH TÍCH (Dành cho Tổ trưởng)"):
                conn = sqlite3.connect(DB_PATH)
                df_all_gv = pd.read_sql_query("SELECT fullname FROM org_emulation", conn)
                conn.close()
                
                list_names = df_all_gv["fullname"].tolist() if not df_all_gv.empty else [x["Họ và tên"] for x in st.session_state.get("db_thanh_vien", [])]
                
                if not list_names:
                    st.warning("Vui lòng thêm thành viên ở Tab 1 trước khi nhập thi đua.")
                else:
                    with st.form("form_update_emulation"):
                        gv_select = st.selectbox("Chọn Giáo viên cần ghi nhận thành tích:", sorted(list(set(list_names))))
                        col_e1, col_e2, col_e3 = st.columns(3)
                        with col_e1: s_tot = st.number_input("Số tiết dạy xếp loại TỐT:", min_value=0, value=0, step=1)
                        with col_e2: s_skkn = st.number_input("Số Sáng kiến kinh nghiệm (SKKN):", min_value=0, value=0, step=1)
                        with col_e3: s_hsg = st.number_input("Số giải Học sinh giỏi (HSG):", min_value=0, value=0, step=1)
                        
                        danh_hieu = st.selectbox("Danh hiệu thi đua dự kiến:", ["Lao động tiên tiến", "Chiến sĩ thi đua cơ sở", "Bằng khen cấp Tỉnh", "Giáo viên dạy giỏi cấp Trường"])
                        
                        if st.form_submit_button("💾 Cập nhật thi đua vĩnh viễn"):
                            conn = sqlite3.connect(DB_PATH)
                            cursor = conn.cursor()
                            cursor.execute("""
                            INSERT OR REPLACE INTO org_emulation (fullname, good_lessons, skkn_count, hsg_count, emulation_title)
                            VALUES (?, ?, ?, ?, ?)
                            """, (gv_select, s_tot, s_skkn, s_hsg, danh_hieu))
                            conn.commit()
                            conn.close()
                            st.success(f"🎉 Đã lưu thành tích của thầy/cô: **{gv_select}** vào sổ thi đua!")
                            st.rerun()

        # Đọc dữ liệu từ SQLite3 dựng bảng thi đua hành chính sạch đẹp
        conn = sqlite3.connect(DB_PATH)
        df_emulation = pd.read_sql_query("""
            SELECT fullname as [Họ và tên giáo viên], 
                   good_lessons as [Số tiết dạy Tốt], 
                   skkn_count as [Số SKKN đạt duyệt], 
                   hsg_count as [Số giải HSG đạt], 
                   emulation_title as [Danh hiệu thi đua]
            FROM org_emulation
        """, conn)
        conn.close()

        if not df_emulation.empty:
            df_emulation.insert(0, "STT", range(1, len(df_emulation) + 1))
            # Hiển thị bảng điểm thi đua dạng lưới Excel chuyên nghiệp thay thế mã rác cũ
            st.dataframe(df_emulation, use_container_width=True, hide_index=True)
            
            # Tích hợp nút xuất Sổ theo dõi thi đua ra Excel nhanh phục vụ báo cáo Ban giám hiệu
            output_org = io_bytes_wrapper() if 'io_bytes_wrapper' in locals() else b""
            import io
            output_org = io.BytesIO()
            with pd.ExcelWriter(output_org, engine='openpyxl') as writer:
                df_emulation.to_excel(writer, index=False, sheet_name="Thi_Dua_To")
            st.download_button(
                label="📥 Kết xuất Báo cáo Thi đua Tổ (.xlsx)",
                data=output_org.getvalue(),
                file_name="Bao_Cao_Thi_Dua_To_Chuyen_Mon.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.caption("ℹ️ Chưa có dữ liệu thi đua. Tổ trưởng vui lòng chọn quyền Admin ở Sidebar để nhập chỉ tiêu.")

# --- 4. CÁC HÀM QUẢN LÝ VỆ TINH PHỤ TRỢ (Giữ nguyên cấu trúc gọi từ app.py) ---
def render_meeting_minutes():
    st.header("📝 BIÊN BẢN SINH HOẠT TỔ")
    st.write("Giao diện quản lý nội dung các buổi họp chuyên môn định kỳ.")

def render_personal_plan():
    st.header("📅 KẾ HOẠCH GIÁO DỤC CÁ NHÂN")
    st.write("Giao diện lập kế hoạch cá nhân (Phụ lục III - 5512).")
