import streamlit as st
import pandas as pd
from google import genai

# --- 1. PHÂN LUỒNG IMPORT CÁC MODULE ĐỘC LẬP ---
from exam_designer import render_exam_designer_section
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_manager import render_khbd_section  
from danh_gia_manager import render_assessment_section

from org_manager import render_org_section
from bien_ban_manager import render_meeting_minutes
from ke_hoach_ca_nhan_manager import render_personal_plan 
from stem_manager import render_stem_section
# --- 2. CẤU HÌNH ĐỌC API KEY TỰ ĐỘNG TỪ TRONG SECRETS ---
API_KEY_HE_THONG = st.secrets.get("GEMINI_API_KEY", "")

def run_ai_prompt_safe(prompt_text):
    api_key = API_KEY_HE_THONG
    if not api_key:
        return "⚠️ Hệ thống chưa được cấu hình API Key trong mục Secrets. Vui lòng liên hệ Admin.", "error"
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt_text,
        )
        return response.text, "gemini-2.5-flash"
    except Exception as main_error:
        error_msg = str(main_error)
        if "503" in error_msg or "UNAVAILABLE" in error_msg or "high demand" in error_msg:
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model='gemini-1.5-flash',
                    contents=prompt_text,
                )
                return response.text, "gemini-1.5-flash (Kênh Dự phòng)"
            except Exception as backup_error:
                return f"Lỗi quá tải hệ thống trên diện rộng: {str(backup_error)}", "error"
        else:
            return f"Lỗi kết nối AI: {error_msg}", "error"

# --- 3. KHỞI TẠO BỘ NHỚ TẠM ĐỒNG BỘ ---
if "db_thanh_vien" not in st.session_state: 
    st.session_state["db_thanh_vien"] = []
if "db_phan_cong_hien_tai" not in st.session_state: 
    st.session_state["db_phan_cong_hien_tai"] = []

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

st.title("🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")
st.caption("Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk")
st.markdown("---")

## ==================================================================================
# --- THANH ĐIỀU HƯỚNG TỔNG ---
## ==================================================================================
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_root_key_2026_v9"
)

# --- 5. XỬ LÝ ĐIỀU HƯỚNG ---
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB","6. Thiết kế bài dạy STEM"], label_visibility="collapsed", key="menu_gv_selectbox_v9")
    
    st.sidebar.success("🔑 Đã kết nối API Key hệ thống từ Manage App.")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "2. Thiết kế Đề KT": 
        render_exam_designer_section("", lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Đánh giá HS": 
        render_assessment_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()
    elif menu == "6. Thiết kế bài dạy STEM":
        render_stem_section()
else:  # Phân hệ Quản lý tổ chuyên môn
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox("Nội dung quản lý", ["1. Quản lý & Phân công chuyên môn", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"], label_visibility="collapsed", key="menu_ql_selectbox_v9")
    
    if menu == "1. Quản lý & Phân công chuyên môn": 
        render_org_section()
    elif menu == "2. Biên bản sinh hoạt": 
        render_meeting_minutes(lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Kế hoạch cá nhân": 
        render_personal_plan(lambda p: run_ai_prompt_safe(p))
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        
        # --- 🌟 ĐỌC DỮ LIỆU TRỰC TIẾP TỪ SQLITE CỦA NHÁNH 1 ---
        import sqlite3
        import os
        
        DB_PATH = "teacher_assistant.db"
        df_tv = pd.DataFrame()
        
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                query = """
                SELECT 
                    m.fullname as [Họ và tên],
                    m.main_subject as [Phân môn chính],
                    a.total_periods as [Số tiết/Tuần]
                FROM org_members m
                LEFT JOIN org_assignments a ON m.fullname = a.fullname
                """
                df_tv = pd.read_sql_query(query, conn)
                conn.close()
            except Exception as e:
                st.error(f"⚠️ Lỗi kết nối cơ sở dữ liệu nội bộ: {str(e)}")
        
        thuc_te_co_du_lieu = False
        if not df_tv.empty:
            if len(df_tv) > 0 and not (df_tv["Phân môn chính"] == "").all():
                thuc_te_co_du_lieu = True

        # --- LUỒNG 1: NẾU CƠ SỞ DỮ LIỆU TRỐNG (CHẾ ĐỘ THỬ NGHIỆM ĐI THI) ---
        if not thuc_te_co_du_lieu:
            st.warning("ℹ️ Hiện tại chưa có dữ liệu giáo viên nào được nhập từ phân hệ '1. Quản lý & Phân công chuyên môn'.")
            
            if st.button("💡 Nạp nhanh dữ liệu mẫu để thử nghiệm biểu đồ", type="primary", use_container_width=True):
                try:
                    conn = sqlite3.connect(DB_PATH)
                    conn.execute("CREATE TABLE IF NOT EXISTS org_members (id INTEGER PRIMARY KEY AUTOINCREMENT, fullname TEXT UNIQUE, position TEXT, main_subject TEXT, email TEXT, phone TEXT, note TEXT)")
                    conn.execute("CREATE TABLE IF NOT EXISTS org_assignments (id INTEGER PRIMARY KEY AUTOINCREMENT, fullname TEXT UNIQUE, subject_class TEXT, homeroom TEXT, concurrent TEXT, total_periods TEXT DEFAULT '0')")
                    
                    danh_sach_demo = [
                        ("Lê Hồng Dưỡng", "Khoa học tự nhiên (Phân môn Vật lí)", "14"),
                        ("Nguyễn Thị Huyền Trang", "Khoa học tự nhiên (Phân môn Vật lí)", "16"),
                        ("Khương Thị Thúy Vân", "Khoa học tự nhiên (Phân môn Sinh học)", "12"),
                        ("Phạm Thùy Ngoan", "Khoa học tự nhiên (Phân môn Hóa học)", "15"),
                        ("Trần Xuân Hạnh", "Giáo dục thể chất", "14")
                    ]
                    for name, subj, periods in danh_sach_demo:
                        conn.execute("INSERT OR REPLACE INTO org_members (fullname, position, main_subject) VALUES (?, 'GV', ?)", (name, subj))
                        conn.execute("INSERT OR REPLACE INTO org_assignments (fullname, total_periods) VALUES (?, ?)", (name, periods))
                    conn.commit()
                    conn.close()
                    st.success("🎉 Đã nạp dữ liệu thử nghiệm trực tiếp vào SQLite! Đang dựng sơ đồ...")
                    st.rerun()
                except Exception as demo_err:
                    st.error(f"Không thể nạp dữ liệu mẫu: {demo_err}")
                
        # --- LUỒNG 2: ĐỒNG BỘ DỮ LIỆU THỰC TẾ THỜI GIAN THỰC TỪ SQLITE ---
        else:
            df_tv["Phân môn chính"] = df_tv["Phân môn chính"].fillna("Chưa phân môn").replace("", "Chưa phân môn")
            df_tv["Họ và tên"] = df_tv["Họ và tên"].fillna("Giáo viên ẩn danh").replace("", "Giáo viên ẩn danh")
            
            # Ép kiểu dữ liệu số tiết sang số để tính toán
            df_tv["Số tiết/Tuần"] = pd.to_numeric(df_tv["Số tiết/Tuần"], errors='coerce').fillna(0).astype(int)
            
            # Nếu số tiết bằng 0 hoặc chưa phân công, gán giá trị định mức mặc định để vẽ biểu đồ không bị bằng 0
            df_tv.loc[df_tv["Số tiết/Tuần"] == 0, "Số tiết/Tuần"] = 14
            
            # Chỉ số tổng quan
            st.markdown("### 📌 Chỉ số tổng quan tổ chuyên môn (Dữ liệu thực tế từ SQLite)")
            m_col1, m_col2, m_col3 = st.columns(3)
            
            tong_so_gv = len(df_tv)
            m_col1.metric(label="👥 Tổng số Giáo viên trong tổ", value=f"{tong_so_gv} Thầy/Cô")
            
            so_phan_mon = df_tv["Phân môn chính"].nunique()
            m_col2.metric(label="📚 Số lượng Môn học/Phân môn", value=f"{so_phan_mon} Nhóm")
            
            tong_tiet = int(df_tv["Số tiết/Tuần"].sum())
            m_col3.metric(label="⏱️ Tổng số tiết định mức / tuần", value=f"{tong_tiet} Tiết")
                
            st.markdown("---")
            
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("##### 📈 Số lượng giáo viên theo Phân môn")
                counts = df_tv["Phân môn chính"].value_counts()
                st.bar_chart(counts, color="#1E3A8A")
                
            with chart_col2:
                st.markdown("##### ⏳ Định mức Tiết dạy/Tuần của từng Giáo viên")
                df_chart_tiet = df_tv.set_index("Họ và tên")[["Số tiết/Tuần"]]
                st.bar_chart(df_chart_tiet, color="#EF4444")
                    
            st.markdown("---")
            st.markdown("### 🗂️ Danh sách trích xuất dữ liệu chi tiết")
            st.dataframe(df_tv, use_container_width=True, hide_index=True)


