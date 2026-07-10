# app.py
import streamlit as st
import pandas as pd
from google import genai
from google.genai import errors  

# --- 1. PHÂN LUỒNG IMPORT CÁC MODULE ĐỘC LẬP ---
from exam_designer import render_exam_designer_section # 🚀 ĐÃ MỞ KHÓA IMPORT TẠI ĐÂY
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_manager import render_khbd_section  
from danh_gia_manager import render_assessment_section

from org_manager import render_org_section
from bien_ban_manager import render_meeting_minutes
from ke_hoach_ca_nhan_manager import render_personal_plan 
from stem_manager import render_stem_section
from chu_nhiem_manager import render_chu_nhiem_section 

# --- 2. CẤU HÌNH ĐỌC API KEY TỰ ĐỘNG TỪ TRONG SECRETS ---
API_KEY_HE_THONG = st.secrets.get("GEMINI_API_KEY", "")

# app.py (Cập nhật Khối gọi API đồng bộ danh mục mô hình mới)
from google import genai
from google.genai import errors
import streamlit as st

API_KEY_HE_THONG = st.secrets.get("GEMINI_API_KEY", "")

def run_ai_prompt_safe(prompt_text, preferred_model="3.5 Flash"):
    """
    Hàm gọi API Gemini thế hệ mới tích hợp cơ chế Fallback (Tự động chuyển đổi mô hình)
    đúng theo các tùy chọn: 3.1 Pro, 3.5 Flash, 3.1 Flash-Lite, Tư duy mở rộng.
    """
    api_key = API_KEY_HE_THONG
    if not api_key:
        return "⚠️ Hệ thống chưa được cấu hình API Key trong mục Secrets. Vui lòng liên hệ Admin.", "error"
    
    # 🌟 ĐỒNG BỘ MÃ MÔ HÌNH CHUẨN KỸ THUẬT THEO CẤU TRÚC PHÂN CẤP DỰ PHÒNG
    model_pool = {
        "3.1 Pro": ["gemini-1.5-pro", "gemini-2.5-flash", "gemini-2.5-flash-8b"],
        "3.5 Flash": ["gemini-2.5-flash", "gemini-2.5-flash-8b"],
        "3.1 Flash-Lite": ["gemini-2.5-flash-8b"],
        "Tư duy mở rộng": ["gemini-2.5-pro", "gemini-1.5-pro", "gemini-2.5-flash"]
    }
    
    # Lấy danh sách mô hình sẽ lần lượt thử nghiệm dựa trên lựa chọn của giáo viên
    models_to_try = model_pool.get(preferred_model, ["gemini-2.5-flash"])
    
    last_error_message = ""
    client = genai.Client(api_key=api_key)
    
    # VÒNG LẶP TỰ ĐỘNG CHUYỂN ĐỔI THÔNG MINH KHI GẶP LỖI HẠN MỨC (QUOTA)
    for model_name in models_to_try:
        try:
            # Nếu người dùng chọn chế độ "Tư duy mở rộng" (Thinking), bổ sung tham số bật tính năng ẩn
            config_params = {}
            if preferred_model == "Tư duy mở rộng" and "pro" in model_name:
                config_params["thinking_config"] = {"thinking_budget": 1024} # Kích hoạt chế độ giải toán chuyên sâu
            
            response = client.models.generate_content(
                model=model_name,
                contents=prompt_text,
                config=config_params if config_params else None
            )
            return response.text, model_name
            
        except errors.APIError as error:  
            last_error_message = f"Mô hình {model_name} lỗi hạn mức hoặc nghẽn mạng. Đang lùi về dòng máy dự phòng..."
            st.toast(last_error_message, icon="⚠️")
            continue  
            
        except Exception as e:
            last_error_message = f"Mô hình {model_name} gặp sự cố cấu trúc: {str(e)}"
            continue
            
    return f"Lỗi quá tải hệ thống trên diện rộng (Tất cả mô hình dự phòng đều cạn hạn mức). Lỗi cuối cùng: {last_error_message}", "error"
# --- 3. KHỞI TẠO BỘ NHỚ TẠM ĐỒNG BỘ ---
if "db_thanh_vien" not in st.session_state: 
    st.session_state["db_thanh_vien"] = []
if "db_phan_cong_hien_tai" not in st.session_state: 
    st.session_state["db_phan_cong_hien_tai"] = []

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

# --- 4. GIAO DIỆN PHẦN CHÍNH (CANH GIỮA, ĐỔI MÀU) ---
st.markdown("<h1 style='text-align: center; color: darkred; font-weight: bold;'>🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #0056b3; font-weight: bold; font-size: 16px;'>Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk</p>", unsafe_allow_html=True)
st.markdown("---")

## ==================================================================================
# --- THANH ĐIỀU HƯỚNG TỔNG (ĐẨY LÊN TRÊN CÙNG SIDEBAR) ---
## ==================================================================================
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_root_key_2026_v9"
)
# ==================================================================================
# --- 5. XỬ LÝ ĐIỀU HƯỚNG & GỌI PHÂN HỆ TÁC NGHIỆP ---
# ==================================================================================
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB","6. Thiết kế bài dạy STEM","7. Kế hoạch công tác chủ nhiệm lớp"], label_visibility="collapsed", key="menu_gv_selectbox_v9")
    
    st.sidebar.success("🔑 Đã kết nối API Key hệ thống từ Manage App.")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "2. Thiết kế Đề KT": 
        render_exam_designer_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Đánh giá HS": 
        render_assessment_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()
    elif menu == "6. Thiết kế bài dạy STEM":
        render_stem_section()
    elif menu == "7. Kế hoạch công tác chủ nhiệm lớp":
        render_chu_nhiem_section(lambda p: run_ai_prompt_safe(p))

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
                except Exception as e:
                    st.error(f"❌ Không thể nạp dữ liệu mẫu: {str(e)}")
        else:
            st.subheader("📋 Danh sách phân công hiện tại:")
            st.dataframe(df_tv, use_container_width=True)
