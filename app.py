import streamlit as st
import pandas as pd
from google import genai

# --- 1. IMPORT CÁC MODULE (Đã đồng bộ đầy đủ các hàm từ org_manager.py) ---
from exam_designer import render_exam_designer_section
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_manager import render_khbd_section
from org_manager import render_org_section, render_meeting_minutes, render_personal_plan

# --- 2. CẤU HÌNH AI (Chuẩn SDK Google GenAI mới cho cuộc thi năm 2026) ---
def run_ai_prompt_safe(prompt_text, api_key):
    if not api_key:
        return "Vui lòng nhập API Key tại Sidebar để sử dụng AI.", "error"
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt_text,
        )
        return response.text, "gemini-2.5-flash"
    except Exception as e:
        return f"Lỗi kết nối AI: {str(e)}", "error"

# --- 3. KHỞI TẠO BỘ NHỚ TẠM (SESSION STATE) ---
if "db_thanh_vien" not in st.session_state: 
    st.session_state["db_thanh_vien"] = []
if "db_phan_cong_hien_tai" not in st.session_state: 
    st.session_state["db_phan_cong_hien_tai"] = []

# Cấu hình giao diện hiển thị diện rộng (Wide)
st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

# --- 4. VÙNG TIÊU ĐỀ CHÍNH ---
st.title("🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")
st.caption("Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk")
st.markdown("---")

# --- 5. BẢNG ĐIỀU KHIỂN SIDEBAR ---
st.sidebar.markdown("## MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")
phan_he = st.sidebar.radio(
    "Phân hệ",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed"
)

# --- 6. XỬ LÝ ĐIỀU HƯỚNG THEO PHÂN HỆ ---
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox(
        "Nội dung giảng dạy", 
        ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB"],
        label_visibility="collapsed"
    )
    
    # Ô nhập API Key động khi giáo viên cần dùng các tính năng liên quan đến AI
    st.sidebar.markdown("### 🔑 CẤU HÌNH AI")
    api_key = st.sidebar.text_input("Nhập Gemini API Key:", type="password")
    
    if menu == "1. Thiết kế KHBD": 
        # Gọi hàm mới trong file khbd_manager.py
        render_khbd_section(lambda p: run_ai_prompt_safe(p, api_key))
    elif menu == "2. Thiết kế Đề KT": 
        render_exam_designer_section("", lambda p: run_ai_prompt_safe(p, api_key))
    elif menu == "3. Đánh giá HS": 
        st.info("💡 Tính năng Đánh giá học sinh đang được phát triển...")
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()

else:  # Phân hệ Quản lý tổ chuyên môn
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox(
        "Nội dung quản lý", 
        ["1. Hệ thống Quản lý và Phân công chuyên môn giảng dạy", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"],
        label_visibility="collapsed"
    )
    
    if menu == "1. Hệ thống Quản lý và Phân công chuyên môn giảng dạy": 
        # Gọi hàm giao diện phân quyền từ org_manager.py
        render_org_section()
    elif menu == "2. Biên bản sinh hoạt": 
        # Gọi hàm quản lý biên bản từ org_manager.py
        render_meeting_minutes()
    elif menu == "3. Kế hoạch cá nhân": 
        # Gọi hàm quản lý kế hoạch từ org_manager.py
        render_personal_plan()
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        df_tv = pd.DataFrame(st.session_state["db_thanh_vien"])
        if not df_tv.empty and "Phân môn chính" in df_tv.columns:
            st.bar_chart(df_tv["Phân môn chính"].value_counts())
        else:
            st.warning("⚠️ Chưa có dữ liệu thành viên hoặc thiếu cột 'Phân môn chính' để lập biểu đồ.")
