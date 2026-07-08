import streamlit as st
import pandas as pd
import time
from google import genai

# Import các module (Đảm bảo các file này nằm cùng thư mục)
from exam_designer import render_exam_designer_section
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager
from org_manager import render_org_section, render_meeting_minutes, render_personal_plan
from khbd_manager import render_khbd_section

# --- CẤU HÌNH AI ---
def run_ai_prompt_safe(prompt_text, api_key):
    # Hàm này dùng chung cho các module
    return "Kết quả từ AI (Đây là ví dụ)", "gemini-model"

# --- KHỞI TẠO SESSION ---
if "db_thanh_vien" not in st.session_state: st.session_state["db_thanh_vien"] = []
if "db_phan_cong_hien_tai" not in st.session_state: st.session_state["db_phan_cong_hien_tai"] = []

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")
st.title("📚 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")

# --- MENU ĐIỀU HƯỚNG ---
phan_he = st.sidebar.radio("CHỌN PHÂN HỆ TÁC NGHIỆP", ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"])

if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    menu = st.sidebar.selectbox("CHỌN NỘI DUNG", 
                                ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB"])
    
    if menu == "1. Thiết kế KHBD": render_khbd_section(run_ai_prompt_safe)
    elif menu == "2. Thiết kế Đề KT": render_exam_designer_section("", run_ai_prompt_safe)
    elif menu == "3. Đánh giá HS": st.info("Đang phát triển tính năng Đánh giá...")
    elif menu == "4. Quản lý điểm (SMAS)": render_grade_manager_section()
    elif menu == "5. Quản lý TKB": render_tkb_manager()

else: # Phân hệ Quản lý tổ
    menu = st.sidebar.selectbox("QUẢN LÝ TỔ CHUYÊN MÔN", 
                                ["1. Hệ thống Quản lý tổ", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"])
    
    if menu == "1. Hệ thống Quản lý tổ": render_org_section()
    elif menu == "2. Biên bản sinh hoạt": render_meeting_minutes()
    elif menu == "3. Kế hoạch cá nhân": render_personal_plan()
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 Thống kê số liệu")
        st.bar_chart(pd.DataFrame(st.session_state.get("db_thanh_vien", [])).get("Phân môn chính", pd.Series()).value_counts())
