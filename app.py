import streamlit as st
# Import các module (Đảm bảo các file này nằm cùng thư mục với app.py)
from exam_designer import render_exam_designer_section
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager
from org_manager import render_org_section

# Cấu hình trang
st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")
st.title("📚 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")

# --- MENU CHÍNH ---
phan_he = st.sidebar.radio("CHỌN PHÂN HỆ TÁC NGHIỆP", 
                           ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"])

if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    menu = st.sidebar.selectbox("CHỌN NỘI DUNG", 
                                ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB"])
    
    if menu == "1. Thiết kế KHBD": st.info("Chức năng đang hoàn thiện...")
    elif menu == "2. Thiết kế Đề KT": render_exam_designer_section("", None)
    elif menu == "3. Đánh giá HS": st.info("Chức năng đang hoàn thiện...")
    elif menu == "4. Quản lý điểm (SMAS)": render_grade_manager_section()
    elif menu == "5. Quản lý TKB": render_tkb_manager()

else: # Phân hệ Quản lý tổ
    menu = st.sidebar.selectbox("QUẢN LÝ TỔ CHUYÊN MÔN", 
                                ["1. Hệ thống Quản lý tổ", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"])
    
    if menu == "1. Hệ thống Quản lý tổ": render_org_section()
    elif menu == "2. Biên bản sinh hoạt": st.write("Chức năng Biên bản - Đang phát triển")
    elif menu == "3. Kế hoạch cá nhân": st.write("Chức năng Kế hoạch cá nhân - Đang phát triển")
    elif menu == "4. Thống kê số liệu": st.write("Chức năng Thống kê - Đang phát triển")
