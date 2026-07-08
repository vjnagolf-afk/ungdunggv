import streamlit as st
# Import các module đã tách riêng
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager
from org_manager import render_org_section
# Import các module khác nếu có (exam_designer...)

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")
st.title("📚 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")

# --- MENU CHÍNH ---
phan_he = st.sidebar.radio("CHỌN PHÂN HỆ", ["Trợ lý Giảng dạy", "Trợ lý Quản lý"])

if phan_he == "Trợ lý Giảng dạy":
    chuc_nang = st.sidebar.selectbox("NỘI DUNG", ["Quản lý điểm (SMAS)", "Thời khóa biểu"])
    
    if chuc_nang == "Quản lý điểm (SMAS)":
        render_grade_manager_section()
    elif chuc_nang == "Thời khóa biểu":
        render_tkb_manager()

else: # Phân hệ Quản lý tổ
    chuc_nang = st.sidebar.selectbox("NỘI DUNG", ["Hệ thống Tổ chuyên môn", "Thống kê số liệu"])
    
    if chuc_nang == "Hệ thống Tổ chuyên môn":
        render_org_section()
    elif chuc_nang == "Thống kê số liệu":
        st.header("📊 Thống kê số liệu")
        # Logic thống kê của thầy ở đây
