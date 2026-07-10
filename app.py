# app.py - Bản vá dứt điểm lỗi ImportError đường dẫn tương đối trên Streamlit Cloud
import streamlit as st
import pandas as pd
import sqlite3
import os
import sys

# 🚀 THUẬT TOÁN ĐƯỜNG DẪN: Ép hệ thống tìm module trong cùng thư mục chạy ứng dụng
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 🚀 BỎ DẤU CHẤM ĐỂ TRÁNH LỖI ĐỊNH DẠNG PACKAGE CỦA PYTHON
from database_manager import check_if_admin_device, inject_demo_data, DB_PATH
from ai_service import run_ai_prompt_safe

# Nhúng các phân hệ tác nghiệp vệ tinh của thầy/cô
from exam_designer import render_exam_designer_section 
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_manager import render_khbd_section  
from danh_gia_manager import render_assessment_section

from org_manager import render_org_section
from bien_ban_manager import render_meeting_minutes
from ke_hoach_ca_nhan_manager import render_personal_plan 
from stem_manager import render_stem_section
from chu_nhiem_manager import render_chu_nhiem_section 

# Khởi chạy quét thiết bị nhận diện Admin ngay khi nạp trang
is_admin_owner = check_if_admin_device()

# Khởi tạo bộ nhớ tạm đồng bộ session
if "db_thanh_vien" not in st.session_state: st.session_state["db_thanh_vien"] = []
if "db_phan_cong_hien_tai" not in st.session_state: st.session_state["db_phan_cong_hien_tai"] = []

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

# Tiêu đề giao diện chính
st.markdown("<h1 style='text-align: center; color: darkred; font-weight: bold;'>🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #0056b3; font-weight: bold; font-size: 16px;'>Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk</p>", unsafe_allow_html=True)
st.markdown("---")

# --- MENU ĐIỀU HƯỚNG TỔNG TẠI SIDEBAR ---
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_root_key_2026_v9"
)

# --- KHỐI HIỂN THỊ Ô NHẬP KEY THEO THIẾT BỊ ĐỐI TƯỢNG ---
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔑 TRẠNG THÁI TÀI KHOẢN")

if is_admin_owner:
    st.sidebar.success("👑 Thiết bị: Chủ dự án (Admin)")
    st.sidebar.caption("Tự động kích hoạt quyền đặc quyền chạy trực tiếp bằng Key hệ thống.")
    st.session_state["gv_api_key_input"] = ""
else:
    st.sidebar.warning("🔒 Thiết bị: Thành viên/Giáo viên")
    st.sidebar.caption("Vui lòng dán API Key cá nhân từ Google AI Studio để mở khóa phân hệ.")
    st.sidebar.text_input("Nhập API Key Gemini của thầy/cô:", type="password", placeholder="AIzaSy...", key="gv_api_key_input")
    if st.session_state["gv_api_key_input"]:
        st.sidebar.success("🟢 Đã nhận diện Key cá nhân.")

# --- KHỐI ĐIỀU HƯỚNG TÁC NGHIỆP CHI TIẾT ---
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB","6. Thiết kế bài dạy STEM","7. Kế hoạch công tác chủ nhiệm lớp"], label_visibility="collapsed", key="menu_gv_selectbox_v9")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p, m: run_ai_prompt_safe(p, m, is_admin_owner))
    elif menu == "2. Thiết kế Đề KT": 
        render_exam_designer_section(lambda p, m: run_ai_prompt_safe(p, m, is_admin_owner))
    elif menu == "3. Đánh giá HS": 
        render_assessment_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()
    elif menu == "6. Thiết kế bài dạy STEM":
        render_stem_section()
    elif menu == "7. Kế hoạch công tác chủ nhiệm lớp":
        render_chu_nhiem_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))

else:  # Phân hệ Quản lý tổ chuyên môn
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox("Nội dung quản lý", ["1. Quản lý & Phân công chuyên môn", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"], label_visibility="collapsed", key="menu_ql_selectbox_v9")
    
    if menu == "1. Quản lý & Phân công chuyên môn": 
        render_org_section()
    elif menu == "2. Biên bản sinh hoạt": 
        render_meeting_minutes(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "3. Kế hoạch cá nhân": 
        render_personal_plan(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        
        df_tv = pd.DataFrame()
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                query = "SELECT m.fullname as [Họ và tên], m.main_subject as [Phân môn chính], a.total_periods as [Số tiết/Tuần] FROM org_members m LEFT JOIN org_assignments a ON m.fullname = a.fullname"
                df_tv = pd.read_sql_query(query, conn)
                conn.close()
            except Exception as e:
                st.error(f"⚠️ Lỗi kết nối cơ sở dữ liệu nội bộ: {str(e)}")
        
        thuc_te_co_du_lieu = not df_tv.empty and len(df_tv) > 0 and not (df_tv["Phân môn chính"] == "").all()

        if not thuc_te_co_du_lieu:
            st.warning("ℹ️ Hiện tại chưa có dữ liệu giáo viên nào được nhập từ phân hệ '1. Quản lý & Phân công chuyên môn'.")
            if st.button("💡 Nạp nhanh dữ liệu mẫu để thử nghiệm biểu đồ", type="primary", use_container_width=True):
                try:
                    inject_demo_data()
                    st.success("🎉 Đã nạp dữ liệu thử nghiệm trực tiếp vào SQLite!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Không thể nạp dữ liệu mẫu: {str(e)}")
        else:
            st.subheader("📋 Danh sách phân công hiện tại:")
            st.dataframe(df_tv, use_container_width=True)
