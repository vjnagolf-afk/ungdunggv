import streamlit as st
import pandas as pd
import sqlite3
import os
import sys
import urllib.request

# --- CẤU HÌNH ĐƯỜNG DẪN ---
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# --- IMPORT CÁC MODULE CHỨC NĂNG (IMPORT Ở ĐẦU FILE) ---
from database_manager import inject_demo_data, DB_PATH
from ai_service import run_ai_prompt_safe
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
from hskt_plan import render_special_ed_section
from teaching_assistant.main import render_teaching_assistant_section

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

# --- XỬ LÝ QUYỀN ADMIN (WHITELIST IP) ---
def get_current_public_ip():
    try:
        return urllib.request.urlopen('https://amazonaws.com').read().decode('utf-8').strip()
    except:
        return "127.0.0.1"

ADMIN_IP_WHITELIST = ["113.161.220.105", "14.161.12.34"]
current_device_ip = get_current_public_ip()
is_admin_owner = current_device_ip in ADMIN_IP_WHITELIST

# Xử lý mật mã dự phòng
if "current_entered_password" not in st.session_state:
    st.session_state["current_entered_password"] = ""

if not is_admin_owner:
    url_params = st.query_params
    if url_params.get("admin") == "123456":
        st.session_state["current_entered_password"] = "123456"
    is_admin_owner = (st.session_state["current_entered_password"] == "123456")

# --- GIAO DIỆN CHÍNH ---
st.markdown("<h1 style='text-align: center; color: darkred; font-weight: bold;'>🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- SIDEBAR & ĐIỀU HƯỚNG ---
st.sidebar.markdown("### MENU HỆ THỐNG")
phan_he = st.sidebar.radio("Chọn phân hệ:", ["Trợ lý Giảng dạy (Giáo viên)", "Hỗ trợ giảng dạy", "Trợ lý Quản lý (Tổ chuyên môn)"], label_visibility="collapsed")

if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB", "6. Thiết kế bài dạy STEM", "7. Kế hoạch công tác chủ nhiệm lớp", "8. Kế hoạch hỗ trợ học sinh khuyết tật"], label_visibility="collapsed")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "2. Thiết kế Đề KT": 
        # Gọi trực tiếp hàm đã import ở đầu file
        render_exam_designer_section(lambda p, m: run_ai_prompt_safe(p, m, is_admin_owner=is_admin_owner))
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
    elif menu == "8. Kế hoạch hỗ trợ học sinh khuyết tật":
        render_special_ed_section(lambda p: run_ai_prompt_safe(p))

elif phan_he == "Hỗ trợ giảng dạy":
    render_teaching_assistant_section()

elif phan_he == "Trợ lý Quản lý (Tổ chuyên môn)":
    # ... phần code tổ chuyên môn của thầy ...
    pass
