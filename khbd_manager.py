import streamlit as st
import docx  
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
from pypdf import PdfReader
import gspread
from datetime import datetime

from math_compiler import process_runs_with_math, generate_plot_stream

# ================= CÁC HÀM HỖ TRỢ GOOGLE SHEETS =================
SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY' 

def get_khbd_sheet():
    try:
        creds_dict = None
        for key in ["gspread_credentials", "GSPREAD_CREDENTIALS", "google_sheet_creds"]:
            if key in st.secrets: creds_dict = st.secrets[key]; break
        if not creds_dict: return None
        gc = gspread.service_account_from_dict(creds_dict)
        return gc.open_by_key(SHEET_ID).worksheet("KHBD")
    except: return None

def save_khbd_to_sheet(ten_bai, lop, bo_sach, thoi_luong, content):
    sheet = get_khbd_sheet()
    if sheet:
        sheet.append_row([ten_bai, lop, bo_sach, thoi_luong, content, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        return True
    return False

def get_all_khbd_from_sheet():
    sheet = get_khbd_sheet()
    return sheet.get_all_records() if sheet else []

def delete_khbd_from_sheet(row_index):
    sheet = get_khbd_sheet()
    if sheet:
        sheet.delete_rows(row_index + 2)
        return True
    return False

# ================= GIAO DIỆN KHBD ĐẦY ĐỦ =================
def render_khbd_section(run_ai_prompt_safe_func):
    st.markdown("<h3 style='text-align: center; color: blue;'>🧠 TRỢ LÝ THIẾT KẾ KẾ HOẠCH BÀI DẠY (KHBD) AI</h3>", unsafe_allow_html=True)
    
    if "ket_qua_khbd" not in st.session_state: st.session_state["ket_qua_khbd"] = ""
    
    # Tăng cỡ chữ cho Tab bằng CSS
    st.markdown("""
        <style>
            button[data-baseweb="tab"] {
                font-size: 20px !important;
                font-weight: bold !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    tab_thiet_ke, tab_thu_vien = st.tabs(["📝 THIẾT KẾ KHBD TỰ ĐỘNG", "🗄️ THƯ VIỆN BÀI SOẠN"])
    
    with tab_thiet_ke:
        ten_bai = st.text_input("Tên bài học / Chủ đề bài dạy:", placeholder="Ví dụ: Bài 4: Tốc độ chuyển động")
        
        # Thiết kế 4 menu trên cùng 1 hàng
        col1, col2, col3, col4 = st.columns(4)
        with col1: lop_khbd = st.selectbox("Lớp:", [f"Lớp {i}" for i in range(6, 13)])
        with col2: kieu_khbd = st.selectbox("Mẫu thiết kế:", ["Chuẩn 5512", "Rút gọn", "STEM"])
        with col3: thoi_luong = st.text_input("Thời lượng (Tiết):", value="2")
        with col4: files_tailieu = st.file_uploader("Tài liệu (docx, pdf, txt):", accept_multiple_files=True)
        
        # Checkbox bám sát 100%
        chk_strict_file = st.checkbox("🚩 Bám sát 100% tài liệu tải lên", value=False)
        
        # Nút chức năng
        c1, c2 = st.columns(2)
        if c1.button("⚡ Thiết kế bài dạy bằng AI", type="primary", use_container_width=True):
            if ten_bai:
                with st.spinner("Đang soạn giáo án (Bộ sách: Kết nối tri thức)..."):
                    ket_qua, _ = run_ai_prompt_safe_func(f"Soạn KHBD chuẩn 5512 bài: {ten_bai} (Bộ sách: Kết nối tri thức). Tài liệu: {chk_strict_file}")
                    st.session_state["ket_qua_khbd"] = ket_qua
        
        if c2.button("💾 Lưu tạm vào Google Sheets", use_container_width=True):
            if st.session_state["ket_qua_khbd"]:
                if save_khbd_to_sheet(ten_bai, lop_khbd, "Kết nối tri thức", thoi_luong, st.session_state["ket_qua_khbd"]):
                    st.success("✅ Đã lưu!")
        
        if st.session_state["ket_qua_khbd"]:
            st.markdown(st.session_state["ket_qua_khbd"])
                
    with tab_thu_vien:
        st.write("### 📂 Danh sách bài soạn")
        ds_bai = get_all_khbd_from_sheet()
        for idx, bai in enumerate(ds_bai):
            with st.expander(f"{bai.get('Tên bài', 'Bài soạn')} ({bai.get('Thời gian', 'N/A')})"):
                c1, c2 = st.columns(2)
                if c1.button("📥 Gọi bài", key=f"load_{idx}"):
                    st.session_state["ket_qua_khbd"] = bai['Nội dung chi tiết']
                    st.rerun()
                if c2.button("🗑️ Xóa bài soạn", key=f"del_{idx}"):
                    if delete_khbd_from_sheet(idx):
                        st.success("✅ Đã xóa!")
                        st.rerun()
