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

# Import từ file math_compiler
from math_compiler import process_runs_with_math, generate_plot_stream

# ================= HÀM HỖ TRỢ GOOGLE SHEETS =================
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

# ================= CÁC HÀM XỬ LÝ VĂN BẢN VÀ EXPORT =================
def extract_context_from_uploaded_files(uploaded_files):
    # ... (Giữ nguyên logic cũ của bạn ở đây) ...
    return "", []

def set_paragraph_spacing(paragraph, before_pt=3.0, after_pt=4.5):
    # ... (Giữ nguyên logic cũ của bạn ở đây) ...
    pass

def export_khbd_to_docx(markdown_content, images_list):
    # ... (Giữ nguyên logic export đã sửa ở phản hồi trước) ...
    return b""

# ================= GIAO DIỆN CHÍNH (ĐÃ FIX LỖI NAMEERROR) =================
def render_khbd_section(run_ai_prompt_safe_func):
    st.markdown("<h3 style='text-align: center; color: blue;'>🧠 TRỢ LÝ THIẾT KẾ KHBD AI</h3>", unsafe_allow_html=True)
    
    if "ket_qua_khbd" not in st.session_state: st.session_state["ket_qua_khbd"] = ""
    
    tab_thiet_ke, tab_thu_vien = st.tabs(["📝 THIẾT KẾ", "🗄️ THƯ VIỆN"])
    
    with tab_thiet_ke:
        ten_bai = st.text_input("Tên bài học:")
        if st.button("⚡ Thiết kế bài dạy"):
            if ten_bai:
                with st.spinner("Đang soạn giáo án..."):
                    ket_qua, _ = run_ai_prompt_safe_func(f"Soạn KHBD chuẩn 5512 bài: {ten_bai}")
                    st.session_state["ket_qua_khbd"] = ket_qua
            else:
                st.warning("Vui lòng nhập tên bài!")
        
        if st.session_state["ket_qua_khbd"]:
            st.markdown(st.session_state["ket_qua_khbd"])
            if st.button("💾 Lưu tạm thời vào Google Sheets"):
                if save_khbd_to_sheet(ten_bai or "Bài không tên", "Lớp 7", "Kết nối", "2 tiết", st.session_state["ket_qua_khbd"]):
                    st.success("✅ Đã lưu vào Sheet!")
                
    with tab_thu_vien:
        st.write("### 📂 Danh sách bài soạn đã lưu")
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
