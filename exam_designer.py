import streamlit as st
import gspread
import sys
import os

# Cấu hình đường dẫn cho các module
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'teaching_assistant')))
from rag_module.latex_formatter import process_science_formulas
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard
from ai_service import run_ai_prompt_safe

# ================= ĐOẠN 1: CẤU HÌNH & GOOGLE SHEETS =================
SPREADSHEET_ID = "1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY"
TAB_NAME = "DE_KT"

def get_exam_sheet():
    try:
        creds_dict = st.secrets.get("gspread_credentials") or st.secrets.get("google_sheet_creds")
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_key(SPREADSHEET_ID)
        return sh.worksheet(TAB_NAME)
    except:
        return None

def sync_exam_to_google_sheet(ten_de, mon, khoi, thoi_gian, noi_dung):
    sheet = get_exam_sheet()
    if sheet:
        sheet.append_row([ten_de, mon, khoi, thoi_gian, noi_dung])
        return True
    return False

def get_all_exams_from_sheet():
    sheet = get_exam_sheet()
    return sheet.get_all_records() if sheet else []

def delete_exam_from_sheet(row_index):
    sheet = get_exam_sheet()
    if sheet:
        sheet.delete_rows(row_index + 2)
        return True
    return False

# ================= ĐOẠN 2: GIAO DIỆN CHÍNH =================
def render_exam_designer_section(run_ai_prompt_safe_func):
    # CSS CỦA THẦY
    st.markdown("""<style>...[GIỮ NGUYÊN CSS CŨ CỦA THẦY]...</style>""", unsafe_allow_html=True)
    
    if "ket_qua_de" not in st.session_state: st.session_state["ket_qua_de"] = ""
    
    # 1. GIAO DIỆN INPUT (COLUMNS, SELECTBOX... GIỮ NGUYÊN)
    # [THẦY DÁN PHẦN GIAO DIỆN CŨ TỪ DÒNG 80 ĐẾN 210 TRONG FILE GỐC VÀO ĐÂY]
    
    # Nút bấm cần được định nghĩa ở đây để tránh NameError
    nut_sinh_de = st.button("🚀 TỰ ĐỘNG KHỞI TẠO MA TRẬN VÀ ĐỀ THI", type="primary", use_container_width=True)
    mo_hinh_uu_tien = st.selectbox("Mô hình xử lý đề thi:", ["3.1 Flash-Lite", "3.5 Flash", "3.1 Pro", "Tư duy mở rộng"], index=0)

    # 2. LOGIC GỌI AI
    if nut_sinh_de:
        with st.spinner("🧠 Hệ thống đang lập bảng..."):
            # Gọi AI
            raw_output, model_thuc_te = run_ai_prompt_safe_func(prompt_vi_mo, mo_hinh_uu_tien)
            
            # ĐỒNG BỘ: BỘ LỌC CÔNG THỨC Ở ĐÂY
            final_content = process_science_formulas(raw_output)
            
            st.session_state["ket_qua_de"] = final_content
            st.session_state["model_dung"] = model_thuc_te
            st.rerun()

    # 3. HIỂN THỊ KẾT QUẢ VÀ XUẤT WORD
    if st.session_state["ket_qua_de"]:
        st.markdown("---")
        st.markdown(st.session_state["ket_qua_de"])
        
        c_save, c_del = st.columns(2)
        with c_save:
            data_word = export_to_docx_vietnam_standard(
                st.session_state["ket_qua_de"], 
                st.session_state.get("save_ten_de", "De_Thi"), 
                school_name=st.session_state.get("save_school", "Trường THCS")
            )
            st.download_button("📥 Tải bản chuẩn (.docx)", data_word, "De_Thi.docx", use_container_width=True)
        with c_del:
            if st.button("🗑️ Xóa kết quả", use_container_width=True):
                st.session_state["ket_qua_de"] = ""
                st.rerun()

    # 4. THƯ MỤC LƯU ĐỀ (Giữ nguyên giao diện cũ của thầy)
    st.markdown("### THƯ MỤC LƯU ĐỀ ĐÃ XD")
    # [THẦY DÁN PHẦN DISPLAY DANH SÁCH TỪ GOOGLE SHEET CŨ VÀO ĐÂY]
