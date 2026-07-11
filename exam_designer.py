import sys
import os
import streamlit as st
import gspread
import re

# Đảm bảo đường dẫn module
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'teaching_assistant')))

from rag_module.latex_formatter import process_science_formulas
from rag_module.document_export import export_to_docx
from ai_service import run_ai_prompt_safe
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard

# ================= CẤU HÌNH GOOGLE SHEETS =================
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

# ================= HÀM GIAO DIỆN CHÍNH =================
def render_exam_designer_section(run_ai_prompt_safe_func):
    
    # [Giữ nguyên phần CSS st.markdown đã có ở trên của thầy]
    
    if "ket_qua_de" not in st.session_state: st.session_state["ket_qua_de"] = ""

    tab_tao_de, tab_thu_muc = st.tabs(["CHỨC NĂNG TẠO ĐỀ KIỂM TRA", "THƯ MỤC LƯU ĐỀ ĐÃ XD"])
    
    with tab_tao_de:
        # [Giữ nguyên phần các ô nhập liệu/selectbox của thầy]
        
        nut_sinh_de = st.button("🚀 TỰ ĐỘNG KHỞI TẠO MA TRẬN VÀ ĐỀ THI", type="primary")

        if nut_sinh_de:
            with st.spinner("🧠 AI đang soạn đề..."):
                # Gọi AI
                raw_output, model_name = run_ai_prompt_safe_func(prompt_vi_mo, mo_hinh_uu_tien)
                
                # CHUẨN HÓA CÔNG THỨC (Đồng bộ qua bộ lọc)
                final_content = process_science_formulas(raw_output)
                st.session_state["ket_qua_de"] = final_content
                st.rerun()

        # HIỂN THỊ KẾT QUẢ ĐÃ XỬ LÝ
        if st.session_state["ket_qua_de"]:
            st.markdown(st.session_state["ket_qua_de"])
            
            c_save, c_del = st.columns([1, 1])
            with c_save:
                # Xuất file Word chuẩn hành chính
                data_word = export_to_docx_vietnam_standard(
                    st.session_state["ket_qua_de"], 
                    st.session_state["save_ten_de"], 
                    school_name=st.session_state["save_school"]
                )
                st.download_button("📥 Tải Đề Kiểm Tra (.docx)", data_word, "De_Thi.docx")
            
            with c_del:
                if st.button("🗑️ Xóa đề hiện tại"):
                    st.session_state["ket_qua_de"] = ""
                    st.rerun()

    with tab_thu_muc:
        # [Giữ nguyên logic get_all_exams_from_sheet của thầy]
        pass
