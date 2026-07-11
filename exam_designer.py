import sys
import os
import streamlit as st
import gspread

# 1. Đường dẫn module
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'teaching_assistant')))
from rag_module.latex_formatter import process_science_formulas
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard
from ai_service import run_ai_prompt_safe

# [PHẦN CẤU HÌNH GOOGLE SHEET VÀ CÁC HÀM GET/SYNC/DELETE CŨ CỦA THẦY VẪN ĐỂ NGUYÊN Ở ĐÂY]
# ... (Thầy giữ nguyên các hàm get_exam_sheet, sync_exam_to_google_sheet...) ...

def render_exam_designer_section(run_ai_prompt_safe_func):
    # [Giữ nguyên phần CSS st.markdown cũ của thầy]
    
    # [PHẦN CÁC Ô NHẬP LIỆU CŨ CỦA THẦY VẪN ĐỂ NGUYÊN]
    # ... col1, col2, col3, col4, các number_input ...
    
    # 2. Xử lý nút bấm KHÔNG CẦN TABS
    if st.button("🚀 TỰ ĐỘNG KHỞI TẠO MA TRẬN VÀ ĐỀ THI", type="primary"):
        with st.spinner("🧠 AI đang soạn đề..."):
            # Gọi AI (Thầy đảm bảo biến prompt_vi_mo đã được định nghĩa từ các ô nhập liệu bên trên)
            raw_output, model_name = run_ai_prompt_safe_func(prompt_vi_mo, mo_hinh_uu_tien)
            
            # --- ĐỒNG BỘ BỘ LỌC CÔNG THỨC ---
            final_content = process_science_formulas(raw_output)
            st.session_state["ket_qua_de"] = final_content
            st.rerun()

    # 3. HIỂN THỊ KẾT QUẢ VÀ NÚT XUẤT WORD CHUẨN (Ngay dưới nút bấm)
    if st.session_state.get("ket_qua_de"):
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(st.session_state["ket_qua_de"])
        
        col_x, col_y = st.columns(2)
        with col_x:
            # Xuất file Word bản chuẩn hành chính
            data_word = export_to_docx_vietnam_standard(
                st.session_state["ket_qua_de"], 
                st.session_state["save_ten_de"], 
                school_name=st.session_state["save_school"]
            )
            st.download_button("📥 Tải đề Word chuẩn", data_word, "De_Thi_Chuan.docx", use_container_width=True)
        with col_y:
            if st.button("🗑️ Xóa đề", use_container_width=True):
                st.session_state["ket_qua_de"] = ""
                st.rerun()

    # 4. THAY VÌ TAB, DÙNG EXPANDER ĐỂ XEM THƯ MỤC LƯU ĐỀ
    with st.expander("📂 XEM THƯ MỤC LƯU ĐỀ ĐÃ XD"):
        # [Chèn đoạn code lấy ds_de và hiển thị bằng st.expander của thầy ở đây]
        pass
