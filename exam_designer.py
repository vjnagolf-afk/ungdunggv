import streamlit as st
import gspread
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard
from rag_module.latex_formatter import process_science_formulas

# --- HÀM HỖ TRỢ GOOGLE SHEETS CỦA THẦY (GIỮ NGUYÊN) ---
def get_exam_sheet():
    try:
        creds_dict = st.secrets.get("gspread_credentials") or st.secrets.get("google_sheet_creds")
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_key("1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY")
        return sh.worksheet("DE_KT")
    except: return None

def sync_exam_to_google_sheet(ten_de, mon, khoi, thoi_gian, noi_dung):
    sheet = get_exam_sheet()
    if sheet:
        sheet.append_row([ten_de, mon, khoi, thoi_gian, noi_dung])
        return True
    return False

# --- HÀM GIAO DIỆN CHÍNH (GỌI TỪ APP.PY) ---
def render_exam_designer_section(run_ai_prompt_safe_func):
    # 1. GIAO DIỆN CŨ (THẦY DÁN PHẦN CODE CŨ CỦA THẦY VÀO ĐÂY)
    # [DÁN TOÀN BỘ CODE CÁC CỘT (COLUMNS), INPUT, BẢNG CỦA THẦY Ở ĐÂY]
    
    # LƯU Ý: Đảm bảo biến 'nut_sinh_de' được gán bằng st.button(...)
    # Đảm bảo biến 'prompt_vi_mo' và 'mo_hinh_uu_tien' lấy giá trị từ các input phía trên

    # 2. XỬ LÝ AI (ĐẶT NGAY DƯỚI CÁC NÚT BẤM CŨ)
    if 'nut_sinh_de' in locals() and nut_sinh_de:
        with st.spinner("🧠 AI đang soạn đề..."):
            ket_qua, model = run_ai_prompt_safe_func(prompt_vi_mo, mo_hinh_uu_tien)
            
            # ĐỒNG BỘ: BỘ LỌC CÔNG THỨC TRƯỚC KHI HIỂN THỊ
            st.session_state["ket_qua_de"] = process_science_formulas(ket_qua)
            st.session_state["model_dung"] = model
            st.rerun()

    # 3. HIỂN THỊ KẾT QUẢ
    if st.session_state.get("ket_qua_de"):
        st.markdown("---")
        st.markdown(st.session_state["ket_qua_de"])
        
        # Nút xuất file
        data_word = export_to_docx_vietnam_standard(
            st.session_state["ket_qua_de"], 
            st.session_state.get("save_ten_de", "De_Thi")
        )
        st.download_button("📥 Tải bản chuẩn (.docx)", data_word, "De_Thi_Chuan.docx", use_container_width=True)
        
        if st.button("🗑️ Xóa kết quả"):
            st.session_state["ket_qua_de"] = ""
            st.rerun()

    # 4. DANH SÁCH LƯU ĐỀ (GIỮ NGUYÊN CODE CŨ CỦA THẦY)
