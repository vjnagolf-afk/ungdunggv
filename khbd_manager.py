import streamlit as st
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def render_khbd_section(api_key, run_ai_prompt_safe):
    st.header("📋 THIẾT KẾ KẾ HOẠCH BÀI DẠY (KHBD) TÍCH HỢP NĂNG LỰC SỐ")
    
    # Logic thiết kế KHBD cũ của thầy
    ten_bai = st.text_input("Tên bài học:")
    lop = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"])
    
    if st.button("🚀 XD KHBD thông minh"):
        if not api_key:
            st.error("Vui lòng nhập Gemini API Key ở sidebar!")
        else:
            with st.spinner("AI đang dựng giáo án..."):
                try:
                    res, _ = run_ai_prompt_safe(f"Soạn giáo án chuẩn 5512 cho bài {ten_bai} lớp {lop}", api_key)
                    st.markdown(res)
                except Exception as e:
                    st.error(f"Lỗi: {e}")