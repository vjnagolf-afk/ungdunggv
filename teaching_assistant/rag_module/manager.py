import sys
import os
import streamlit as st
from datetime import datetime
from ai_service import run_ai_prompt_safe # Import từ file gốc ở thư mục ngoài

# Sử dụng import tương đối để tránh lỗi đường dẫn
from .processor import process_and_vectorize, query_rag, backup_to_googlesheet

def render_rag():
    # ... (giữ nguyên phần logic của thầy)
def render_rag():
    uploaded_file = st.file_uploader("Tải lên tài liệu:", type=["pdf", "docx"])
    
    if uploaded_file:
        # Lưu file tạm để xử lý
        with open("temp_file" + os.path.splitext(uploaded_file.name)[1], "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        if st.button("Bắt đầu xử lý tài liệu"):
            with st.spinner("Đang băm nhỏ và nhúng tài liệu..."):
                vectorstore = process_and_vectorize("temp_file" + os.path.splitext(uploaded_file.name)[1])
                st.session_state["vectorstore"] = vectorstore
                st.success("Xử lý xong! Bạn có thể đặt câu hỏi.")

    if "vectorstore" in st.session_state:
        question = st.text_input("Nhập câu hỏi của thầy/cô:")
        if question:
            context = query_rag(st.session_state["vectorstore"], question)
            # Sau đó gửi context + question vào hàm run_ai_prompt_safe của thầy
            st.write("Dựa trên tài liệu, tôi trả lời như sau...")
