import sys
import os
import streamlit as st
from datetime import datetime

# 1. Ép đường dẫn để import các file ở thư mục gốc (app.py, ai_service.py,...)
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))

# 2. Import các module cần thiết
from ai_service import run_ai_prompt_safe 
from .processor import process_and_vectorize, query_rag, backup_to_googlesheet

def render_rag():
    st.subheader("📚 AI Hỏi - Đáp theo tài liệu")
    
    uploaded_file = st.file_uploader("Tải lên tài liệu (PDF, Docx):", type=["pdf", "docx"])
    
    if uploaded_file:
        # Lưu file tạm
        temp_path = "temp_file" + os.path.splitext(uploaded_file.name)[1]
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        if st.button("Bắt đầu xử lý tài liệu"):
            with st.spinner("Đang băm nhỏ và nhúng tài liệu..."):
                vectorstore = process_and_vectorize(temp_path)
                st.session_state["vectorstore"] = vectorstore
                st.success("Xử lý xong! Bạn có thể đặt câu hỏi.")

    # 3. Khu vực đặt câu hỏi
    if "vectorstore" in st.session_state:
        question = st.text_input("Nhập câu hỏi của thầy/cô:")
        if question:
            with st.spinner("AI đang tìm kiếm tài liệu..."):
                # Lấy dữ liệu liên quan từ processor
                context = query_rag(st.session_state["vectorstore"], question)
                
                # Tạo prompt gửi cho AI
                prompt = f"Dựa vào ngữ cảnh tài liệu sau đây, hãy trả lời câu hỏi của giáo viên.\n\nNgữ cảnh: {context}\n\nCâu hỏi: {question}"
                
                # Gọi AI
                response = run_ai_prompt_safe(prompt)
                
                st.markdown("---")
                st.markdown("### 🤖 Phản hồi từ AI:")
                st.write(response)
                
                # Sao lưu tự động vào GSheet
                backup_to_googlesheet({
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'query': question,
                    'response': response
                })
                st.success("✅ Đã sao lưu vào Nhật ký giảng dạy!")
