import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))
import streamlit as st

def render_teaching_assistant_section():
    st.title("🌱 Hỗ trợ Giảng dạy")
    
    # Định nghĩa danh sách các tab
    tabs = st.tabs([
        "Hỏi-Đáp (RAG)", "Trò chơi", "Chấm bài", "Học liệu", 
        "Mô phỏng", "Phân tích", "Ngân hàng đề", "Sinh Video", "Tương tác", "Cá nhân hóa"
    ])
    
    # Kết nối tới từng module con
    with tabs[0]: # Tab Hỏi-Đáp (RAG)
        from teaching_assistant.rag_module.manager import render_rag
        render_rag()
        
    with tabs[1]:
        st.info("Module Trò chơi đang được phát triển...")
    
    # ... Các tab khác tương tự
