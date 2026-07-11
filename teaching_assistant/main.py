import streamlit as st
import sys
import os

# Thêm thư mục gốc của dự án vào đường dẫn tìm kiếm của Python
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

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
        
    with tabs[2]:
        st.info("Module Chấm bài đang được phát triển...")
        
    with tabs[3]:
        st.info("Module Học liệu đang được phát triển...")
        
    with tabs[4]:
        st.info("Module Mô phỏng đang được phát triển...")
        
    with tabs[5]:
        st.info("Module Phân tích đang được phát triển...")
        
    with tabs[6]:
        st.info("Module Ngân hàng đề đang được phát triển...")
        
    with tabs[7]:
        st.info("Module Sinh Video đang được phát triển...")
        
    with tabs[8]:
        st.info("Module Tương tác đang được phát triển...")
        
    with tabs[9]:
        st.info("Module Cá nhân hóa đang được phát triển...")
