import streamlit as st
import sys
import os

# Đường dẫn gốc
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

def render_teaching_assistant_section():
    st.title("🌱 Hỗ trợ Giảng dạy")
    
    # Định nghĩa các tab chức năng
    tabs = st.tabs([
        "Hỏi-Đáp (RAG)", 
        "Ra đề kiểm tra", 
        "Xây dựng KHBD", 
        "Học liệu", 
        "Phân tích"
    ])
    
    # Liên kết module RAG (đã hoàn thiện)
    with tabs[0]:
        from teaching_assistant.rag_module.manager import render_rag
        render_rag()
        
    # Liên kết module Ra đề (Thầy cần tạo file exam_module/manager.py)
    with tabs[1]:
        # Gọi trực tiếp tệp ở thư mục gốc
        import exam_designer
        exam_designer.render_exam_interface() # Thầy đảm bảo hàm chính trong file exam_designer.py tên là render_exam_interface            
    # Liên kết module KHBD (Thầy cần tạo file lesson_plan_module/manager.py)
    with tabs[2]:
        try:
            from teaching_assistant.lesson_plan_module.manager import render_lesson_plan
            render_lesson_plan()
        except ImportError:
            st.warning("Module KHBD đang được thiết lập. Vui lòng kiểm tra file `lesson_plan_module/manager.py`.")

    # ... các tab khác ...
