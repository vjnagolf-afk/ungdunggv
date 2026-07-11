import streamlit as st

def render_rag():
    st.subheader("📚 AI Hỏi - Đáp theo tài liệu")
    
    uploaded_file = st.file_uploader("Tải lên tài liệu của bạn (PDF, Docx):", type=["pdf", "docx"])
    
    if uploaded_file:
        st.success(f"Đã tải lên: {uploaded_file.name}")
        # Tại đây bạn sẽ gọi hàm xử lý từ processor.py
        user_query = st.text_input("Nhập câu hỏi của bạn về tài liệu này:")
        if st.button("Hỏi AI"):
            st.write("Đang truy xuất thông tin...")
            # Logic gọi AI sẽ được viết ở đây
