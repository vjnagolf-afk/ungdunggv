import streamlit as st

def render_khbd_section(run_ai_prompt_safe):
    st.header("📋 THIẾT KẾ KHBD THÔNG MINH")
    ten_bai = st.text_input("Tên bài học:")
    lop = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"])
    
    if st.button("🚀 XD KHBD"):
        # Gọi hàm AI từ app.py truyền qua
        res, model = run_ai_prompt_safe(f"Soạn giáo án bài {ten_bai}", "api_key_dummy")
        st.markdown(res)
