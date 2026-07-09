import streamlit as st

def render_stem_page():
    st.markdown("## 🚀 CHỨC NĂNG XÂY DỰNG DỰ ÁN GIÁO DỤC STEM")
    st.write("Hỗ trợ giáo viên thiết kế các chủ đề STEM tích hợp.")
    
    # Khu vực nhập liệu
    ten_chu_de = st.text_input("Tên dự án / Chủ đề STEM:", placeholder="Ví dụ: Hệ thống tiết kiệm năng lượng thông minh")
    
    col1, col2 = st.columns(2)
    with col1:
        mon_chu_dao = st.selectbox("Môn học chủ đạo:", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"])
    with col2:
        doi_tuong = st.selectbox("Đối tượng học sinh:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"])
        
    tich_hop_thiet_bi = st.checkbox("Tích hợp thiết bị vi điều khiển (VD: ESP8266, Cảm biến)")
    
    if st.button("Tạo cấu trúc bài giảng STEM"):
        with st.spinner("Đang kết nối AI..."):
            # Gọi hàm xử lý Prompt AI tại đây
            st.success("Đã hoàn thiện bản thảo dự án STEM!")