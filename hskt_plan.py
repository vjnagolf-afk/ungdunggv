import streamlit as st

def render_special_ed_section(run_ai_prompt_safe_func):
    st.markdown("""<style>
        .title-box { background-color: #EBF5FB; padding: 15px; border-radius: 10px; border-left: 5px solid #2980B9; margin-bottom: 20px; }
        button[kind="primary"] { background-color: #5DADE2 !important; color: white !important; }
    </style>""", unsafe_allow_html=True)
    
    st.markdown("<div class='title-box'><h3 style='text-align: center; color: #154360; margin: 0;'>🌱 HỖ TRỢ XÂY DỰNG KẾ HOẠCH HỖ TRỢ HSKT</h3></div>", unsafe_allow_html=True)

    tab_thiet_ke, tab_quan_ly = st.tabs(["📝 XÂY DỰNG KẾ HOẠCH", "🗂️ QUẢN LÝ HỒ SƠ"])
    
    with tab_thiet_ke:
        # Bốn menu nằm trên 1 hàng
        col1, col2, col3, col4 = st.columns(4)
        ten_hs = col1.text_input("Họ và tên HS:")
        dang_kt = col2.selectbox("Dạng khuyết tật:", ["Khuyết tật vận động", "Khuyết tật trí tuệ (Chậm phát triển)", "Khuyết tật nghe", "Khuyết tật nhìn (Thị giác)", "Khuyết tật phát triển (Tự kỷ, Tăng động)", "Khuyết tật thần kinh, tâm thần", "Khác"])
        ky_hoc = col3.selectbox("Kế hoạch GDHSKT:", ["Kế hoạch HK I", "Kế hoạch HK II", "Cả Năm"])
        mon_hoc = col4.selectbox("Chọn môn học:", ["Ngữ văn", "Toán", "Tiếng Anh", "Giáo dục công dân", "Lịch sử và Địa lí", "KHTN (Lý)", "KHTN (Hóa học)", "KHTN(Sinh học)", "Công nghệ", "Tin học", "Giáo dục thể chất", "Nghệ thuật", "Giáo dục địa phương", "Hoạt động trải nghiệm, hướng nghiệp"])
        
        # Nút tải file mẫu cùng hàng
        c_f1, c_f2 = st.columns(2)
        uploaded_file = c_f1.file_uploader(f"Tải mẫu KH môn {mon_hoc}:")
        info_file = c_f2.file_uploader("Tải mẫu 'Thông tin về HSKT':")
        
        if st.button("✨ TẠO KẾ HOẠCH HỖ TRỢ TỰ ĐỘNG BẰNG AI", type="primary", use_container_width=True):
            with st.spinner("AI đang thiết lập kế hoạch..."):
                prompt = f"Lập kế hoạch hỗ trợ HSKT dạng {dang_kt}, môn {mon_hoc}, {ky_hoc}."
                # Gọi hàm AI
                ket_qua, _ = run_ai_prompt_safe_func(prompt)
                st.markdown(ket_qua)
                    
    with tab_quan_ly:
        st.write("📂 Quản lý hồ sơ cá nhân học sinh.")
