import streamlit as st

def render_special_ed_section(run_ai_prompt_safe_func):
    st.markdown("""
        <style>
            .title-box { background-color: #EBF5FB; padding: 15px; border-radius: 10px; border-left: 5px solid #2980B9; margin-bottom: 20px; }
            button[kind="primary"] { background-color: #5DADE2 !important; color: white !important; }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<div class='title-box'><h3 style='text-align: center; color: #154360; margin: 0;'>🌱 HỖ TRỢ XÂY DỰNG KẾ HOẠCH HỖ TRỢ HSKT</h3></div>", unsafe_allow_html=True)

    tab_thiet_ke, tab_quan_ly = st.tabs(["📝 XÂY DỰNG KẾ HOẠCH", "🗂️ QUẢN LÝ HỒ SƠ"])
    
    with tab_thiet_ke:
        with st.container(border=True):
            st.markdown("<h5 style='color: #2980B9;'>1. Thông tin học sinh</h5>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            ten_hs = col1.text_input("Họ và tên:")
            dang_kt = col2.selectbox("Dạng khuyết tật:", [
                "1. Khuyết tật vận động", "2. Khuyết tật trí tuệ (Chậm phát triển)",
                "3. Khuyết tật nghe", "4. Khuyết tật nhìn (Thị giác)",
                "5. Khuyết tật phát triển (Tự kỷ, Tăng động)",
                "6. Khuyết tật thần kinh, tâm thần", "7. Khác"
            ])
        
        with st.container(border=True):
            st.markdown("<h5 style='color: #2980B9;'>2. Kế hoạch & Môn học</h5>", unsafe_allow_html=True)
            col3, col4 = st.columns(2)
            ky_hoc = col3.selectbox("Kế hoạch GDHSKT:", ["Kế hoạch HK I", "Kế hoạch HK II", "Cả Năm"])
            mon_hoc = col4.selectbox("Chọn môn học:", [
                "Ngữ văn", "Toán", "Tiếng Anh", "Giáo dục công dân", "Lịch sử và Địa lí", 
                "KHTN (Lý)", "KHTN (Hóa học)", "KHTN (Sinh học)", "Công nghệ", 
                "Tin học", "Giáo dục thể chất", "Nghệ thuật (Âm nhạc và Mĩ thuật)", 
                "Giáo dục địa phương", "Hoạt động trải nghiệm, hướng nghiệp"
            ])
            
            uploaded_file = st.file_uploader(f"Tải lên file mẫu kế hoạch môn {mon_hoc}:")
            
            c_save, c_del = st.columns(2)
            if uploaded_file and c_save.button("💾 Lưu file môn học"):
                st.session_state[f"file_{mon_hoc}"] = uploaded_file
                st.success(f"Đã lưu file mẫu cho {mon_hoc}")
            if c_del.button("🗑️ Xóa file môn học"):
                st.session_state.pop(f"file_{mon_hoc}", None)
                st.info("Đã xóa file.")

        if st.button("✨ TẠO KẾ HOẠCH HỖ TRỢ TỰ ĐỘNG BẰNG AI", type="primary", use_container_width=True):
            with st.spinner("AI đang thiết lập kế hoạch..."):
                file_info = "Có sử dụng file mẫu." if f"file_{mon_hoc}" in st.session_state else "Không có file mẫu."
                prompt = f"Lập kế hoạch hỗ trợ HSKT dạng {dang_kt}, môn {mon_hoc}, {ky_hoc}. {file_info}"
                ket_qua, _ = run_ai_prompt_safe_func(prompt, "3.1 Flash-Lite")
                st.markdown(ket_qua)
                    
    with tab_quan_ly:
        st.write("📂 Quản lý hồ sơ cá nhân học sinh.")
