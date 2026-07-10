import streamlit as st

def render_special_ed_section(run_ai_prompt_safe_func):
    # CSS: Tăng cỡ chữ và nút bấm
    st.markdown("""<style>
        .big-font { font-size: 18px !important; }
        .stButton button { font-size: 16px !important; }
        .stSelectbox, .stTextInput { font-size: 16px !important; }
    </style>""", unsafe_allow_html=True)
    
    st.markdown("<h3 style='text-align: center; color: #154360;'>🌱 XÂY DỰNG KẾ HOẠCH HỖ TRỢ HSKT</h3>", unsafe_allow_html=True)

    # Khởi tạo kho lưu trữ file vĩnh viễn trong session
    if "my_files" not in st.session_state: st.session_state["my_files"] = {}

    tab_thiet_ke, tab_quan_ly = st.tabs(["📝 XÂY DỰNG KẾ HOẠCH", "🗂️ QUẢN LÝ HỒ SƠ"])
    
    with tab_thiet_ke:
        # Hàng menu
        col1, col2, col3, col4, col5 = st.columns(5)
        ten_hs = col1.text_input("Họ và tên HS:")
        lop_hs = col2.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Khối 10", "Khối 11", "Khối 12"])
        dang_kt = col3.selectbox("Dạng khuyết tật:", ["Vận động", "Trí tuệ", "Nghe", "Nhìn", "Tự kỷ", "Thần kinh", "Khác"])
        ky_hoc = col4.selectbox("Kế hoạch GDHSKT:", ["Kế hoạch HK I", "Kế hoạch HK II", "Cả Năm"])
        mon_hoc = col5.selectbox("Chọn môn học:", ["Ngữ văn", "Toán", "Tiếng Anh", "KHTN (Lý)", "KHTN (Hóa)", "KHTN (Sinh)", "Tin học", "Nghệ thuật", "HĐTN, hướng nghiệp"])
        
        # Khu vực quản lý file: Tải - Lưu - Xóa
        st.markdown("---")
        for label in ["KH Mẫu", "Thông tin HSKT", "Nội dung học tập HK"]:
            c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
            c1.markdown(f"**{label}**")
            
            # Hiển thị file đã lưu nếu có
            if label in st.session_state["my_files"]:
                c2.success(f"Đã có file: {st.session_state['my_files'][label].name}")
                if c4.button("🗑️", key=f"del_{label}"):
                    del st.session_state["my_files"][label]
                    st.rerun()
            else:
                uploaded = c2.file_uploader(f"Chọn file:", key=f"up_{label}")
                if c3.button("💾 Lưu", key=f"save_{label}") and uploaded:
                    st.session_state["my_files"][label] = uploaded
                    st.rerun()
        
        st.markdown("---")
        
        if st.button("✨ TẠO KẾ HOẠCH HỖ TRỢ HSKT BẰNG AI", type="primary", use_container_width=True):
            if not ten_hs:
                st.warning("Vui lòng nhập tên học sinh!")
            else:
                with st.spinner("Đang phân tích dữ liệu..."):
                    prompt = f"Lập kế hoạch hỗ trợ HS {ten_hs}, lớp {lop_hs}, dạng {dang_kt}, môn {mon_hoc}, kỳ {ky_hoc}. Bám sát file mẫu đã lưu."
                    ket_qua, _ = run_ai_prompt_safe_func(prompt)
                    st.session_state["ket_qua_hskt"] = ket_qua
                    st.markdown(ket_qua)
        
        if "ket_qua_hskt" in st.session_state and st.session_state["ket_qua_hskt"]:
            st.download_button("📥 Tải kế hoạch (Word)", data=st.session_state["ket_qua_hskt"], file_name=f"KH_{ten_hs}.docx", use_container_width=True)
                    
    with tab_quan_ly:
        st.write("📂 Danh sách hồ sơ đã lưu:")
        for name, f in st.session_state["my_files"].items():
            st.write(f"- {name}: {f.name}")
