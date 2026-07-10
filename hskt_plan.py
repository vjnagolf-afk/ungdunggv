import streamlit as st

def render_special_ed_section(run_ai_prompt_safe_func):
    # CSS căn chỉnh: Tăng size chữ và giữ bố cục hàng ngang
    st.markdown("""<style>
        .big-font { font-size: 18px !important; }
        .stButton button { font-size: 14px !important; }
        .stTextInput, .stSelectbox { font-size: 16px !important; }
    </style>""", unsafe_allow_html=True)

    # Khởi tạo session để giữ file không bị mất
    if "my_files" not in st.session_state: st.session_state["my_files"] = {}

    # Hàng 1: 5 Menu trên 1 hàng
    col1, col2, col3, col4, col5 = st.columns(5)
    ten_hs = col1.text_input("Họ và tên HS:")
    lop_hs = col2.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Khối 10", "Khối 11", "Khối 12"])
    dang_kt = col3.selectbox("Dạng khuyết tật:", ["Vận động", "Trí tuệ", "Nghe", "Nhìn", "Tự kỷ", "Thần kinh", "Khác"])
    ky_hoc = col4.selectbox("Kế hoạch GDHSKT:", ["Kế hoạch HK I", "Kế hoạch HK II", "Cả Năm"])
    mon_hoc = col5.selectbox("Chọn môn học:", ["Ngữ văn", "Toán", "Tiếng Anh", "KHTN (Lý)", "KHTN (Hóa)", "KHTN (Sinh)", "Tin học", "Nghệ thuật", "HĐTN, hướng nghiệp"])

    st.markdown("---")

    # Bố cục hàng file: Nhãn | Upload | Nút Lưu
    for label in ["KH Mẫu", "Thông tin HSKT", "Nội dung học tập HK"]:
        c_label, c_up, c_save = st.columns([1, 4, 1])
        c_label.markdown(f"**{label}**")
        
        uploaded = c_up.file_uploader(f"Chọn file:", key=f"up_{label}")
        if c_save.button("💾 Lưu", key=f"save_{label}") and uploaded:
            st.session_state["my_files"][label] = uploaded
            st.success(f"Đã lưu {label}")
        
        # Nếu đã lưu, hiện thông báo
        if label in st.session_state["my_files"]:
            c_up.info(f"Đang dùng file: {st.session_state['my_files'][label].name}")

    st.markdown("---")

    # Nút Tạo AI
    if st.button("✨ TẠO KẾ HOẠCH HỖ TRỢ HSKT BẰNG AI", type="primary", use_container_width=True):
        with st.spinner("Đang xử lý..."):
            prompt = f"Lập KH hỗ trợ HS {ten_hs}, môn {mon_hoc}, kỳ {ky_hoc}. Bám sát file mẫu đã lưu."
            ket_qua, _ = run_ai_prompt_safe_func(prompt)
            st.session_state["ket_qua_hskt"] = ket_qua
            st.markdown(ket_qua)

    # Nút Tải Word
    if "ket_qua_hskt" in st.session_state and st.session_state["ket_qua_hskt"]:
        st.download_button("📥 Tải kế hoạch (Word)", data=st.session_state["ket_qua_hskt"], file_name="Ke_hoach_HSKT.docx", use_container_width=True)
