import streamlit as st

def render_special_ed_section(run_ai_prompt_safe_func):
    # CSS Tối ưu cho giao diện dịu nhẹ, nhân văn
    st.markdown("""
        <style>
            .st-emotion-cache-1kyxreq { justify-content: center; }
            div[data-testid="stTabs"] button {
                font-size: 18px !important;
                color: #2E86C1 !important; /* Xanh lam dịu */
                font-weight: bold !important;
            }
            .title-box {
                background-color: #EBF5FB;
                padding: 15px;
                border-radius: 10px;
                border-left: 5px solid #2980B9;
                margin-bottom: 20px;
            }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div class="title-box">
            <h3 style='text-align: center; color: #154360; margin: 0;'>
                🌱 TRỢ LÝ XÂY DỰNG KẾ HOẠCH HỖ TRỢ HỌC SINH KHUYẾT TẬT / HÒA NHẬP
            </h3>
        </div>
    """, unsafe_allow_html=True)

    tab_thiet_ke, tab_quan_ly = st.tabs(["📝 XÂY DỰNG KẾ HOẠCH CÁ NHÂN (IEP)", "🗂️ HỒ SƠ LƯU TRỮ"])
    
    with tab_thiet_ke:
        # THÔNG TIN CƠ BẢN
        with st.container(border=True):
            st.markdown("<h5 style='color: #2980B9;'>1. Thông tin cơ bản của học sinh</h5>", unsafe_allow_html=True)
            col_ten, col_lop, col_dang = st.columns([2, 1, 1.5])
            with col_ten:
                ten_hs = st.text_input("Họ và tên (hoặc Mã HS):", placeholder="Ví dụ: Nguyễn Văn A")
            with col_lop:
                lop_hs = st.selectbox("Lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"])
            with col_dang:
                dang_kt = st.selectbox("Dạng khuyết tật / Khó khăn:", [
                    "Khuyết tật học tập (Khó đọc, khó tính toán...)", 
                    "Khuyết tật trí tuệ", 
                    "Khuyết tật thính giác / Ngôn ngữ", 
                    "Khuyết tật thị giác", 
                    "Tự kỷ / Rối loạn phổ tự kỷ",
                    "Khuyết tật vận động",
                    "Khác"
                ])
                
        # ĐÁNH GIÁ VÀ MỤC TIÊU
        with st.container(border=True):
            st.markdown("<h5 style='color: #2980B9;'>2. Đánh giá hiện tại & Mục tiêu</h5>", unsafe_allow_html=True)
            col_diem, col_muc = st.columns(2)
            with col_diem:
                kha_nang = st.text_area("Khả năng hiện tại (Điểm mạnh/Hạn chế):", 
                                        placeholder="- Điểm mạnh: Thích vẽ, ngoan ngoãn.\n- Hạn chế: Mất tập trung, tiếp thu chậm môn Toán.", height=120)
            with col_muc:
                muc_tieu = st.text_area("Mục tiêu giáo dục trong học kỳ:", 
                                        placeholder="- Giao tiếp hòa đồng với bạn bè.\n- Hoàn thành phép cộng trừ cơ bản.", height=120)

        # TÀI LIỆU Y TẾ & NÚT LỆNH
        st.markdown("<br>", unsafe_allow_html=True)
        col_file, col_btn = st.columns([1, 1])
        with col_file:
            ho_so_y_te = st.file_uploader("Tải lên Hồ sơ Y tế/Đánh giá tâm lý (Tùy chọn):", type=["pdf", "docx", "jpg"])
        
        with col_btn:
            st.markdown("<br>", unsafe_allow_html=True)
            nut_tao = st.button("✨ TẠO KẾ HOẠCH HỖ TRỢ TỰ ĐỘNG BẰNG AI", type="primary", use_container_width=True)

        if nut_tao:
            if not ten_hs:
                st.warning("Vui lòng nhập tên hoặc mã học sinh để tiếp tục!")
            else:
                with st.spinner("AI đang phân tích và thiết lập kế hoạch hỗ trợ chuyên biệt..."):
                    # Chỗ này sẽ gọi hàm AI của bạn
                    # prompt = f"Viết kế hoạch hỗ trợ học sinh khuyết tật tên {ten_hs}, dạng {dang_kt}..."
                    st.success("Đã tạo kế hoạch thành công!")
                    # Hiển thị kết quả tạm thời
                    st.info("Khu vực hiển thị kết quả AI sinh ra sẽ nằm ở đây.")
                    
    with tab_quan_ly:
        st.markdown("### 📂 Hồ sơ học sinh đang được hỗ trợ")
        st.write("Tại đây sẽ hiển thị danh sách các kế hoạch đã được lưu đồng bộ trên Google Sheets.")