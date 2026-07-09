import streamlit as st
import io
from docx import Document

# Hàm hỗ trợ tạo file Word từ nội dung văn bản
def create_word_file(title, content):
    doc = Document()
    doc.add_heading(title, 0)
    doc.add_paragraph(content)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def render_stem_section():
    st.markdown("## 🚀 HỆ SINH THÁI GIÁO DỤC STEM")
    st.markdown("---")
    
    # ---------------------------------------------------------
    # KHỞI TẠO BỘ NHỚ (SESSION STATE) CHO TẤT CẢ CÁC THẺ
    # ---------------------------------------------------------
    if "stem_generated_content" not in st.session_state:
        st.session_state.stem_generated_content = ""
    if "stem_saved_projects" not in st.session_state:
        st.session_state.stem_saved_projects = {}
    if "stem_ai_suggestions" not in st.session_state:
        st.session_state.stem_ai_suggestions = ""

    # ---------------------------------------------------------
    # TẠO 3 THẺ (TABS) ĐỘC LẬP
    # ---------------------------------------------------------
    tab1, tab2, tab3 = st.tabs([
        "💡 1. CÁC SẢN PHẨM STEM", 
        "🛠️ 2. XÂY DỰNG DỰ ÁN", 
        "📁 3. QUẢN LÝ DỰ ÁN ĐÃ LƯU"
    ])

    # =========================================================
    # THẺ 1: GỢI Ý VÀ DANH SÁCH SẢN PHẨM STEM
    # =========================================================
    with tab1:
        st.subheader("Khám phá & Đề xuất chủ đề STEM")
        st.info("Trợ lý AI sẽ phân tích chương trình học và xu hướng công nghệ để đề xuất các chủ đề STEM phù hợp nhất.")
        
        if st.button("Trợ lý AI nghiên cứu và gợi ý chủ đề", use_container_width=True):
            with st.spinner("AI đang tổng hợp dữ liệu..."):
                # GỌI API AI TẠI ĐÂY ĐỂ LẤY GỢI Ý. Dưới đây là nội dung mô phỏng (Mockup):
                st.session_state.stem_ai_suggestions = """
                ### 📌 Danh sách dự án STEM nổi bật:
                1. **Hệ thống giám sát và cảnh báo chất lượng không khí:** Tích hợp bộ vi điều khiển theo dõi nồng độ bụi mịn.
                2. **Mô hình nhà kính nông nghiệp tự động hóa:** Áp dụng hệ thống cảm biến đo độ ẩm và tự động tưới tiêu.
                
                ### ✨ Đề xuất chủ đề mới từ AI (Tối ưu cho KHTN Lớp 9):
                *   **Hệ thống tiết kiệm năng lượng thông minh:** Sử dụng vi điều khiển ESP8266 và cảm biến chuyển động để tự động bật/tắt thiết bị điện, giúp giảm thiểu lãng phí điện năng trong trường học.
                *   **Thiết kế công cụ hỗ trợ học tập hòa nhập:** Ứng dụng công nghệ in 3D và vật liệu tái chế để chế tạo các mô hình trực quan hỗ trợ học sinh có nhu cầu đặc biệt.
                """
        
        # Hiển thị kết quả gợi ý nếu đã có trong bộ nhớ
        if st.session_state.stem_ai_suggestions:
            with st.container(border=True):
                st.markdown(st.session_state.stem_ai_suggestions)

    # =========================================================
    # THẺ 2: CHỨC NĂNG XÂY DỰNG DỰ ÁN GIÁO DỤC STEM
    # =========================================================
    with tab2:
        st.subheader("1. Thông tin chung")
        ten_chu_de = st.text_input("Tên dự án / Chủ đề STEM:", 
                                   placeholder="Ví dụ: Thiết kế hệ thống tiết kiệm năng lượng sử dụng cảm biến...")
        
        col1, col2 = st.columns(2)
        with col1:
            mon_chu_dao = st.selectbox("Môn học chủ đạo:", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"])
            lop_hoc = st.selectbox("Lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"])
        with col2:
            thoi_luong = st.text_input("Thời lượng thực hiện:", placeholder="Ví dụ: 3 Tiết")
        
        st.subheader("2. Yêu cầu tích hợp & Phân hóa")
        tich_hop_ai_iot = st.checkbox("🔌 Tích hợp ứng dụng AI hoặc Vi điều khiển (Ví dụ: Arduino, ESP8266)")
        tich_hop_hoa_nhap = st.checkbox("🤝 Phân hóa hoạt động và câu hỏi dành cho học sinh khuyết tật (Giáo dục hòa nhập)")
        
        st.subheader("3. Tài liệu tham khảo (Học liệu nguồn)")
        tai_lieu_nguon = st.file_uploader("Tải lên file bài học, tài liệu tham khảo (PDF, Word, TXT)", accept_multiple_files=True)
        
        st.markdown("---")
        
        if st.button("Kích hoạt AI thiết kế tiến trình STEM", type="primary", use_container_width=True):
            if not ten_chu_de:
                st.warning("Vui lòng nhập tên chủ đề STEM!")
            else:
                with st.spinner("Hệ thống đang phân tích tài liệu và thiết kế bài dạy..."):
                    # GỌI API AI THIẾT KẾ GIÁO ÁN TẠI ĐÂY (Nội dung giả lập):
                    noi_dung_ai = f"""
                    # GIÁO ÁN STEM: {ten_chu_de}
                    **Môn học:** {mon_chu_dao} | **Đối tượng:** {lop_hoc} | **Thời lượng:** {thoi_luong}
                    
                    ## BƯỚC 1: XÁC ĐỊNH VẤN ĐỀ
                    Học sinh tìm hiểu về thực trạng...
                    
                    ## BƯỚC 2: NGHIÊN CỨU KIẾN THỨC NỀN
                    Nghiên cứu nguyên lý hoạt động của các thiết bị...
                    """
                    st.session_state.stem_generated_content = noi_dung_ai
                    st.success("Tạo kế hoạch bài dạy STEM thành công!")

        # Khu vực hiển thị kết quả và tải file bên trong Thẻ 2
        if st.session_state.stem_generated_content != "":
            st.markdown("### 📖 KẾT QUẢ THIẾT KẾ")
            with st.container(border=True):
                st.markdown(st.session_state.stem_generated_content)
            
            col_download, col_save = st.columns(2)
            with col_download:
                docx_file = create_word_file(ten_chu_de, st.session_state.stem_generated_content)
                st.download_button(
                    label="📥 Tải giáo án về máy (File Word)",
                    data=docx_file,
                    file_name=f"Giao_an_STEM_{ten_chu_de}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with col_save:
                if st.button("💾 Lưu dự án vào hệ thống", use_container_width=True):
                    st.session_state.stem_saved_projects[ten_chu_de] = st.session_state.stem_generated_content
                    st.toast(f"Đã lưu thành công dự án: {ten_chu_de}", icon="✅")

    # =========================================================
    # THẺ 3: QUẢN LÝ DỰ ÁN ĐÃ LƯU
    # =========================================================
    with tab3:
        st.subheader("Danh sách các dự án đang lưu trữ")
        
        if len(st.session_state.stem_saved_projects) > 0:
            danh_sach_du_an = list(st.session_state.stem_saved_projects.keys())
            
            for ten_da in danh_sach_du_an:
                with st.expander(f"📌 {ten_da}"):
                    # Hiển thị nội dung dự án
                    st.markdown(st.session_state.stem_saved_projects[ten_da])
                    
                    # Nút xóa dự án
                    if st.button("🗑️ Xóa dự án này", key=f"btn_del_{ten_da}"):
                        del st.session_state.stem_saved_projects[ten_da]
                        st.rerun() # Tải lại giao diện ngay lập tức
        else:
            st.info("Hiện tại chưa có dự án nào được lưu trong hệ thống.")
