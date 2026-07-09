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
    
    # KHỞI TẠO BỘ NHỚ (SESSION STATE)
    if "stem_generated_content" not in st.session_state:
        st.session_state.stem_generated_content = ""
    if "stem_saved_projects" not in st.session_state:
        st.session_state.stem_saved_projects = {}
    if "stem_ai_suggestions" not in st.session_state:
        st.session_state.stem_ai_suggestions = ""

    # LỆNH TẠO 3 THẺ ĐỘC LẬP
    tab1, tab2, tab3 = st.tabs([
        "💡 1. CÁC SẢN PHẨM STEM", 
        "🛠️ 2. CHỨC NĂNG XÂY DỰNG DỰ ÁN GIÁO DỤC STEM", 
        "📁 3. QUẢN LÝ DỰ ÁN ĐÃ LƯU"
    ])

    # =========================================================
    # THẺ 1: CÁC SẢN PHẨM STEM
    # =========================================================
    with tab1:
        st.subheader("Danh sách & Gợi ý Chủ đề STEM từ AI")
        
        if st.button("Trợ lý AI nghiên cứu và gợi ý chủ đề mới", use_container_width=True):
            with st.spinner("AI đang tổng hợp dữ liệu..."):
                # Nội dung AI trả về
                st.session_state.stem_ai_suggestions = """
                ### 📌 Các dự án STEM phổ biến:
                1. **Mô hình nhà kính nông nghiệp thông minh** (KHTN, Công nghệ)
                2. **Hệ thống cảnh báo rò rỉ khí gas tự động** (KHTN, Tin học)
                
                ### ✨ Đề xuất chủ đề mới (Tối ưu cho KHTN Lớp 9):
                *   **Dự án 1: Hệ thống tiết kiệm năng lượng thông minh:** Sử dụng vi điều khiển ESP8266 và cảm biến để thiết kế hệ thống tự động ngắt điện, giảm thiểu lãng phí điện năng trong lớp học.
                *   **Dự án 2: Công cụ học tập hòa nhập:** Ứng dụng STEM chế tạo các thiết bị vật lý trực quan hỗ trợ học sinh có nhu cầu giáo dục đặc biệt trong môn Khoa học tự nhiên.
                """
        
        # Hiển thị kết quả
        if st.session_state.stem_ai_suggestions:
            with st.container(border=True):
                st.markdown(st.session_state.stem_ai_suggestions)

    # =========================================================
    # THẺ 2: CHỨC NĂNG XÂY DỰNG DỰ ÁN GIÁO DỤC STEM
    # =========================================================
    with tab2:
        st.subheader("1. Thông tin chung")
        ten_chu_de = st.text_input("Tên dự án / Chủ đề STEM:", 
                                   placeholder="Ví dụ: Thiết kế hệ thống tiết kiệm năng lượng...")
        
        col1, col2 = st.columns(2)
        with col1:
            # Mặc định là KHTN
            mon_chu_dao = st.selectbox("Môn học chủ đạo:", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"], index=0)
            # Mặc định là Lớp 9
            lop_hoc = st.selectbox("Lớp:", ["Lớp 9", "Lớp 8", "Lớp 7", "Lớp 6"], index=0)
        with col2:
            thoi_luong = st.text_input("Thời lượng thực hiện:", placeholder="Ví dụ: 3 Tiết")
        
        st.subheader("2. Yêu cầu tích hợp & Phân hóa")
        # Tick sẵn các tùy chọn công nghệ và hòa nhập
        tich_hop_ai_iot = st.checkbox("🔌 Tích hợp ứng dụng AI hoặc Vi điều khiển (Ví dụ: Arduino, ESP8266)", value=True)
        tich_hop_hoa_nhap = st.checkbox("🤝 Phân hóa hoạt động và câu hỏi dành cho học sinh khuyết tật", value=True)
        
        st.subheader("3. Tài liệu tham khảo (Học liệu nguồn)")
        tai_lieu_nguon = st.file_uploader("Tải lên file bài học, tài liệu tham khảo (PDF, Word)", accept_multiple_files=True)
        
        st.markdown("---")
        
        if st.button("Kích hoạt AI thiết kế tiến trình STEM", type="primary", use_container_width=True):
            if not ten_chu_de:
                st.warning("Vui lòng nhập tên chủ đề STEM!")
            else:
                with st.spinner("Hệ thống đang thiết kế..."):
                    noi_dung_ai = f"""
                    # GIÁO ÁN STEM: {ten_chu_de}
                    **Môn học:** {mon_chu_dao} | **Đối tượng:** {lop_hoc} | **Thời lượng:** {thoi_luong}
                    
                    ## BƯỚC 1: XÁC ĐỊNH VẤN ĐỀ
                    Học sinh xác định bài toán lãng phí điện năng...
                    
                    ## BƯỚC 2: NGHIÊN CỨU KIẾN THỨC NỀN
                    Nghiên cứu nguyên lý vi điều khiển và mạch điện...
                    """
                    st.session_state.stem_generated_content = noi_dung_ai
                    st.success("Thiết kế thành công!")

        # Khu vực hiển thị và lưu
        if st.session_state.stem_generated_content != "":
            st.markdown("### 📖 KẾT QUẢ THIẾT KẾ")
            with st.container(border=True):
                st.markdown(st.session_state.stem_generated_content)
            
            col_download, col_save = st.columns(2)
            with col_download:
                docx_file = create_word_file(ten_chu_de, st.session_state.stem_generated_content)
                st.download_button(
                    label="📥 Tải giáo án (File Word)",
                    data=docx_file,
                    file_name=f"Giao_an_STEM.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with col_save:
                # Nút này sẽ đẩy dự án sang thẻ 3
                if st.button("💾 Lưu dự án vào hệ thống", use_container_width=True):
                    st.session_state.stem_saved_projects[ten_chu_de] = st.session_state.stem_generated_content
                    st.toast("Đã lưu thành công! Hãy kiểm tra ở Thẻ 3.", icon="✅")

    # =========================================================
    # THẺ 3: QUẢN LÝ DỰ ÁN ĐÃ LƯU
    # =========================================================
    with tab3:
        st.subheader("Danh sách các dự án đang lưu trữ")
        
        if len(st.session_state.stem_saved_projects) > 0:
            danh_sach_du_an = list(st.session_state.stem_saved_projects.keys())
            
            for ten_da in danh_sach_du_an:
                with st.expander(f"📌 {ten_da}"):
                    st.markdown(st.session_state.stem_saved_projects[ten_da])
                    
                    if st.button("🗑️ Xóa dự án này", key=f"btn_del_{ten_da}"):
                        del st.session_state.stem_saved_projects[ten_da]
                        st.rerun() 
        else:
            st.info("Hiện tại chưa có dự án nào được lưu. Thầy hãy tạo và lưu dự án ở Thẻ 2 nhé.")
