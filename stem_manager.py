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
    # THẺ 1: CÁC SẢN PHẨM STEM (ĐÃ NÂNG CẤP BỘ LỌC AI)
    # =========================================================
    with tab1:
        st.subheader("⚙️ Bộ lọc tùy chỉnh yêu cầu AI gợi ý")
        
        # Tạo 3 cột cho 3 Menu
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            chon_khoi = st.selectbox("1. Chọn khối lớp:", ["Khối 9", "Khối 8", "Khối 7", "Khối 6"], index=0)
        with col_m2:
            chon_mon = st.selectbox("2. Môn học chủ đạo:", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"], index=0)
        with col_m3:
            chon_chu_de = st.selectbox("3. Chọn Chủ đề:", [
                "Tiết kiệm năng lượng", 
                "Bảo vệ môi trường", 
                "Nông nghiệp thông minh", 
                "Nhà thông minh (Smart Home)", 
                "Chăm sóc sức khỏe", 
                "Chủ đề tự do"
            ])
            
        st.markdown("**Các lĩnh vực và yêu cầu tích hợp bổ sung:**")
        
        # Menu dạng tag cho phép chọn nhiều môn học tích hợp
        mon_tich_hop = st.multiselect(
            "Các môn học cần tích hợp (Mô hình STEM/STEAM):", 
            ["Toán học (Math)", "Khoa học (Science)", "Công nghệ (Technology)", "Kỹ thuật (Engineering)", "Nghệ thuật (Art)"]
        )
        
        # Các hộp kiểm (Checkboxes)
        tich_hop_ai = st.checkbox("🔌 Tích hợp ứng dụng AI hoặc Vi điều khiển (Ví dụ: Arduino, ESP8266)", value=True)
        tich_hop_khuyet_tat = st.checkbox("🤝 Phân hóa hoạt động và câu hỏi dành cho học sinh khuyết tật", value=True)
        
        st.markdown("---")
        
        # Nút gọi AI
        if st.button("Trợ lý AI nghiên cứu và gợi ý chủ đề mới", use_container_width=True):
            with st.spinner("AI đang phân tích các tiêu chí và tổng hợp dữ liệu..."):
                
                # Biến cờ (flag) để hiển thị chữ Có/Không trong mô phỏng kết quả
                str_ai = "Có" if tich_hop_ai else "Không"
                str_ht = "Có" if tich_hop_khuyet_tat else "Không"
                str_mon = ", ".join(mon_tich_hop) if mon_tich_hop else "Không yêu cầu"
                
                # Nội dung mô phỏng AI trả về dựa trên các tùy chọn của thầy
                st.session_state.stem_ai_suggestions = f"""
                ### 📌 Danh sách dự án STEM đề xuất cho {chon_khoi} - Môn {chon_mon}
                *(Chủ đề: {chon_chu_de} | Tích hợp: {str_mon})*
                
                **1. Hệ thống giám sát và điều khiển {chon_chu_de.lower()} tự động**
                *   **Mô tả:** Sử dụng vi điều khiển ESP8266 kết nối cảm biến để thu thập dữ liệu môi trường, tự động điều chỉnh thiết bị nhằm tối ưu hóa hiệu suất.
                *   **Giáo dục hòa nhập:** Cung cấp sơ đồ mạch điện in nổi, các khối lệnh lập trình kéo thả màu sắc tương phản cao, phân công vai trò giám sát dữ liệu phù hợp với từng nhu cầu của học sinh.
                
                **2. Mô hình trực quan ứng dụng AI trong {chon_chu_de.lower()}**
                *   **Mô tả:** Xây dựng mô hình vật lý kết hợp camera AI để nhận diện và phân loại, giúp học sinh nắm bắt thực tiễn của Khoa học Tự nhiên.
                """
        
        # Hiển thị kết quả gợi ý
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
            mon_chu_dao_t2 = st.selectbox("Môn học chủ đạo (Thẻ 2):", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"], index=0)
            lop_hoc_t2 = st.selectbox("Lớp (Thẻ 2):", ["Lớp 9", "Lớp 8", "Lớp 7", "Lớp 6"], index=0)
        with col2:
            thoi_luong = st.text_input("Thời lượng thực hiện:", placeholder="Ví dụ: 3 Tiết")
        
        st.subheader("2. Yêu cầu tích hợp & Phân hóa")
        tich_hop_ai_iot_t2 = st.checkbox("🔌 Tích hợp ứng dụng AI hoặc Vi điều khiển (Ví dụ: Arduino, ESP8266)", value=True, key="cb1_t2")
        tich_hop_hoa_nhap_t2 = st.checkbox("🤝 Phân hóa hoạt động và câu hỏi dành cho học sinh khuyết tật", value=True, key="cb2_t2")
        
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
                    **Môn học:** {mon_chu_dao_t2} | **Đối tượng:** {lop_hoc_t2} | **Thời lượng:** {thoi_luong}
                    
                    ## BƯỚC 1: XÁC ĐỊNH VẤN ĐỀ
                    Học sinh xác định bài toán thực tiễn...
                    
                    ## BƯỚC 2: NGHIÊN CỨU KIẾN THỨC NỀN
                    Nghiên cứu nguyên lý vi điều khiển và mạch điện...
                    """
                    st.session_state.stem_generated_content = noi_dung_ai
                    st.success("Thiết kế thành công!")

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
