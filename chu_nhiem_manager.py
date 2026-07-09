import streamlit as st
import io
from docx import Document
from st_quill import st_quill

def export_to_word(title, content_html):
    """Hàm xử lý chuyển đổi văn bản sang file Word (.docx)"""
    doc = Document()
    doc.add_heading(title, level=1)
    
    # Loại bỏ các thẻ HTML cơ bản từ Quill để đưa văn bản sạch vào file Word
    clean_text = content_html.replace("<p>", "").replace("</p>", "\n")
    clean_text = clean_text.replace("<strong>", "").replace("</strong>", "")
    clean_text = clean_text.replace("<br>", "\n").replace("<br/>", "\n")
    
    # Chia nhỏ văn bản theo dòng và ghi vào file Word
    for line in clean_text.split("\n"):
        if line.strip():
            doc.add_paragraph(line)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def render_chu_nhiem_section(run_ai_prompt_safe=None):
    st.subheader("📋 7. KẾ HOẠCH CÔNG TÁC CHỦ NHIỆM LỚP")
    st.caption("Hệ sinh thái số hỗ trợ giáo viên chủ nhiệm quản lý lớp học và lập kế hoạch thông minh.")
    
    tab_tong_quan, tab_hang_thang = st.tabs([
        "📊 Đặc điểm tình hình & Kế hoạch năm học", 
        "📅 Kế hoạch công tác theo Tháng (AI Tự động)"
    ])
    
    # =========================================================================
    # THẺ 1: ĐẶC ĐIỂM TÌNH HÌNH & KẾ HOẠCH NĂM HỌC
    # =========================================================================
    with tab_tong_quan:
        st.write("### 📌 ĐẶC ĐIỂM TÌNH HÌNH LỚP")
        col_thuan_loi, col_kho_khan = st.columns(2)
        with col_thuan_loi:
            st.text_area("Thuận lợi", value="- Nhà trường luôn quan tâm đến công tác chủ nhiệm, có những kế hoạch chỉ đạo kịp thời đến tập thể lớp...\n- Tập thể lớp có tinh thần đoàn kết...", height=150, key="ta_thuan_loi")
        with col_kho_khan:
            st.text_area("Khó khăn", value="- Sự phát triển của công nghệ thông tin có tác động đến tâm sinh lí các em, một số em mê chơi, lơ là việc học...", height=150, key="ta_kho_khan")
            
        st.write("---")
        st.write("### 📝 KẾ HOẠCH NĂM HỌC")
        col_ren_luyen, col_muc_dich = st.columns(2)
        with col_ren_luyen:
            st.text_area("Rèn luyện", value="Về rèn luyện: + Có hồ sơ quản lí học sinh khoa học, cụ thể, rõ ràng, chính xác...", height=150, key="ta_ren_luyen")
        with col_muc_dich:
            st.text_area("Mục đích yêu cầu", value="- Luôn kính trọng người trên, thầy cô giáo, cán bộ và nhân viên nhà trường...", height=150, key="ta_muc_dich")
            
        if st.button("💾 Lưu Kế hoạch năm học", type="primary", key="btn_save_nam_hoc"):
            st.success("Đã cập nhật và lưu trữ kế hoạch năm học thành công!")

    # =========================================================================
    # THÈ 2: KẾ HOẠCH CÔNG TÁC THEO THÁNG - ĐỊNH DẠNG RICH TEXT & XUẤT WORD
    # =========================================================================
    with tab_hang_thang:
        st.write("#### 🛠 CẤU HÌNH THÔNG TIN CHỦ NHIỆM")
        col_khoi, col_lop, col_thang = st.columns(3)
        
        with col_khoi:
            selected_khoi = st.selectbox("Chọn Khối lớp:", ["Khối lớp 6", "Khối lớp 7", "Khối lớp 8", "Khối lớp 9"], key="sb_khoi_lop")
        with col_lop:
            lop_dict = {"Khối lớp 6": ["6A","6B","6C","6D","6E","6F"], "Khối lớp 7": ["7A","7B","7C","7D","7E","7F"], "Khối lớp 8": ["8A","8B","8C","8D","8E","8F"], "Khối lớp 9": ["9A","9B","9C","9D","9E","9F","9G"]}
            selected_lop = st.selectbox("Chọn Lớp chủ nhiệm:", lop_dict.get(selected_khoi, ["6A"]), key="sb_lop_chu_nhiem")
        with col_thang:
            thang_options = [f"Tháng {i}/2026" for i in range(8, 13)] + [f"Tháng {i}/2027" for i in range(1, 6)]
            selected_thang = st.selectbox("Chọn Tháng công tác:", thang_options, key="sb_thang_cong_tac")
            
        st.write("---")
        ghi_chu_them = st.text_input("Yêu cầu bổ sung đặc biệt cho tháng này (nếu có):", placeholder="Ví dụ: Chuẩn bị đại hội Chi đội, phụ đạo học sinh yếu...", key="txt_ghi_chu_them")
        
        # Khởi tạo giá trị ban đầu trong session_state để lưu trữ kế hoạch tránh bị mất khi rerun
        if "id_ke_hoach_content" not in st.session_state:
            st.session_state["id_ke_hoach_content"] = "<p><i>Kế hoạch trống. Vui lòng nhấn nút Khởi tạo bên dưới để AI tự động biên soạn nội dung...</i></p>"
            
        # Nút bấm kích hoạt AI lập kế hoạch chủ động
        if st.button("🚀 Khởi tạo Kế hoạch bằng AI", type="primary", key="btn_chu_nhiem_ai"):
            if run_ai_prompt_safe is not None:
                with st.spinner(f"AI đang thiết lập kế hoạch {selected_thang}..."):
                    prompt_he_thong = f"""
                    Bạn là trợ lý AI cho giáo viên chủ nhiệm THCS Việt Nam. Hãy lập bản kế hoạch công tác cho lớp {selected_lop} trong {selected_thang}.
                    Đầu ra CỦA BẠN PHẢI sử dụng mã HTML cơ bản (thẻ <p>, <strong>, <br/>) để hiển thị chính xác cấu trúc xuống dòng đẹp mắt:
                    
                    <strong>KẾ HOẠCH CÔNG TÁC CHỦ NHIỆM LỚP {selected_lop} - {selected_thang.upper()}</strong><br/><br/>
                    <strong>1. Chủ điểm:</strong> [Xác định chủ điểm giáo dục tương ứng tháng, ví dụ Tháng 2 là 'Mừng Đảng - Mừng Xuân']<br/>
                    - [Nhiệm vụ 1]<br/>
                    - [Nhiệm vụ 2]<br/><br/>
                    <strong>2. Nội dung hoạt động:</strong><br/>
                    - [Các dòng việc lớn cần triển khai, kiểm tra nề nếp, duy trì sĩ số trước sau tết, chăm sóc CTMN...]<br/><br/>
                    <strong>* KẾ HOẠCH TỪNG TUẦN:</strong><br/>
                    - <strong>TUẦN 22:</strong> [Nhiệm vụ cụ thể từng mục nề nếp, sĩ số, sinh hoạt đầu giờ, tổng hợp sổ đầu bài...]<br/>
                    - <strong>TUẦN 23:</strong> [Nhiệm vụ cụ thể...]<br/>
                    """
                    response = run_ai_prompt_safe(prompt_he_thong)
                    # Chuyển đổi các dấu xuống dòng thông thường thành thẻ <br/> để Quill đọc hiểu định dạng dọc
                    st.session_state["id_ke_hoach_content"] = response.replace("\n", "<br/>")
            else:
                st.info("Hệ thống kết nối AI đang được đồng bộ...")

        st.write("#### 📝 KHUNG SOẠN THẢO KẾ HOẠCH THÁNG (Có thể chỉnh sửa trực tiếp)")
        
        # BƯỚC ĐỘT PHÁ: Sử dụng st_quill để tạo khung soạn thảo văn bản giàu định dạng giống hệt như ảnh mẫu
        edited_content = st_quill(
            value=st.session_state["id_ke_hoach_content"],
            html=True,
            toolbar=["bold", "italic", "underline", "strike", {"list": "ordered"}, {"list": "bullet"}, {"align": []}, "color", "background", "font", "size"],
            key="quill_editor"
        )
        
        # Cập nhật lại nội dung đã sửa đổi vào session state
        st.session_state["id_ke_hoach_content"] = edited_content
        
        # CHỨC NĂNG XUẤT FILE WORD (.DOCX)
        st.write("")
        file_name_doc = f"Ke_hoach_chu_nhiem_{selected_lop.replace(' ', '_')}_{selected_thang.replace('/', '_')}.docx"
        word_data = export_to_word(f"KẾ HOẠCH CHỦ NHIỆM LỚP {selected_lop} - {selected_thang.upper()}", edited_content)
        
        st.download_button(
            label="📥 Tải xuống file Word (.docx)",
            data=word_data,
            file_name=file_name_doc,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_download_word"
        )
