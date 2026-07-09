import streamlit as st
import io
from docx import Document

def export_to_word(title, content_text):
    """Hàm xử lý tạo và tải file Word (.docx) sạch sẽ từ nội dung text"""
    doc = Document()
    doc.add_heading(title, level=1)
    
    # Ghi từng dòng văn bản vào file Word
    for line in content_text.split("\n"):
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
    # THÈ 2: KẾ HOẠCH CÔNG TÁC THEO THÁNG - HIỂN THỊ ĐỊNH DẠNG DỌC & TẢI WORD
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
        ghi_chu_them = st.text_input("Yêu cầu bổ sung đặc biệt cho tháng này (nếu có):", placeholder="Ví dụ: Chuẩn bị văn nghệ, uốn nắn học sinh cá biệt...", key="txt_ghi_chu_them")
        
        # Khởi tạo vùng lưu trữ nội dung trong phiên làm việc (session state)
        if "content_ke_hoach_thang" not in st.session_state:
            st.session_state["content_ke_hoach_thang"] = ""
            
        # Nút bấm ra lệnh cho AI soạn thảo tự động
        if st.button("🚀 Khởi tạo Kế hoạch bằng AI", type="primary", key="btn_chu_nhiem_ai"):
            if run_ai_prompt_safe is not None:
                with st.spinner(f"AI đang thiết lập kế hoạch {selected_thang}..."):
                    prompt_he_thong = f"""
                    Bạn là trợ lý AI cho giáo viên chủ nhiệm THCS Việt Nam. Hãy lập một bản kế hoạch công tác chủ nhiệm cực kỳ chi tiết cho lớp {selected_lop} trong {selected_thang}.
                    
                    YÊU CẦU ĐẦU RA PHẢI ĐƯỢC PHÂN TÁCH DÒNG RÕ RÀNG THEO ĐÚNG ĐỊNH DẠNG CẤU TRÚC SAU:
                    
                    KẾ HOẠCH THÁNG {selected_thang.split('/')[0].replace('Tháng ', '')}/{selected_thang.split('/')[-1]}
                    1. Chủ điểm: [Tên chủ điểm tương ứng, ví dụ Tháng 2 là 'Mừng Đảng - Mừng Xuân']
                    - [Nhiệm vụ thi đua 1]
                    - [Nhiệm vụ thi đua 2]
                    
                    2. Nội dung hoạt động:
                    - [Đầu việc lớn 1: Duy trì sĩ số học sinh, nề nếp...]
                    - [Đầu việc lớn 2: Vệ sinh cảnh quan trường lớp, công trình măng non...]
                    - [Đầu việc lớn 3: Phong trào thi đua tại liên đội...]
                    
                    * KẾ HOẠCH TỪNG TUẦN:
                    - TUẦN 22: (Ổn định nề nếp, sinh hoạt 15 phút đầu giờ thường xuyên theo đúng chủ đề, nhắc nhở giờ giấc học tập, học bài và làm bài đầy đủ, tổng hợp sổ đầu bài vào thứ 7...)
                    - TUẦN 23: (...)
                    - TUẦN 24: (...)
                    
                    Yêu cầu bổ sung từ giáo viên (nếu có): {ghi_chu_them}
                    Cấm viết tất cả văn bản trên cùng một dòng ngang. Hãy dùng dấu xuống dòng liên tục giữa các mục để tạo cấu trúc dọc đẹp mắt.
                    """
                    response = run_ai_prompt_safe(prompt_he_thong)
                    st.session_state["content_ke_hoach_thang"] = response
            else:
                st.info("Hệ thống kết nối AI đang được đồng bộ...")

        # HIỂN THỊ KHUNG VĂN BẢN ĐỊNH DẠNG DỌC (Cho phép giáo viên chỉnh sửa trực tiếp)
        st.write("#### 📝 KHUNG SOẠN THẢO KẾ HOẠCH THÁNG")
        
        edited_text = st.text_area(
            label="Nội dung kế hoạch hiển thị theo cấu trúc dọc (Bấm vào để tự do sửa đổi):",
            value=st.session_state["content_ke_hoach_thang"],
            height=400,
            key="ta_main_editor"
        )
        
        # Cập nhật dữ liệu chỉnh sửa của giáo viên vào bộ nhớ tạm
        st.session_state["content_ke_hoach_thang"] = edited_text
        
        # NÚT XUẤT FILE WORD (.DOCX) ĐÃ ĐƯỢC ĐỒNG BỘ
        if st.session_state["content_ke_hoach_thang"].strip():
            st.write("")
            file_name_doc = f"Ke_hoach_chu_nhiem_{selected_lop}_{selected_thang.replace('/', '_')}.docx"
            word_file_bytes = export_to_word(f"KẾ HOẠCH CHỦ NHIỆM LỚP {selected_lop} - {selected_thang.upper()}", edited_text)
            
            st.download_button(
                label="📥 Tải xuống file Word (.docx)",
                data=word_file_bytes,
                file_name=file_name_doc,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="btn_download_word"
            )
