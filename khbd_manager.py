import streamlit as st

def render_khbd_section(run_ai_prompt_safe_func):
    st.subheader("📋 THIẾT KẾ KHBD THÔNG MINH")
    
    # --- CÁC THÀNH PHẦN NHẬP LIỆU (Giống hệt trong ảnh của bạn) ---
    ten_bai_hoc = st.text_input("Tên bài học:", placeholder="Ví dụ: Bài 1 - Sơ lược về bảng tuần hoàn các nguyên tố hóa học")
    
    khoi_lop = st.selectbox(
        "Khối lớp:",
        ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Lớp 10", "Lớp 11", "Lớp 12"]
    )
    
    # Bổ sung thêm cấu hình môn học để AI viết chính xác hơn
    mon_hoc = st.selectbox(
        "Môn học:",
        ["Ngữ văn", "Toán", "Tiếng Anh", "Khoa học tự nhiên", "Lịch sử & Địa lý", "Tin học", "Công nghệ"]
    )

    # --- XỬ LÝ KHI BẤM NÚT "XD KHBD" ---
    if st.button("🪄 XD KHBD"):
        if not ten_bai_hoc:
            st.warning("⚠️ Vui lòng nhập Tên bài học trước khi tạo giáo án!")
        else:
            with st.spinner("🧠 Trợ lý AI đang thiết kế kế hoạch bài dạy chi tiết..."):
                # Cấu trúc Prompt chuẩn Công văn 5512 để gửi cho Gemini
                prompt_yeu_cau = f"""
                Bạn là một chuyên gia giáo dục và trợ lý giáo viên bậc THCS/THPT tại Việt Nam.
                Hãy thiết kế một Kế hoạch bài dạy (KHBD / Giáo án) chi tiết cho:
                - Tên bài học: {ten_bai_hoc}
                - Khối lớp: {khoi_lop}
                - Môn học: {mon_hoc}
                
                Yêu cầu cấu trúc bài soạn phải tuân thủ nghiêm ngặt Phụ lục IV của Công văn 5512/BGDĐT bao gồm:
                I. MỤC TIÊU (Về Kiến thức, Năng lực cốt lõi, Năng lực đặc thù, Phẩm chất).
                II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (Của giáo viên và học sinh).
                III. TIẾN TRÌNH DẠY HỌC: Thiết kế chi tiết Hoạt động 1: Xác định vấn đề/Nhiệm vụ học tập (Mở đầu); Hoạt động 2: Hình thành kiến thức mới; Hoạt động 3: Luyện tập; Hoạt động 4: Vận dụng.
                Mỗi hoạt động phải chỉ rõ: Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện (Giao nhiệm vụ, Thực hiện nhiệm vụ, Báo cáo thảo luận, Kết luận/Nhận định).
                
                Hãy trình bày bằng định dạng Markdown rõ ràng, dễ đọc, chuyên nghiệp.
                """
                
                # Gọi hàm AI dùng chung được truyền từ app.py
                ket_qua_ai, model_used = run_ai_prompt_safe_func(prompt_yeu_cau)
                
                # Hiển thị kết quả giáo án
                st.markdown("---")
                st.success(f"✨ Đã thiết kế xong giáo án bằng mô hình: {model_used}")
                
                # Vùng hiển thị nội dung giáo án dạng văn bản định dạng
                st.markdown(ket_qua_ai)
                
                # Tính năng cộng điểm: Cho phép giáo viên tải giáo án về máy (.txt hoặc sao chép)
                st.download_button(
                    label="📥 Tải giáo án về máy (.txt)",
                    data=ket_qua_ai,
                    file_name=f"Giao_an_{ten_bai_hoc.replace(' ', '_')}.txt",
                    mime="text/plain"
                )
