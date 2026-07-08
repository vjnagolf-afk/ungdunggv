import streamlit as st
import docx  # Dùng để đọc file Word
from pypdf import PdfReader  # Dùng để đọc file PDF

# --- HÀM TRÍCH XUẤT VĂN BẢN TỪ FILE TẢI LÊN ---
def extract_text_from_files(uploaded_files):
    combined_text = ""
    for file in uploaded_files:
        try:
            if file.name.endswith('.docx'):
                doc = docx.Document(file)
                fullText = [para.text for para in doc.paragraphs]
                combined_text += f"\n--- NỘI DUNG TỪ FILE: {file.name} ---\n" + "\n".join(fullText)
            elif file.name.endswith('.pdf'):
                reader = PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() or ""
                combined_text += f"\n--- NỘI DUNG TỪ FILE: {file.name} ---\n" + text
            elif file.name.endswith('.txt'):
                combined_text += f"\n--- NỘI DUNG TỪ FILE: {file.name} ---\n" + file.read().decode("utf-8")
        except Exception as e:
            st.error(f"Lỗi khi đọc file {file.name}: {str(e)}")
    return combined_text

# --- GIAO DIỆN CHÍNH CỦA PHÂN HỆ ---
def render_khbd_section(run_ai_prompt_safe_func):
    st.markdown("<h3 style='text-align: center; color: red;'>📖 CHỨC NĂNG SOẠN GIÁO ÁN AI THEO MẪU KHÁCH</h3>", unsafe_allow_index=True)
    
    # --- DÒNG 1: NHẬP TÊN BÀI DẠY ---
    ten_bai = st.text_input("Tên bài dạy / Chủ đề:", key="khbd_ten_bai")
    
    # --- DÒNG 2: MÔN HỌC, THỜI LƯỢNG, LỚP ---
    col_mon, col_tg, col_lop = st.columns([2, 1, 1])
    with col_mon:
        mon_hoc = st.selectbox("Môn học:", ["Khoa học tự nhiên", "Toán học", "Ngữ văn", "Tiếng Anh", "Lịch sử & Địa lý", "Tin học", "Công nghệ"])
    with col_tg:
        thoi_luong = st.text_input("Thời lượng:", placeholder="Số")
    with col_lop:
        lop = st.text_input("Lớp:", placeholder="Ví dụ: 6A")
        
    # --- DÒNG 3: CÁC TÙY CHỌN CHECKBOX ---
    tich_hop_ai = st.checkbox("Tích hợp giáo dục AI (Năng lực số và AI)", value=True)
    uati_bam_sat = st.checkbox("Ưu tiên bám sát 100% tài liệu tải lên", value=True)
    
    # --- DÒNG 4: KHU VỰC CÁC NÚT BẤM VÀ TẢI FILE ---
    col_f1, col_f2, col_blank, col_btn1, col_btn2 = st.columns([1.2, 1.2, 1, 1.2, 1.4])
    
    with col_f1:
        # Cho phép tải lên nhiều file học liệu cùng lúc (Word, PDF, TXT)
        tai_hoc_lieu = st.file_uploader("📁 Tải file học liệu", type=["docx", "pdf", "txt"], accept_multiple_files=True, label_visibility="collapsed")
    with col_f2:
        tai_giao_an_mau = st.file_uploader("📄 Tải giáo án mẫu", type=["docx", "pdf", "txt"], label_visibility="collapsed")
        st.caption("Chưa chọn mẫu" if not tai_giao_an_mau else f"✅ Đã chọn mẫu: {tai_giao_an_mau.name}")
        
    # Khởi tạo vùng chứa nội dung giáo án trong session để lưu kết quả
    if "ket_qua_giao_an" not in st.session_state:
        st.session_state["ket_qua_giao_an"] = ""

    with col_btn1:
        # Nút xuất/tải file về máy bài soạn đã tạo
        st.download_button(
            label="💾 Tải file KHBD về máy",
            data=st.session_state["ket_qua_giao_an"],
            file_name=f"KHBD_{ten_bai.replace(' ', '_') if ten_bai else 'Chua_dat_ten'}.txt",
            mime="text/plain",
            disabled=(st.session_state["ket_qua_giao_an"] == "")
        )
    with col_btn2:
        nut_chay_ai = st.button("⚡ Khởi tạo giáo án bằng AI", type="primary")

    # --- DÒNG 5: HIỂN THỊ DANH SÁCH FILE ĐÃ TẢI LÊN CÓ NÚT XÓA ---
    st.markdown("**Danh sách học liệu tham khảo đã tải lên (Bấm X để xóa file lỗi):**")
    
    # Quản lý danh sách file trong bộ nhớ tạm để xử lý xóa
    if "danh_sách_files" not in st.session_state:
        st.session_state["danh_sách_files"] = []
        
    if tai_hoc_lieu:
        st.session_state["danh_sách_files"] = tai_hoc_lieu

    # Giao diện hộp hiển thị danh sách file
    with st.container(border=True):
        if not st.session_state["danh_sách_files"]:
            st.caption("Chưa nạp học liệu tham khảo nào.")
        else:
            # Tạo danh sách file có nút bấm xóa từng file bên cạnh
            for idx, f in enumerate(st.session_state["danh_sách_files"]):
                c_name, c_del = st.columns([0.9, 0.1])
                c_name.write(f"📄 {f.name} ({round(f.size/1024, 1)} KB)")
                if c_del.button("❌", key=f"del_{idx}"):
                    st.session_state["danh_sách_files"].pop(idx)
                    st.rerun()

    # --- XỬ LÝ LOGIC KHI BẤM NÚT KHỞI TẠO GIÁO ÁN BẰNG AI ---
    if nut_chay_ai:
        if not ten_bai:
            st.warning("⚠️ Vui lòng nhập Tên bài dạy / Chủ đề trước khi khởi tạo!")
        elif not st.session_state["danh_sách_files"]:
            st.warning("⚠️ Vui lòng tải lên ít nhất một file học liệu tham khảo để AI đọc nguồn dữ liệu!")
        else:
            with st.spinner("🧠 Trợ lý AI đang đọc kỹ các file nguồn và thiết kế bài soạn..."):
                # 1. Trích xuất toàn bộ dữ liệu chữ từ các file nguồn học liệu
                nguon_van_ban = extract_text_from_files(st.session_state["danh_sách_files"])
                
                # 2. Trích xuất dữ liệu từ giáo án mẫu nếu có
                mau_giao_an_text = ""
                if tai_giao_an_mau:
                    mau_giao_an_text = extract_text_from_files([tai_giao_an_mau])

                # 3. Xây dựng Prompt kỹ thuật cực kỳ chi tiết ép AI tuân thủ cấu trúc của bạn
                prompt_yeu_cau = f"""
                Bạn là một chuyên gia giáo dục cao cấp tại Việt Nam. Nhiệm vụ của bạn là biên soạn Kế hoạch bài dạy (Giáo án) chi tiết.
                
                THÔNG TIN BÀI DẠY:
                - Tên bài dạy / Chủ đề: {ten_bai}
                - Môn học: {mon_hoc}
                - Thời lượng: {thoi_luong} tiết (Yêu cầu phân chia nội dung và tiến trình khít và bám sát chính xác tổng số {thoi_luong} tiết này).
                - Khối lớp: {lop}
                
                YÊU CẦU BẮT BUỘC KHHI SOẠN BÀI:
                1. { 'BÁM SÁT 100% FILE NGUỒN TÀI LIỆU: Đọc kỹ phần nội dung văn bản trích xuất từ file nguồn bên dưới. Không tự ý bịa đặt kiến thức nằm ngoài phạm vi tài liệu này cung cấp.' if uati_bam_sat else '' }
                2. BỐ CỤC BÀI SOẠN: Phải tuân thủ chuẩn chỉnh, chính xác theo cấu trúc PHỤ LỤC IV CÔNG VĂN 5512/BGDĐT bao gồm:
                   I. MỤC TIÊU (Kiến thức; Năng lực; Phẩm chất).
                   II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU.
                   III. TIẾN TRÌNH DẠY HỌC (Gồm đầy đủ 4 hoạt động: 1. Mở đầu/Xác định vấn đề; 2. Hình thành kiến thức mới; 3. Luyện tập; 4. Vận dụng). Mỗi hoạt động bắt buộc có đủ: Mục tiêu, Nội dung, Sản phẩm, Tổ chức thực hiện.
                3. { 'TÍCH HỢP NĂNG LỰC SỐ VÀ AI: Ở phần Mục tiêu (Năng lực) và trong các bước Tổ chức thực hiện ở mục Tiến trình dạy học, phải lồng ghép chi tiết các hoạt động yêu cầu học sinh sử dụng thiết bị số, ứng dụng công nghệ thông tin hoặc khai thác các công cụ trí tuệ nhân tạo (AI) để giải quyết nhiệm vụ học tập.' if tich_hop_ai else '' }
                
                {f'MẪU THIẾT KẾ GIÁO ÁN THAM KHẢO (Hãy bắt chước phong cách trình bày từ mẫu này): {mau_giao_an_text}' if mau_giao_an_text else ''}
                
                DỮ LIỆU NGUỒN HỌC LIỆU ĐỂ TRÍCH XUẤT KIẾN THỨC BÀI HỌC:
                {nguon_van_ban}
                
                Hãy trình bày văn bản rõ ràng, chuyên nghiệp dưới định dạng Markdown, phân tách tiêu đề bằng các dấu gạch ngang đẹp mắt.
                """
                
                # Gọi hàm chạy Gemini AI
                ket_qua_ai, model_used = run_ai_prompt_safe_func(prompt_yeu_cau)
                
                # Lưu kết quả vào session state để giữ lại trạng thái hiển thị và cho phép download
                st.session_state["ket_qua_giao_an"] = ket_qua_ai
                st.success(f"✨ Đã khởi tạo bài soạn thành công bằng mô hình {model_used}!")
                st.rerun()

    # --- DÒNG 6: KHU VỰC HIỂN THỊ NỘI DUNG GIÁO ÁN XEM TRƯỚC ---
    st.markdown("**Nội dung giáo án hiển thị xem trước:**")
    with st.container(border=True):
        if st.session_state["ket_qua_giao_an"]:
            st.markdown(st.session_state["ket_qua_giao_an"])
        else:
            st.write("Nội dung bài soạn sau khi khởi tạo bằng AI sẽ hiển thị ở đây...")
