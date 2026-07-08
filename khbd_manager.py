import streamlit as st
import docx  
from docx.shared import Inches, Pt
import io
import re  # <-- THÊM THƯ VIỆN NÀY ĐỂ SỬA LỖI NAMEEROR
from pypdf import PdfReader

# --- HÀM TRÍCH XUẤT VĂN BẢN VÀ HÌNH ẢNH TỪ FILE NGUỒN ---
def extract_context_from_uploaded_files(uploaded_files):
    combined_text = ""
    extracted_images = [] 
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.docx'):
                doc = docx.Document(file)
                for paragraph in doc.paragraphs:
                    if paragraph.text:
                        combined_text += paragraph.text + "\n"
                for table in doc.tables:
                    for row in table.rows:
                        text_row = [cell.text for cell in row.cells]
                        combined_text += " | ".join(text_row) + "\n"
                for rel in doc.part.relations.values():
                    if "image" in rel.target_ref:
                        extracted_images.append(rel.target_part.blob)
                        
            elif file.name.endswith('.pdf'):
                reader = PdfReader(file)
                for page in reader.pages:
                    combined_text += (page.extract_text() or "") + "\n"
                    for img_file_object in page.images:
                        extracted_images.append(img_file_object.data)
                        
            elif file.name.endswith('.txt'):
                combined_text += file.read().decode("utf-8") + "\n"
        except Exception as e:
            st.error(f"Lỗi khi xử lý file {file.name}: {str(e)}")
            
    return combined_text, extracted_images

# --- HÀM KHỞI TẠO FILE WORD CHUẨN ĐỊNH DẠNG (CHỐNG LỖI CÔNG THỨC & BIỂU BẢNG) ---
def export_khbd_to_docx(markdown_content, images_list):
    doc = docx.Document()
    
    # Định dạng lề trang chuẩn hành chính Việt Nam (Top: 2cm, Bottom: 2cm, Left: 3cm, Right: 1.5cm)
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)

    lines = markdown_content.split('\n')
    in_table = False
    table_data = []

    for line in lines:
        # Tự động dựng biểu bảng thực tế chống lỗi tràn khung dòng của Markdown
        if line.strip().startswith('|') and line.strip().endswith('|'):
            if '---|' in line or ':---|' in line:
                continue
            in_table = True
            cells = [c.strip() for c in line.split('|')[1:-1]]
            table_data.append(cells)
            continue
        else:
            if in_table and table_data:
                num_rows = len(table_data)
                num_cols = len(table_data) if num_rows > 0 else 0
                if num_cols > 0:
                    word_table = doc.add_table(rows=num_rows, cols=num_cols)
                    word_table.style = 'Table Grid'
                    for r_idx, row in enumerate(table_data):
                        for c_idx, val in enumerate(row):
                            if c_idx < num_cols:
                                word_table.cell(r_idx, c_idx).text = val
                in_table = False
                table_data = []

        # Tối ưu font chữ tránh lỗi định dạng Toán học, Hóa học
        if line.startswith('#'):
            level = len(line) - len(line.lstrip('#'))
            title_text = line.lstrip('#').strip()
            heading = doc.add_heading(title_text, level=min(level, 3))
            heading.style.font.name = 'Times New Roman'
        elif line.strip():
            # Tự động chèn lại hình ảnh minh họa từ kho lưu trữ nhị phân
            if "[Hình ảnh minh họa]" in line and images_list:
                try:
                    img_stream = io.BytesIO(images_list[0]) # Lấy ảnh đầu tiên trong kho nạp
                    doc.add_picture(img_stream, width=Inches(4.5))
                except:
                    p = doc.add_paragraph(line)
                    p.style.font.name = 'Times New Roman'
            else:
                p = doc.add_paragraph()
                p.style.font.name = 'Times New Roman'
                # Xử lý nhanh các chỉ số hóa học viết thường dưới dòng (Subscript)
                parts = re.split(r'(\d+)', line)
                for part in parts:
                    run = p.add_run(part)
                    if part.isdigit() and any(x in line for x in ['H2O', 'CO2', 'Fe', 'O2', 'H2SO4', 'N2', 'CH4', 'C2H5OH']):
                        run.font.subscript = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
# --- GIAO DIỆN CHÍNH CỦA PHÂN HỆ ---
def render_khbd_section(run_ai_prompt_safe_func):
    
    # 1. Khởi tạo 2 Thẻ điều hướng theo cấu trúc thiết kế mới
    tab_xay_dung, tab_luu_khbd = st.tabs(["📝 XÂY DỰNG KẾ HOẠCH BÀI DẠY AI", "🗄️ LƯU KHBD ĐÃ XD"])
    
    # Khởi tạo bộ nhớ Session State chống mất dữ liệu khi chuyển tab hoặc bấm nút
    if "ket_qua_giao_an" not in st.session_state:
        st.session_state["ket_qua_giao_an"] = ""
    if "lich_su_khbd" not in st.session_state:
        st.session_state["lich_su_khbd"] = []
    if "kho_anh_trich_xuat" not in st.session_state:
        st.session_state["kho_anh_trich_xuat"] = []

    # ==================== THẺ 1: XÂY DỰNG KẾ HOẠCH BÀI DẠY AI ====================
    with tab_xay_dung:
        st.markdown("<h3 style='text-align: center; color: red;'>📖 CHỨC NĂNG XÂY DỰNG KẾ HOẠCH BÀI DẠY TỐI ƯU HÓA CAO</h3>", unsafe_allow_html=True)
        
        ten_bai = st.text_input("Tên bài dạy / Chủ đề:", key="khbd_ten_bai")
        
        col_mon, col_tg, col_lop = st.columns(3)
        with col_mon:
            mon_hoc = st.selectbox("Môn học:", ["Khoa học tự nhiên", "Toán học", "Ngữ văn", "Tiếng Anh", "Lịch sử & Địa lý", "Tin học", "Công nghệ"])
        with col_tg:
            thoi_luong = st.text_input("Thời lượng:", placeholder="Ví dụ: 3 tiết")
        with col_lop:
            lop = st.text_input("Lớp:", placeholder="Ví dụ: 6A")
            
        tich_hop_ai = st.checkbox("Tích hợp giáo dục AI (Năng lực số và AI)", value=True)
        uati_bam_sat = st.checkbox("Ưu tiên bám sát 100% nội dung tài liệu nguồn tải lên", value=True)
        
        st.markdown("**📁 Hệ thống tải lên học liệu tham khảo đa file (Hỗ trợ nạp cùng lúc nhiều tài liệu):**")
        tai_hoc_lieu = st.file_uploader(
            "Kéo thả tất cả các file tài liệu nền tảng tại đây", 
            type=["docx", "pdf", "txt"], 
            accept_multiple_files=True,
            key="hoc_lieu_uploader"
        )
        
        col_btn1, col_blank, col_btn2 = st.columns([1.8, 1.5, 1.7])
        
        # Nhập lưu ý đặc biệt từ giáo viên để gửi cho AI
        st.markdown("**💬 Yêu cầu ràng buộc khác (Để AI làm căn cứ bổ sung khi soạn bài):**")
        yeu_cau_khac = st.text_area(
            "Nhập lưu ý đặc biệt...",
            placeholder="Ví dụ: Thiết kế thêm bảng phụ lục so sánh các chất, viết rõ phương trình hóa học cân bằng, sử dụng công thức toán định dạng rõ ràng...",
            label_visibility="collapsed",
            height=100
        )
        
        # --- LOGIC XỬ LÝ KHI BẤM NÚT KHỞI TẠO ---
        with col_btn2:
            st.write("") # Tạo khoảng cách dòng
            st.write("")
            nut_chay_ai = st.button("⚡ Khởi tạo kế hoạch bài dạy bằng AI", type="primary", use_container_width=True)
            
        if nut_chay_ai:
            if not ten_bai:
                st.warning("⚠️ Vui lòng nhập Tên bài dạy trước!")
            elif not tai_hoc_lieu:
                st.warning("⚠️ Vui lòng nạp học liệu tham khảo để AI bám sát dữ liệu nguồn!")
            else:
                with st.spinner("🧠 Trợ lý AI đang nghiên cứu kỹ dữ liệu nguồn đa file và tiến hành lập tiến trình bài dạy..."):
                    from khbd_manager import extract_context_from_uploaded_files # Đảm bảo import đúng hàm
                    văn_bản_nguồn, danh_sách_ảnh = extract_context_from_uploaded_files(tai_hoc_lieu)
                    st.session_state["kho_anh_trich_xuat"] = danh_sách_ảnh

                    prompt_yeu_cau = f"""
                    Bạn là Chuyên gia viết giáo án cấp cao. Hãy soạn một Kế hoạch bài dạy cực kỳ chi tiết, đầy đủ chữ, không viết tóm tắt.
                    Tên bài: {ten_bai} | Môn: {mon_hoc} | Thời lượng: {thoi_luong} | Lớp: {lop}
                    
                    YÊU CẦU NỘI DUNG VÀ CẤU TRÚC BẮT BUỘC:
                    1. BỐ CỤC: Tuân thủ 100% cấu trúc PHỤ LỤC IV của Công văn 5512/BGDĐT:
                       I. MỤC TIÊU (1. Kiến thức, 2. Năng lực, 3. Phẩm chất)
                       II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
                       III. TIẾN TRÌNH DẠY HỌC: Thiết kế phân bổ tiến trình hợp lý cho tổng thời lượng là {thoi_luong}. Gồm đầy đủ 4 Hoạt động: Hoạt động 1: Mở đầu; Hoạt động 2: Hình thành kiến thức mới; Hoạt động 3: Luyện tập; Hoạt động 4: Vận dụng.
                       *Mỗi hoạt động BẮT BUỘC viết đầy đủ nội dung chữ cho cả 4 mục nhỏ: a) Mục tiêu; b) Nội dung; c) Sản phẩm; d) Tổ chức thực hiện (Giao nhiệm vụ -> Thực hiện -> Báo cáo thảo luận -> Kết luận nhận định).
                    
                    2. ĐỘ CHÍNH XÁC & ĐỊNH DẠNG:
                       - Bám sát hoàn toàn 100% nội dung kiến thức từ tài liệu nguồn dưới đây. Không tự chế thuật ngữ, định lý ngoài nguồn.
                       - CÔNG THỨC TOÁN/HÓA HỌC: Trình bày rõ các công thức, phương trình phản ứng cân bằng (ví dụ: viết rõ chỉ số dạng H2O, CO2, Fe2(SO4)3 hoặc biểu thức toán).
                       - BIỂU BẢNG: Phần so sánh hoặc phiếu học tập phải dùng bảng định dạng Markdown bằng ký tự '|' để hệ thống tự chuyển đổi sang Word.
                       - HÌNH ẢNH MINH HỌA: Tại các bước lý thuyết phù hợp, hãy chèn chính xác dòng chữ dòng đơn là "[Hình ảnh minh họa]" để hệ thống nhúng ảnh từ file gốc.
                    
                    3. TÍCH HỢP NĂNG LỰC SỐ VÀ AI: Thiết kế các hoạt động yêu cầu học sinh làm việc trên máy tính, khai thác học liệu số hoặc sử dụng công cụ AI trợ giúp giải bài tập.
                    4. CĂN CỨ BỔ SUNG KHÁC: {yeu_cau_khac}
                    
                    DỮ LIỆU FILE NGUỒN TÀI LIỆU THAM KHẢO:
                    {văn_bản_nguồn}
                    """
                    
                    ket_qua_ai, model_used = run_ai_prompt_safe_func(prompt_yeu_cau)
                    st.session_state["ket_qua_giao_an"] = ket_qua_ai

        with col_btn1:
            st.write("") # Tạo khoảng cách dòng
            st.write("")
            # Tạo sẵn luồng tải file Word dựa trên dữ liệu đã lưu trong Session
            from khbd_manager import export_khbd_to_docx
            docx_data = export_khbd_to_docx(st.session_state["ket_qua_giao_an"], st.session_state["kho_anh_trich_xuat"]) if st.session_state["ket_qua_giao_an"] else b""
            st.download_button(
                label="📥 Tải file Word (.docx) chuẩn biểu bảng",
                data=docx_data,
                file_name=f"KHBD_{ten_bai.replace(' ', '_') if ten_bai else 'BGD_5512'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                disabled=(st.session_state["ket_qua_giao_an"] == ""),
                use_container_width=True
            )

        # Hiển thị nội dung bài soạn
        st.markdown("**📊 Nội dung bài soạn hiển thị xem trước:**")
        with st.container(border=True):
            if st.session_state["ket_qua_giao_an"]:
                st.markdown(st.session_state["ket_qua_giao_an"])
                if st.button("📥 Lưu vào Thư viện hệ thống"):
                    if ten_bai:
                        st.session_state["lich_su_khbd"].append({"Tên bài": ten_bai, "Môn": mon_hoc, "Lớp": lop, "Nội dung": st.session_state["ket_qua_giao_an"]})
                        st.success("✅ Đã lưu giáo án thành công!")
            else:
                st.caption("Bài soạn sau khi khởi tạo bằng AI sẽ hiển thị tại đây...")

    # ==================== THÈ 2: LƯU TRỮ KẾ HOẠCH BÀI DẠY ĐÃ XD ====================
    with tab_luu_khbd:
        st.markdown("### 🗄️ THƯ VIỆN LƯU TRỮ KẾ HOẠCH BÀI DẠY ĐÃ XÂY DỰNG")
        if not st.session_state["lich_su_khbd"]:
            st.info("Chưa có bài soạn nào được lưu trong phiên này.")
        else:
            for idx, item in enumerate(st.session_state["lich_su_khbd"]):
                with st.expander(f"📚 {idx+1}. {item['Tên bài']} - Lớp {item['Lớp']}"):
                    st.markdown(item["Nội dung"])
