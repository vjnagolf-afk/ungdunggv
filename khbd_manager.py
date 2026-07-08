import streamlit as st
import docx  
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
from pypdf import PdfReader

# --- HÀM TRÍCH XUẤT VĂN BẢN VÀ LỌC ẢNH TRÙNG LẶP ---
def extract_context_from_uploaded_files(uploaded_files):
    combined_text = ""
    extracted_images = [] 
    seen_image_sizes = set()
    for file in uploaded_files:
        try:
            if file.name.endswith('.docx'):
                doc = docx.Document(file)
                for paragraph in doc.paragraphs:
                    if paragraph.text: combined_text += paragraph.text + "\n"
                for table in doc.tables:
                    for row in table.rows:
                        text_row = [cell.text for cell in row.cells]
                        combined_text += " | ".join(text_row) + "\n"
                for rel in doc.part.relations.values():
                    if "image" in rel.target_ref:
                        img_blob = rel.target_part.blob
                        if len(img_blob) not in seen_image_sizes:
                            seen_image_sizes.add(len(img_blob))
                            extracted_images.append(img_blob)
            elif file.name.endswith('.pdf'):
                reader = PdfReader(file)
                for page in reader.pages:
                    combined_text += (page.extract_text() or "") + "\n"
                    for img_file_object in page.images:
                        img_blob = img_file_object.data
                        if len(img_blob) not in seen_image_sizes:
                            seen_image_sizes.add(len(img_blob))
                            extracted_images.append(img_blob)
            elif file.name.endswith('.txt'):
                combined_text += file.read().decode("utf-8") + "\n"
        except Exception as e:
            st.error(f"Lỗi khi xử lý file {file.name}: {str(e)}")
    return combined_text, extracted_images

# --- HÀM SET KHOẢNG CÁCH ĐOẠN 3 - 4.5 PT CHUẨN UX ---
def set_paragraph_spacing(paragraph, before_pt=3.0, after_pt=4.5):
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(before_pt * 20)))
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    p_pr.append(spacing)

# --- HÀM CHUYỂN ĐỔI BIỂU THỨC THÀNH CÔNG THỨC TOÁN ĐỨNG ---
def add_math_expression(doc, text_line):
    if "[frac:" in text_line or "[root:" in text_line:
        p = doc.add_paragraph()
        set_paragraph_spacing(p)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        m_table = doc.add_table(rows=1, cols=3)
        m_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        frac_match = re.search(r'\[frac:\s*([^/]+)/([^\]]+)\]', text_line)
        root_match = re.search(r'\[root:\s*([^|]+)\|([^\]]+)\]', text_line)
        if frac_match:
            tu, mau = frac_match.group(1).strip(), frac_match.group(2).strip()
            cell = m_table.cell(0, 0)
            cell.text = f"{tu}\n---\n{mau}"
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
        if root_match:
            bac, b_thuc = root_match.group(1).strip(), root_match.group(2).strip()
            cell = m_table.cell(0, 1)
            cell.text = f"^{bac}√({b_thuc})"
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
    else:
        p = doc.add_paragraph()
        set_paragraph_spacing(p)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        run = p.add_run(text_line)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)

# --- HÀM XUẤT FILE WORD TỔNG HỢP NÂNG CAO CẬP NHẬT TIẾT VÀ HOẠT ĐỘNG ---
def export_khbd_to_docx(markdown_content, images_list):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)

    MAU_DO = RGBColor(255, 0, 0)
    MAU_XANH_DUONG = RGBColor(0, 51, 153)
    MAU_DEN = RGBColor(0, 0, 0)

    lines = markdown_content.split('\n')
    in_table = False
    table_data = []
    used_img_idx = 0

    for line in lines:
        clean_line = line.strip().replace('**', '').replace('###', '').replace('##', '').replace('#', '')
        
        # LỌC SẠCH DẤU GẠCH NGANG THỪA ở đầu dòng đối với các đề mục nhỏ cụ thể
        if re.match(r'^-\s*((\d+\.)|([a-d]\)))', clean_line):
            clean_line = re.sub(r'^-\s*', '', clean_line)

        if line.strip().startswith('|') and line.strip().endswith('|'):
            if '---|' in line or ':---|' in line: continue
            in_table = True
            cells = [c.strip().replace('**', '') for c in line.split('|')[1:-1]]
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
                                cell = word_table.cell(r_idx, c_idx)
                                cell.text = val
                                for para in cell.paragraphs:
                                    set_paragraph_spacing(para, 2.0, 3.0)
                                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                    for r in para.runs:
                                        r.font.name = 'Times New Roman'
                                        r.font.size = Pt(14)
                                        r.font.color.rgb = MAU_DEN
                in_table = False
                table_data = []

        if not clean_line: continue

        if "[Hình ảnh minh họa]" in line and images_list:
            if used_img_idx < len(images_list):
                try:
                    p = doc.add_paragraph()
                    set_paragraph_spacing(p)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    img_stream = io.BytesIO(images_list[used_img_idx])
                    doc.add_picture(img_stream, width=Inches(4.5))
                    used_img_idx += 1
                    continue
                except: pass

        # 1. ĐỊNH DẠNG TIÊU ĐỀ CHÍNH HOẶC "TIẾT ... " (IN HOA ĐẬM, CĂN GIỮA, CHỮ ĐỎ)
        if any(x in clean_line.upper() for x in ["MÔN HỌC:", "LỚP:", "BÀI:", "KẾ HOẠCH BÀI DẠY"]) or re.match(r'^TIẾT\s+\d+', clean_line.upper()):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 6.0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_line.upper())
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.color.rgb = MAU_DO

        # 2. ĐỊNH DẠNG "HOẠT ĐỘNG X" LỚN (IN ĐẬM, CĂN GIỮA, CHỮ XANH DƯƠNG)
        elif re.match(r'^HOẠT\s+ĐỘNG\s+\d+($|[^.\d])', clean_line.upper()):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(clean_line.upper())
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.color.rgb = MAU_XANH_DUONG

        # 3. ĐỊNH DẠNG "HOẠT ĐỘNG X.Y" NHỎ (IN ĐẬM, CĂN TRÁI, CHỮ XANH HOẶC ĐEN)
        elif re.match(r'^HOẠT\s+ĐỘNG\s+\d+\.\d+', clean_line.upper()):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 3.5, 4.0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(clean_line.upper())
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.color.rgb = MAU_XANH_DUONG

        # 4. ĐỊNH DẠNG ĐỀ MỤC LỚN I, II, III... (IN ĐẬM, CĂN TRÁI, CHỮ XANH DƯƠNG)
        elif re.match(r'^(I|II|III|IV|V|VI)\.', clean_line) or re.match(r'^\d+\.', clean_line):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(clean_line)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.color.rgb = MAU_XANH_DUONG

        # 5. CÁC NỘI DUNG THÔNG THƯỜNG KHÁC (CĂN ĐỀU CẢ 2 BÊN)
        else:
            add_math_expression(doc, clean_line)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
# --- GIAO DIỆN CHÍNH CỦA PHÂN HỆ ---
def render_khbd_section(run_ai_prompt_safe_func):
    
    tab_xay_dung, tab_luu_khbd = st.tabs(["... XÂY DỰNG KẾ HOẠCH BÀI DẠY AI", "🗄️ LƯU KHBD ĐÃ XD"])
    
    if "ket_qua_giao_an" not in st.session_state: st.session_state["ket_qua_giao_an"] = ""
    if "lich_su_khbd" not in st.session_state: st.session_state["lich_su_khbd"] = []
    if "kho_anh_trich_xuat" not in st.session_state: st.session_state["kho_anh_trich_xuat"] = []

    # ==================== THỂ 1: XÂY DỰNG KẾ HOẠCH BÀI DẠY AI ====================
    with tab_xay_dung:
        st.markdown("<h3 style='text-align: center; color: red;'>📖 CHỨC NĂNG XÂY DỰNG KẾ HOẠCH BÀI DẠY TỐI ƯU HÓA CAO</h3>", unsafe_allow_html=True)
        
        ten_bai = st.text_input("Tên bài dạy / Chủ đề:", key="khbd_ten_bai")
        
        col_mon, col_tg, col_lop = st.columns(3)
        with col_mon:
            mon_hoc = st.selectbox("Môn học:", ["Khoa học tự nhiên", "Toán học", "Ngữ văn", "Tiếng Anh", "Lịch sử & Địa lý", "Tin học", "Công nghệ"])
        with col_tg:
            thoi_luong = st.text_input("Thời lượng:", placeholder="Ví dụ: 2 Tiết")
        with col_lop:
            lop = st.text_input("Lớp:", placeholder="Ví dụ: 7A")
            
        tich_hop_ai = st.checkbox("Tích hợp giáo dục AI (Năng lực số và AI)", value=True)
        uati_bam_sat = st.checkbox("Ưu tiên bám sát 100% nội dung tài liệu nguồn tải lên", value=True)
        
        st.markdown("**📁 Hệ thống tải lên học liệu tham khảo đa file (Hỗ trợ nạp cùng lúc nhiều tài liệu):**")
        tai_hoc_lieu = st.file_uploader("Kéo thả tất cả các file tài liệu tại đây", type=["docx", "pdf", "txt"], accept_multiple_files=True, key="hoc_lieu_uploader")
        
        col_btn1, col_blank, col_btn2 = st.columns([2.0, 1.3, 1.7])
        
        st.markdown("**💬 Yêu cầu ràng buộc khác (Để AI làm căn cứ bổ sung khi soạn bài):**")
        yeu_cau_khac = st.text_area("Nhập lưu ý...", placeholder="Ví dụ: Giữ nguyên bảng số liệu thực hành tốc độ chuyển động trong tài liệu nguồn...", label_visibility="collapsed", height=100)
        
        with col_btn2:
            st.write(""); st.write("")
            nut_chay_ai = st.button("⚡ Khởi tạo kế hoạch bài dạy bằng AI", type="primary", use_container_width=True)
            
        if nut_chay_ai:
            if not ten_bai:
                st.warning("⚠️ Vui lòng nhập Tên bài dạy trước!")
            elif not tai_hoc_lieu:
                st.warning("⚠️ Vui lòng nạp học liệu tham khảo để AI bám sát dữ liệu nguồn!")
            else:
                with st.spinner("🧠 Trợ lý AI đang nghiên cứu kỹ dữ liệu nguồn đa file và tiến hành lập tiến trình bài dạy..."):
                    from khbd_manager import extract_context_from_uploaded_files
                    văn_bản_nguồn, danh_sách_ảnh = extract_context_from_uploaded_files(tai_hoc_lieu)
                    st.session_state["kho_anh_trich_xuat"] = danh_sách_ảnh

                    prompt_yeu_cau = f"""
                    Bạn là Chuyên gia viết giáo án cấp cao bậc THCS/THPT tại Việt Nam. Hãy soạn một Kế hoạch bài dạy cực kỳ chi tiết, đầy đủ chữ theo đúng yêu cầu cấu trúc và định dạng sau.
                    
                    TIÊU ĐỀ BÀI HỌC (Viết in hoa ở đầu giáo án):
                    MÔN HỌC: {mon_hoc.upper()}
                    LỚP: {lop.upper()}
                    BÀI: {ten_bai.upper()} ({thoi_luong})
                    
                    YÊU CẦU CẤU TRÚC PHỤ LỤC IV CÔNG VĂN 5512/BGDĐT:
                    I. MỤC TIÊU (Bắt buộc chia nhỏ thành đúng 4 đề mục con sau, KHÔNG ĐƯỢC THÊM dấu gạch ngang '-' ở trước số thứ tự):
                       1. Kiến thức
                       2. Năng lực (Bao gồm Năng lực chung và Năng lực đặc thù của môn học)
                       3. Năng lực số và AI (Thiết kế mục tiêu học sinh ứng dụng thiết bị số, biết viết prompt cho chatbot AI làm trợ lý học tập)
                       4. Phẩm chất
                    II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
                    III. TIẾN TRÌNH DẠY HỌC (Phân bổ tiến trình hợp lý cho tổng số {thoi_luong}. Ví dụ nếu có nhiều tiết, ghi rõ dòng độc lập 'TIẾT 1:', 'TIẾT 2:',... ở đầu phần phân bổ).
                       - Tiến trình gồm 4 Hoạt động lớn: 'HOẠT ĐỘNG 1: MỞ ĐẦU', 'HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI', 'HOẠT ĐỘNG 3: LUYỆN TẬP', 'HOẠT ĐỘNG 4: VẬN DỤNG'.
                       - Nếu trong hoạt động lớn có chia các hoạt động nhỏ, đặt tên dạng: 'HOẠT ĐỘNG 1.1: ...', 'HOẠT ĐỘNG 1.2: ...'
                       - Mỗi hoạt động (hoặc hoạt động nhỏ) bắt buộc trình bày đủ 4 mục sau và KHÔNG ĐƯỢC THÊM dấu gạch ngang '-' ở trước chữ:
                         a) Mục tiêu:
                         b) Nội dung:
                         c) Sản phẩm:
                         d) Tổ chức thực hiện:
                    
                    YÊU CẦU ĐỊNH DẠNG CÔNG THỨC TOÁN, BIỂU BẢNG VÀ CÚ PHÁP:
                    - BẢO TỒN NGUYÊN VẸN BIỂU BẢNG: Đối với bất kỳ bảng số liệu, bảng so sánh, hoặc biểu bảng nào xuất hiện trong tài liệu nguồn/SGK được cung cấp ở bên dưới, bạn BẮT BUỘC phải trích xuất và giữ nguyên cấu trúc (số hàng, số cột) và sao chép 100% đầy đủ toàn bộ nội dung chữ/số liệu bên trong các ô. KHÔNG ĐƯỢC TỰ Ý TÓM TẮT, lược bỏ, hoặc viết chung chung nội dung bảng. Thiết kế chính xác dạng bảng Markdown bằng ký tự '|'.
                    - BỎ TOÀN BỘ các ký tự dấu sao kép '**' ở đầu và cuối các từ hoặc các mục.
                    - Toàn bộ danh sách nội dung thông thường dùng duy nhất ký tự gạch ngang '-' ở đầu dòng.
                    - CÔNG THỨC TOÁN HỌC KHỐI ĐỨNG: Khi xuất hiện công thức phân số hoặc căn thức phức tạp, bạn BẮT BUỘC phải bọc trong thẻ định dạng sau để hệ thống biên dịch sang Word dạng khối đứng:
                      + Phân số: bọc dạng [frac: tử_số/mẫu_số] (Ví dụ: [frac: 25/30])
                      + Căn thức: bọc dạng [root: bậc_căn|biểu_thức_dưới_căn] (Ví dụ: [root: 5|(3x - 5y)])
                    - HÌNH ẢNH: Tại vị trí lý thuyết phù hợp, ghi một dòng độc lập là "[Hình ảnh minh họa]".
                    
                    CĂN CỨ BỔ SUNG KHÁC: {yeu_cau_khac}
                    
                    DỮ LIỆU FILE NGUỒN TÀI LIỆU THAM KHẢO (Đọc thật kỹ các phần biểu bảng và dữ liệu để sao chép):
                    {văn_bản_nguồn}
                    """
                    ket_qua_ai, model_used = run_ai_prompt_safe_func(prompt_yeu_cau)
                    st.session_state["ket_qua_giao_an"] = ket_qua_ai

        with col_btn1:
            st.write(""); st.write("")
            from khbd_manager import export_khbd_to_docx
            docx_data = export_khbd_to_docx(st.session_state["ket_qua_giao_an"], st.session_state["kho_anh_trich_xuat"]) if st.session_state["ket_qua_giao_an"] else b""
            st.download_button(
                label="📥 Tải file Word (.docx) chuẩn về máy",
                data=docx_data,
                file_name=f"KHBD_{ten_bai.replace(' ', '_') if ten_bai else 'BGD_5512'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                disabled=(st.session_state["ket_qua_giao_an"] == ""),
                use_container_width=True
            )

        st.markdown("**📊 Nội dung bài soạn hiển thị xem trước:**")
        with st.container(border=True):
            if st.session_state["ket_qua_giao_an"]:
                st.markdown(st.session_state["ket_qua_giao_an"])
                if st.button("📥 Lưu vào Thư viện hệ thống", use_container_width=True):
                    if ten_bai:
                        st.session_state["lich_su_khbd"].append({"Tên bài": ten_bai, "Môn": mon_hoc, "Lớp": lop, "Nội dung": st.session_state["ket_qua_giao_an"], "Kho_anh": st.session_state["kho_anh_trich_xuat"]})
                        st.success("✅ Đã lưu giáo án vào Thư viện thành công!")
            else:
                st.caption("Bài soạn sau khi khởi tạo bằng AI sẽ hiển thị tại đây...")

    # ==================== THÈ 2: LƯU TRỮ KẾ HOẠCH BÀI DẠY ĐÃ XD ====================
    with tab_luu_khbd:
        st.markdown("### 🗄️ THƯ VIỆN LƯU TRỮ KẾ HOẠCH BÀI DẠY ĐÃ XÂY DỰNG")
        if not st.session_state["lich_su_khbd"]:
            st.info("Chưa có bài soạn nào được lưu trong phiên này.")
        else:
            for idx, item in enumerate(st.session_state["lich_su_khbd"]):
                col_exp, col_del = st.columns([0.88, 0.12])
                with col_exp:
                    with st.expander(f"📚 {idx+1}. {item['Tên bài']} - Lớp {item['Lớp']}"):
                        st.markdown(item["Nội dung"])
                        from khbd_manager import export_khbd_to_docx
                        saved_docx = export_khbd_to_docx(item["Nội dung"], item.get("Kho_anh", []))
                        st.download_button(
                            label="📥 Tải lại bản Word (.docx)",
                            data=saved_docx,
                            file_name=f"Luu_tru_{item['Tên bài'].replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_saved_docx_{idx}"
                        )
                with col_del:
                    st.write("") 
                    if st.button("🗑️ Xóa bài", key=f"del_khbd_{idx}", use_container_width=True, type="secondary"):
                        st.session_state["lich_su_khbd"].pop(idx)
                        st.success(f"Đã xóa bài số {idx+1}!")
                        st.rerun()
