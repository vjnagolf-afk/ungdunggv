# khbd_manager.py - ĐOẠN 1: CẤU HÌNH & TRÍCH XUẤT TÀI LIỆU (BẢN VÁ TOÀN DIỆN)
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
import gspread
from datetime import datetime

# Nhúng bộ biên dịch toán và đồ thị thông minh để dùng cho bài soạn
from math_compiler import process_runs_with_math, generate_plot_stream

# ================= CẤU HÌNH GOOGLE SHEETS =================
SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY' 

def get_khbd_sheet():
    try:
        creds_dict = None
        priority_keys = ["gspread_credentials", "GSPREAD_CREDENTIALS", "google_sheet_creds", "google_creds", "GOOGLE_KEY"]
        for key in priority_keys:
            if key in st.secrets:
                creds_dict = st.secrets[key]
                break
                
        if creds_dict is None:
            for key in st.secrets.keys():
                node = st.secrets[key]
                if hasattr(node, "get") or isinstance(node, dict):
                    if node.get("type") == "service_account" or "private_key" in node:
                        creds_dict = node
                        break
                        
        if creds_dict is None:
            return None
            
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_key(SHEET_ID)
        return sh.worksheet("KHBD")
    except:
        return None

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

def set_paragraph_spacing(paragraph, before_pt=3.0, after_pt=4.5):
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(before_pt * 20)))
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    p_pr.append(spacing)
# khbd_manager.py - ĐOẠN 2: THUẬT TOÁN KẾT XUẤT GIÁO ÁN CHUẨN WORD (VÁ LỖI ĐỀ MỤC & TOÁN HỌC)
# khbd_manager.py - ĐOẠN 2: THUẬT TOÁN KẾT XUẤT GIÁO ÁN CHUẨN ĐỊNH DẠNG MÀU SẮC YÊU CẦU
# khbd_manager.py - ĐOẠN 2: THUẬT TOÁN KẾT XUẤT GIÁO ÁN CHUẨN ĐỊNH DẠNG MÀU SẮC YÊU CẦU
# khbd_manager.py - ĐOẠN 2: THUẬT TOÁN KẾT XUẤT GIÁO ÁN VĂN BẢN XUÔI (TUYỆT ĐỐI KHÔNG DÙNG BẢNG)
def export_khbd_to_docx(markdown_content, images_list):
    # 🚀 1. SƠ CHẾ VĂN BẢN TOÀN CỤC: Khử sạch thẻ bảng và làm sạch dấu tiêu đề Markdown
    markdown_content = re.sub(r'(?m)^#+\s*', '', markdown_content) # Xóa dấu #
    markdown_content = markdown_content.replace("<br>", "\n").replace("<br/>", "\n") # Ép thẻ dòng HTML thành xuống dòng Word
    
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)

    MAU_DO = RGBColor(255, 0, 0)
    MAU_XANH_DUONG = RGBColor(0, 51, 153)
    MAU_DEN = RGBColor(0, 0, 0)

    # Khóa cứng phôi chữ in hành chính Times New Roman 14pt
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)

    lines = markdown_content.split('\n')
    passed_intro = False

    for line in lines:
        cleaned_line = line.strip().replace('**', '')
        if not cleaned_line: 
            continue

        # 🚀 2. THUẬT TOÁN BÓC TÁCH KHỬ TOÀN BỘ KÝ TỰ BẢNG PHẲNG CỦA AI
        # Nếu AI cố tình sinh dòng dạng cột '| Thời gian | Hoạt động |', hệ thống tự bẻ khóa đưa về text thường
        if cleaned_line.startswith('|'):
            cleaned_line = cleaned_line.replace('|', '  ').strip()
            # Bỏ qua các đường kẻ khung ngăn cách bảng của Markdown
            if all(c in '-: ' for c in cleaned_line): 
                continue

        # Loại bỏ lời dẫn thừa đầu trang
        if not passed_intro:
            if any(x in cleaned_line.upper() for x in ["KẾ HOẠCH BÀI DẠY", "MÔN HỌC:", "LỚP:", "BÀI:", "TIẾT ", "I. MỤC TIÊU"]):
                passed_intro = True
            else:
                continue

        # Giấu hoàn toàn dòng tên bộ sách giáo khoa
        if any(x in cleaned_line.upper() for x in ["BỘ SÁCH:", "CÁNH DIỀU", "KẾT NỐI TRI THỨC", "CHÂN TRỜI SÁNG TẠO"]):
            continue

        # 🚀 3. TƯỚC BỎ DẤU GẠCH NGANG TRƯỚC CÁC TIÊU ĐỀ ĐỀ MỤC CHÍNH CHUẨN VĂN BẢN
        if cleaned_line.startswith('-'):
            sub_text = re.sub(r'^-+\s*', '', cleaned_line).strip()
            if (re.match(r'^(I|II|III|IV|V|VI|VII)\.', sub_text) or 
                re.match(r'^[A-D]\.', sub_text) or
                re.match(r'^\d+\.', sub_text) or 
                re.match(r'^[a-d]\)', sub_text) or
                any(x in sub_text.upper() for x in ["MÔN HỌC:", "LỚP:", "BÀI:", "KẾ HOẠCH BÀI DẠY", "THỜI LƯỢNG:"])):
                cleaned_line = sub_text

        # Chuẩn hóa định dạng dấu đầu dòng văn bản thường
        if cleaned_line.startswith('*'):
            cleaned_line = re.sub(r'^\*+\s*', '- ', cleaned_line)
        elif cleaned_line.startswith('-') and not re.match(r'^-\s*[a-zA-Z0-9I]', cleaned_line):
            cleaned_line = re.sub(r'^-+\s*', '- ', cleaned_line)

        # Nhận diện hàm số để vẽ đồ thị tự động
        if '[GRAPH:' in cleaned_line:
            match = re.search(r'\[GRAPH:\s*(.+?)\]', cleaned_line)
            if match:
                eq = match.group(1)
                img_stream = generate_plot_stream(eq)
                doc.add_picture(img_stream, width=Inches(4.5))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

        # Định dạng khối tiêu đề lớn căn giữa (IN HOA, ĐỎ/XANH, 14pt)
        if any(x in cleaned_line.upper() for x in ["MÔN HỌC:", "LỚP:", "BÀI:", "KẾ HOẠCH BÀI DẠY", "THỜI LƯỢNG:"]) or re.match(r'^TIẾT\s+\d+', cleaned_line.upper()):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(cleaned_line.upper())
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.color.rgb = MAU_DO if "KẾ HOẠCH BÀI DẠY" in cleaned_line.upper() else MAU_XANH_DUONG
            continue

        # 🚀 4. ĐỊNH DẠNG ĐỀ MỤC LỚN (I, II, A, B...) -> IN ĐẬM, MÀU XANH DƯƠNG 14PT
        if re.match(r'^(I|II|III|IV|V|VI|VII)\.', cleaned_line) or re.match(r'^[A-D]\.\s+', cleaned_line):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            process_runs_with_math(p, cleaned_line)
            for r in p.runs:
                r.bold = True
                r.font.name = 'Times New Roman'
                r.font.size = Pt(14)
                r.font.color.rgb = MAU_XANH_DUONG
            continue

        # 🚀 5. ĐỊNH DẠNG TIỂU MỤC SỐ (1. Về kiến thức, 2. Về năng lực...) -> IN ĐẬM, MÀU ĐỎ 14PT
        if re.match(r'^\d+\.\s+Về\s+(kiến thức|năng lực|phẩm chất)', cleaned_line) or re.match(r'^\d+\.\s+Về\s+Mục\s+tiêu', cleaned_line):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            process_runs_with_math(p, cleaned_line)
            for r in p.runs:
                r.bold = True
                r.font.name = 'Times New Roman'
                r.font.size = Pt(14)
                r.font.color.rgb = MAU_DO
            continue
        # 🚀 6. THUẬT TOÁN PHÂN TÁCH CÔ LẬP KHỐI IN ĐẬM ĐẦU DÒNG CHỮ ĐEN CHUẨN XUÔI DÒNG
        # Sửa lại Regex phân cấp đúng 2 nhóm ngoặc đơn: Nhóm 1 (Từ khóa), Nhóm 2 (Nội dung sau dấu :)
        keyword_match = re.match(r'^((?:-?\s*Mục tiêu:|-?\s*Nội dung:|-?\s*Sản phẩm:|-?\s*Tổ chức thực hiện:|-?\s*Bước\s+\d+:|Thành phần|Thời gian|[a-d]\)\s+Năng lực\s+(?:chung|đặc thù|số và AI):))(.*)', cleaned_line, re.IGNORECASE)
        
        if keyword_match:
            keyword = keyword_match.group(1).strip()
            rest_of_text = keyword_match.group(2).strip() # 🚀 ĐÃ SỬA THÀNH GROUP(2) CHUẨN XÁC
            
            # Làm sạch dấu gạch ngang của từ khóa khi đưa vào phôi Word
            keyword_clean = re.sub(r'^-+\s*', '', keyword)
            
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 3.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if keyword.startswith('-'):
                p.paragraph_format.left_indent = Inches(0.25)
            
            # Run 1: Ghi riêng từ khóa đầu mục -> ÉP IN ĐẬM CHỮ ĐEN
            run_key = p.add_run(keyword_clean + " ")
            run_key.bold = True
            run_key.font.name = 'Times New Roman'
            run_key.font.size = Pt(14)
            run_key.font.color.rgb = MAU_DEN
            
            # Run 2: Ghi phần nội dung diễn giải phía sau -> CHỮ THƯỜNG MÀU ĐEN (Nhúng quét toán)
            process_runs_with_math(p, rest_of_text)
            for r in p.runs[1:]:
                r.bold = False
                r.font.name = 'Times New Roman'
                r.font.size = Pt(14)
                r.font.color.rgb = MAU_DEN
            continue
        # Đoạn văn giáo án nội dung thường phẳng (Times New Roman 14pt thường)
        p = doc.add_paragraph()
        set_paragraph_spacing(p)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if cleaned_line.startswith('-'):
            p.paragraph_format.left_indent = Inches(0.25)
            
        process_runs_with_math(p, cleaned_line)
        for r in p.runs:
            r.font.name = 'Times New Roman'
            if not r.font.size:
                r.font.size = Pt(14)
            r.font.color.rgb = MAU_DEN

    # Xuất luồng nhị phân an toàn để trả file về nút download_button
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# khbd_manager.py - ĐOẠN 3: GIAO DIỆN STREAMLIT PHÂN HỆ KHBD HOÀN CHỈNH TÍCH HỢP Ô CHECKBOX
def render_khbd_section(run_ai_prompt_safe_func):
    st.markdown("<h3 style='text-align: center; color: blue;'>🧠 TRỢ LÝ THIẾT KẾ KẾ HOẠCH BÀI DẠY (KHBD) AI PHÁT TRIỂN NĂNG LỰC</h3>", unsafe_allow_html=True)
    
    tab_thiet_ke, tab_thu_vien = st.tabs(["📝 THIẾT KẾ KHBD TỰ ĐỘNG", "🗄️ THƯ VIỆN BÀI SOẠN"])
    
    if "ket_qua_khbd" not in st.session_state: st.session_state["ket_qua_khbd"] = ""
    if "lich_su_khbd" not in st.session_state: st.session_state["lich_su_khbd"] = []
    if "images_khbd" not in st.session_state: st.session_state["images_khbd"] = []

    with tab_thiet_ke:
        st.write("Nhập thông tin bài học và tải lên tài liệu tham khảo để AI lập tiến trình dạy học Công văn 5512.")
        
        ten_bai = st.text_input("Tên bài học / Chủ đề bài dạy:", placeholder="Ví dụ: Bài 4: Tốc độ chuyển động - Khoa học tự nhiên 7")
        
        col_lop, col_bo = st.columns(2)
        with col_lop:
            lop_khbd = st.text_input("Lớp dạy:", value="Lớp 7")
        with col_bo:
            bo_sach = st.selectbox("Bộ sách giáo khoa:", ["Kết nối tri thức với cuộc sống", "Cánh Diều", "Chân trời sáng tạo", "Chương trình GDPT 2018"])

        col_tg, col_loai = st.columns(2)
        with col_tg:
            thoi_luong = st.text_input("Thời lượng bài dạy (Tiết):", value="2 tiết")
        with col_loai:
            kieu_khbd = st.selectbox("Mẫu cấu trúc thiết kế:", ["Chuẩn Công văn 5512 (Đầy đủ 4 hoạt động)", "Rút gọn (Tiết kiệm thời gian)", "Giáo án hoạt động trải nghiệm/STEM"])

        st.markdown("**Tải lên tài liệu tham khảo thô (Đề cương, Sách, file nội dung bài học nếu có):**")
        files_tailieu = st.file_uploader("Chọn tệp (.docx, .pdf, .txt):", type=["docx", "pdf"], accept_multiple_files=True, key="khbd_files_upload")

        context_data = ""
        if files_tailieu:
            context_data, st.session_state["images_khbd"] = extract_context_from_uploaded_files(files_tailieu)
            st.success(f"📊 Đã nạp thành công văn bản tham khảo và trích xuất được {len(st.session_state['images_khbd'])} hình ảnh!")

        col_chk1, col_chk2 = st.columns(2)
        with col_chk1:
            chk_ai_digital = st.checkbox("Tích hợp năng lực số và AI", value=True)
        with col_chk2:
            chk_strict_file = st.checkbox("Bám sát 100% file tài liệu tải lên", value=False)

        yeu_cau_rieng = st.text_area("Yêu cầu sư phạm bổ sung (Tùy chọn):", placeholder="Ví dụ: Thiết kế thêm một trò chơi khởi động sôi nổi; lồng ghép công thức toán học chi tiết...")

        col_btn1, col_blank, col_btn2 = st.columns([2.2, 1.0, 1.8])
        
        with col_btn2:
            st.write(""); st.write("")
            nut_tao_khbd = st.button("⚡ Tiến hành thiết kế bài dạy bằng AI", type="primary", use_container_width=True)

        if nut_tao_khbd:
            if not ten_bai:
                st.warning("⚠️ Vui lòng điền tên bài học hoặc chủ đề giảng dạy!")
            else:
                with st.spinner("🧠 AI đang phân tích dữ liệu, bám sát Công văn 5512 để thiết kế tiến trình..."):
                    prompt_requirements = ""
                    if chk_ai_digital:
                        prompt_requirements += """
                        - TẠI MỤC I.2 (VỀ NĂNG LỰC): Bắt buộc phải bổ sung thêm riêng một tiểu mục: 'c) Năng lực số và AI' (Mô tả chi tiết tiêu chí học sinh biết khai thác tài nguyên số hoặc ứng dụng AI gì để thực hành, học tập trong bài soạn này).
                        - TẠI TIẾN TRÌNH CÁC HOẠT ĐỘNG DẠY HỌC: Thiết kế lồng ghép hoạt động giao nhiệm vụ yêu cầu học sinh sử dụng ứng dụng mô phỏng hoặc công cụ công nghệ số.
                        """
                    if chk_strict_file:
                        prompt_requirements += f"\n- QUY ĐỊNH BẮT BUỘC: Khai thác và bám sát chính xác 100% kiến thức từ tệp tài liệu tham khảo thô sau đây: {context_data[:3000]}"

                    prompt_khbd = f"""
                    Bạn là Chuyên gia Phương pháp dạy học cao cấp tại Việt Nam. Hãy soạn một Kế hoạch bài dạy (KHBD) chi tiết theo Công văn 5512 cho bài học:
                    - Tên bài: {ten_bai}
                    - Lớp: {lop_khbd} - Thuộc bộ sách: {bo_sach}
                    - Thời lượng: {thoi_luong}
                    - Kiểu cấu trúc mẫu: {kieu_khbd}
                    
                    YÊU CẦU ĐỊNH HƯỚNG SƯ PHẠM QUY ĐỊNH (THI SÁT TUYỆT ĐỐI):
                    {prompt_requirements}
                    - Yêu cầu bổ sung riêng: {yeu_cau_rieng if yeu_cau_rieng else "Không có"}
                    
                    QUY ĐỊNH ĐỊNH DẠNG CÔNG THỨC TOÁN HỌC & ĐỒ THỊ (Nghiêm ngặt):
                    - Tuyệt đối không viết chữ toán dạng phẳng dính văn bản thường. Tất cả công thức như v = s/t, các phân số chồng tầng, đại lượng vật lý bắt buộc phải đặt cô lập gọn gàng trong cặp dấu đô-la $...$ độc lập (Ví dụ: $v = \\frac{{s}}{{t}}$, $1 \\text{{ m/s}} = 3.6 \\text{{ km/h}}$).
                    - Nếu có đồ thị hàm số, xuất chuỗi dạng [GRAPH: tên_hàm] (Ví dụ: [GRAPH: 3*x]) để Word tự động vẽ hình.
                    """
                    ket_qua, model_thuc_te = run_ai_prompt_safe_func(prompt_khbd)
                    st.session_state["ket_qua_khbd"] = ket_qua

        if st.session_state["ket_qua_khbd"]:
            st.markdown("### 📋 Giáo án bài dạy xem trước:")
            st.markdown(st.session_state["ket_qua_khbd"])

        with col_btn1:
            st.write(""); st.write("")
            title_file_khbd = f"KHBD_{lop_khbd}_{ten_bai[:20].replace(' ', '_')}" if ten_bai else "Ke_Hoach_Bai_Day"
            
            if st.session_state["ket_qua_khbd"]:
                docx_data_khbd = export_khbd_to_docx(st.session_state["ket_qua_khbd"], st.session_state["images_khbd"])
                is_disabled_khbd = False
                
                sheet_khbd = get_khbd_sheet()
                if sheet_khbd is not None:
                    try:
                        sheet_khbd.append_row([ten_bai, lop_khbd, bo_sach, thoi_luong, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                    except: pass
            else:
                docx_data_khbd = b""
                is_disabled_khbd = True

            st.download_button(
                label="📥 Tải tệp Giáo án Word (.docx) bản chuẩn hành chính",
                data=docx_data_khbd,
                file_name=f"{title_file_khbd}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                disabled=is_disabled_khbd,
                use_container_width=True
            )

    with tab_thu_vien:
        st.write("🗄️ Quản lý kho lưu trữ giáo án trực tuyến:")
        st.markdown(f"🔗 [Bấm vào đây để mở trực tiếp Google Sheets quản lý bài soạn](https://google.com{SHEET_ID})")
