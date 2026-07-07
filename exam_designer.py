# exam_designer.py
import streamlit as st
import io
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader

def read_uploaded_docx(uploaded_file):
    try:
        doc = Document(uploaded_file)
        return "\n".join([para.text for para in doc.paragraphs])
    except: return "Lỗi đọc file Word"

def read_uploaded_pdf(uploaded_file):
    try:
        reader = PdfReader(uploaded_file)
        return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
    except: return "Lỗi đọc file PDF"

def export_to_docx_vietnam_standard(text_content, title_name, school_name="TRƯỜNG THCS NGUYỄN CHÍ THANH", group_name="TỔ KHOA HỌC TỰ NHIÊN - GDTC"):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    
    admin_table = doc.add_table(rows=1, cols=2)
    admin_table.autofit = False
    admin_table.columns[0].width = Inches(3.2)
    admin_table.columns[1].width = Inches(3.8)
    
    cell_l = admin_table.rows[0].cells[0]
    p_left = cell_l.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.add_run(f"{school_name.upper()}\n").bold = True
    p_left.add_run(f"{group_name.upper()}\n").bold = True
    p_left.add_run("Số: ..... /BB-TCM").font.size = Pt(11)
    
    cell_r = admin_table.rows[0].cells[1]
    p_right = cell_r.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n").bold = True
    p_right.add_run("Độc lập - Tự do - Hạnh phúc\n").bold = True
    p_right.add_run("***************").font.size = Pt(11)
    
    doc.add_paragraph()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run(title_name.upper()).bold = True
    
    in_table = False
    table_data = []
    
    def process_runs(paragraph, text):
        bold_parts = text.split('**')
        for i, b_part in enumerate(bold_parts):
            is_bold = (i % 2 != 0)
            sub_sup_parts = re.split(r'(<sub>.*?</sub>|<sup>.*?</sup>)', b_part)
            for part in sub_sup_parts:
                if not part: continue
                if part.startswith('<sub>') and part.endswith('</sub>'):
                    run = paragraph.add_run(part[5:-6]) 
                    run.bold = is_bold
                    run.font.subscript = True 
                elif part.startswith('<sup>') and part.endswith('</sup>'):
                    run = paragraph.add_run(part[5:-6]) 
                    run.bold = is_bold
                    run.font.superscript = True 
                else:
                    run = paragraph.add_run(part)
                    run.bold = is_bold

    def build_table():
        if not table_data: return
        cols = len(table_data[0])
        table = doc.add_table(rows=len(table_data), cols=cols)
        table.style = 'Table Grid'
        for r_idx, row in enumerate(table_data):
            for c_idx, cell_val in enumerate(row):
                if c_idx < cols:
                    cell = table.cell(r_idx, c_idx)
                    p = cell.paragraphs[0]
                    p.text = "" 
                    process_runs(p, cell_val.strip())
        doc.add_paragraph() 
        
    for line in text_content.split('\n'):
        cleaned_line = line.strip()
        
        if cleaned_line.startswith('|') and cleaned_line.endswith('|'):
            in_table = True
            row_data = [cell.strip() for cell in cleaned_line.split('|')[1:-1]]
            if all(re.match(r'^[-: ]+$', cell) for cell in row_data): continue
            table_data.append(row_data)
            continue
            
        if in_table:
            build_table()
            in_table = False
            table_data = []
            
        if not cleaned_line: continue
            
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        if cleaned_line.startswith('#'):
            process_runs(p, cleaned_line.replace('#', '').strip())
            for run in p.runs: run.bold = True
        else:
            process_runs(p, cleaned_line)
            
    if in_table: build_table()
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def render_exam_designer_section(api_key_input, run_ai_prompt_safe_func):
    # CSS Tùy chỉnh bám sát UI mẫu
    st.markdown("""
    <style>
    .header-pink { background-color: #FCE4EC; color: #880E4F; padding: 10px; text-align: center; font-weight: bold; font-size: 16px; border-radius: 4px; margin-bottom: 15px;}
    .header-green { background-color: #E8F5E9; color: #1B5E20; padding: 10px; text-align: center; font-weight: bold; font-size: 16px; border-radius: 4px; margin-bottom: 15px;}
    .footer-red { color: #D32F2F; font-weight: bold; font-style: italic; font-size: 14px; text-align: center; margin-top: 30px; padding-top: 10px; border-top: 1px solid #ccc;}
    div[data-testid="stNumberInput"] label { display: none !important; } /* Ẩn label mặc định của number_input để làm form inline */
    </style>
    """, unsafe_allow_html=True)

    st.markdown("### 📝 CHỨC NĂNG: TẠO ĐỀ KIỂM TRA AI")
    
    # Dòng trên cùng: Hình thức đề & Upload
    col_top1, col_top2 = st.columns([1, 1])
    with col_top1:
        c_lbl, c_sel = st.columns([1, 2])
        c_lbl.markdown("<div style='margin-top: 8px;'>Hình thức đề:</div>", unsafe_allow_html=True)
        hinh_thuc = c_sel.selectbox("", ["Trắc nghiệm kết hợp tự luận", "100% Trắc nghiệm", "100% Tự luận"], label_visibility="collapsed")
        
        mon_de = st.text_input("Môn học:", value="Khoa học tự nhiên")
        khoi_de = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"], index=3)
        thoi_gian_de = st.text_input("Thời gian:", value="45 phút")

    with col_top2:
        uploaded_files_de = st.file_uploader(
            "TẢI TÀI LIỆU LÊN (Giới hạn kiến thức/Đề cương):", 
            type=["pdf", "docx"], 
            accept_multiple_files=True
        )
        if not uploaded_files_de:
            st.markdown("*Chưa có tài liệu nào được tải lên hệ thống.*", unsafe_allow_html=True)
        else:
            st.success(f"Đã tải lên {len(uploaded_files_de)} tài liệu.")

    st.markdown("<hr style='margin: 10px 0px;'>", unsafe_allow_html=True)

    # KHU VỰC CHIA CỘT TRẮC NGHIỆM - TỰ LUẬN
    col_tn, spacer, col_tl = st.columns([10, 1, 10])
    
    with col_tn:
        st.markdown("<div class='header-pink'>PHẦN TRẮC NGHIỆM</div>", unsafe_allow_html=True)
        
        # Dòng 1
        c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
        c1.markdown("<b style='color:#C62828;'>Tổng số câu TNKQ:</b>", unsafe_allow_html=True)
        tong_so_tn = c2.number_input("Tổng TN", min_value=0, value=16)
        c3.markdown("<b>Tổng điểm TN:</b>", unsafe_allow_html=True)
        tong_diem_tn = c4.number_input("Tổng điểm TN", min_value=0.0, value=4.0, format="%.1f")
        
        # Dòng 2
        c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
        c1.markdown("Số câu nhiều lựa chọn:")
        tn_1_dap_an = c2.number_input("TN 1 ĐA", min_value=0, value=12)
        c3.markdown("Tổng điểm dòng này:")
        diem_tn_1 = c4.number_input("Điểm TN 1", min_value=0.0, value=2.0, format="%.1f")
        
        # Dòng 3
        c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
        c1.markdown("Số câu đúng sai:")
        tn_dung_sai = c2.number_input("TN Đ/S", min_value=0, value=2)
        c3.markdown("Tổng điểm dòng này:")
        diem_tn_2 = c4.number_input("Điểm TN 2", min_value=0.0, value=1.0, format="%.1f")

        # Dòng 4
        c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
        c1.markdown("Số câu điền khuyết:")
        tn_dien_khuyen = c2.number_input("TN ĐK", min_value=0, value=1)
        c3.markdown("Tổng điểm dòng này:")
        diem_tn_3 = c4.number_input("Điểm TN 3", min_value=0.0, value=0.5, format="%.1f")

        # Dòng 5
        c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
        c1.markdown("Số câu trả lời ngắn:")
        tn_tra_loi_ngan = c2.number_input("TN TLN", min_value=0, value=1)
        c3.markdown("Tổng điểm dòng này:")
        diem_tn_4 = c4.number_input("Điểm TN 4", min_value=0.0, value=0.5, format="%.1f")

    with col_tl:
        st.markdown("<div class='header-green'>PHẦN TỰ LUẬN</div>", unsafe_allow_html=True)
        
        c1, c2, c3, c4 = st.columns([3, 2, 2, 2])
        c1.markdown("<b style='color:#1565C0;'>TỔNG SỐ CÂU TỰ LUẬN:</b>", unsafe_allow_html=True)
        tong_so_tl = c2.number_input("Tổng TL", min_value=0, value=5)
        c3.markdown("<b>ĐIỂM:</b>", unsafe_allow_html=True)
        tong_diem_tl = c4.number_input("Tổng điểm TL", min_value=0.0, value=6.0, format="%.1f")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        diem_tl_list = []
        # Sinh form nhập điểm tự luận động dựa trên số câu
        for i in range(int(tong_so_tl)):
            rc1, rc2, rc3, rc4 = st.columns([1, 2, 2, 2])
            rc2.markdown(f"Câu {i+1}")
            diem_cau = rc3.number_input(f"Điểm câu {i+1}", min_value=0.0, value=1.0, format="%.1f", key=f"diem_tl_{i}")
            rc4.markdown("ĐIỂM")
            diem_tl_list.append(diem_cau)

    st.markdown("<hr style='margin: 15px 0px;'>", unsafe_allow_html=True)

    # KHU VỰC ĐIỀU KHIỂN & TỶ LỆ
    c_btn, c_chk = st.columns([2, 5])
    btn_tao = c_btn.button("⚙ Tự động tạo ma trận & đề thi", type="primary", use_container_width=True)
    uu_tien_de = c_chk.checkbox("Yêu cầu bám sát kiến thức trong tài liệu tải lên", value=True)

    st.markdown("<b>Tỷ lệ mức độ nhận thức (%):</b>", unsafe_allow_html=True)
    c_nb1, c_nb2, c_th1, c_th2, c_vd1, c_vd2, c_vdc1, c_vdc2 = st.columns([1,1,1,1,1,1,1,1])
    c_nb1.markdown("Nhận biết:")
    nb = c_nb2.number_input("NB", value=40)
    c_th1.markdown("Thông hiểu:")
    th = c_th2.number_input("TH", value=30)
    c_vd1.markdown("Vận dụng:")
    vd = c_vd2.number_input("VD", value=20)
    c_vdc1.markdown("Vận dụng cao:")
    vdc = c_vdc2.number_input("VDC", value=10)

    yeu_cau_khac = st.text_area("Nhập yêu cầu khác (Tùy chọn):", placeholder="Nhập yêu cầu khác ....")

    # XỬ LÝ SỰ KIỆN TẠO ĐỀ
    if btn_tao:
        if not api_key_input: 
            st.error("Thầy cần cấu hình Gemini API Key tại thanh bên!")
        else:
            with st.spinner("Hệ thống đang phân tích tài liệu và cấu trúc để sinh Ma trận & Đề thi..."):
                try:
                    content_de_nguon = ""
                    if uploaded_files_de:
                        for file in uploaded_files_de:
                            content_de_nguon += f"\n--- TÀI LIỆU: {file.name} ---\n"
                            if file.name.endswith('.docx'): content_de_nguon += read_uploaded_docx(file)
                            else: content_de_nguon += read_uploaded_pdf(file)
                    
                    diem_tl_str = ", ".join([f"Câu {i+1} ({diem_tl_list[i]} điểm)" for i in range(int(tong_so_tl))])

                    prompt_de = f"""Đóng vai một chuyên gia khảo thí xuất sắc. Hãy thiết kế Đề kiểm tra định kỳ môn {mon_de} {khoi_de}. Hình thức: {hinh_thuc}.
Thời gian: {thoi_gian_de}.
Cấu trúc điểm (Tỷ lệ {nb}-{th}-{vd}-{vdc}):
- PHẦN TRẮC NGHIỆM ({tong_so_tn} câu - {tong_diem_tn} điểm):
  + {tn_1_dap_an} câu nhiều lựa chọn (Tổng {diem_tn_1} điểm)
  + {tn_dung_sai} câu đúng/sai (Tổng {diem_tn_2} điểm)
  + {tn_dien_khuyen} câu điền khuyết (Tổng {diem_tn_3} điểm)
  + {tn_tra_loi_ngan} câu trả lời ngắn (Tổng {diem_tn_4} điểm)
- PHẦN TỰ LUẬN ({tong_so_tl} câu - {tong_diem_tl} điểm). Điểm chi tiết từng câu: {diem_tl_str}.
Yêu cầu khác: {yeu_cau_khac}
"""
                    if uu_tien_de and content_de_nguon:
                        prompt_de += f"\n\nBẮT BUỘC BÁM SÁT 100% NỘI DUNG TÀI LIỆU NGUỒN SAU ĐÂY:\n{content_de_nguon}"

                    prompt_de += """\n
LƯU Ý ĐỊNH DẠNG (BẮT BUỘC):
1. TUYỆT ĐỐI KHÔNG DÙNG LaTeX ($ hay $$).
2. Công thức Hóa/Toán dùng thẻ HTML (VD: H<sub>2</sub>O, x<sup>2</sup>).

TRÌNH BÀY ĐẦY ĐỦ 3 PHẦN (Dùng bảng Markdown |---|---| cho Ma trận và Đặc tả):
PHẦN 1. MA TRẬN ĐỀ KIỂM TRA
PHẦN 2. BẢNG ĐẶC TẢ CHI TIẾT
PHẦN 3. ĐỀ KIỂM TRA & ĐÁP ÁN"""

                    result_text, _ = run_ai_prompt_safe_func(prompt_de, api_key_input)
                    
                    # Lưu vào Session State để duy trì hiển thị
                    st.session_state["ket_qua_de_vua_tao"] = result_text
                    st.success("✅ Đã tạo đề thi thành công!")
                except Exception as error_ai: 
                    st.error(f"Lỗi hệ thống AI: {error_ai}")

    # KHU VỰC HIỂN THỊ KẾT QUẢ VÀ NÚT TẢI XUỐNG
    if "ket_qua_de_vua_tao" in st.session_state:
        col_dl1, col_dl2 = st.columns([8, 2])
        with col_dl2:
            st.download_button(
                label="📥 Tải về file Word (.docx)", 
                data=export_to_docx_vietnam_standard(st.session_state["ket_qua_de_vua_tao"], "ĐỀ KIỂM TRA"), 
                file_name="De_Kiem_Tra_AI.docx",
                type="primary",
                use_container_width=True
            )
        
        st.markdown("<div style='border: 1px solid #ccc; padding: 15px; border-radius: 5px; background-color: #fff;'>", unsafe_allow_html=True)
        st.markdown(st.session_state["ket_qua_de_vua_tao"], unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # Footer bản quyền
    st.markdown("<div class='footer-red'>© Bản quyền thuộc về Tác giả: Lê Hồng Dưỡng | Đơn vị: Trường THCS Nguyễn Chí Thanh – phường Tân Lập - tỉnh Đắk Lắk</div>", unsafe_allow_html=True)
