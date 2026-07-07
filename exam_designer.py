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
    
    def build_table():
        if not table_data: return
        cols = len(table_data[0])
        table = doc.add_table(rows=len(table_data), cols=cols)
        table.style = 'Table Grid'
        for r_idx, row in enumerate(table_data):
            for c_idx, cell_val in enumerate(row):
                if c_idx < cols:
                    clean_val = cell_val.replace('**', '').strip()
                    table.cell(r_idx, c_idx).text = clean_val
        doc.add_paragraph() 
        
    for line in text_content.split('\n'):
        cleaned_line = line.strip()
        
        if cleaned_line.startswith('|') and cleaned_line.endswith('|'):
            in_table = True
            row_data = [cell.strip() for cell in cleaned_line.split('|')[1:-1]]
            if all(re.match(r'^[-: ]+$', cell) for cell in row_data):
                continue
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
            p.add_run(cleaned_line.replace('#', '').strip()).bold = True
        elif '**' in cleaned_line:
            parts = cleaned_line.split('**')
            for i, part in enumerate(parts):
                run = p.add_run(part)
                if i % 2 != 0: run.bold = True
        else:
            p.add_run(cleaned_line)
            
    if in_table: build_table()
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def render_exam_designer_section(api_key_input, run_ai_prompt_safe_func):
    """
    Hàm đóng gói giao diện và logic thiết kế đề thi để nhúng vào file chính.
    """
    st.header("📊 THIẾT KẾ HỆ THỐNG ĐỀ KIỂM TRA ĐÁNH GIÁ ĐỊNH KỲ")
    tab_thiet_ke, tab_kho_luu_tru = st.tabs(["✨ Thiết kế đề thi mới", "📂 Thư mục lưu trữ đề đã dựng"])
    
    with tab_thiet_ke:
        col_de1, col_de2 = st.columns(2)
        with col_de1:
            ten_de = st.text_input("Tên kỳ kiểm tra thiết lập:", placeholder="Ví dụ: Kiểm tra định kỳ học kỳ I")
            mon_de = st.text_input("Môn học kiểm tra:", value="Khoa học tự nhiên")
            khoi_de = st.selectbox("Khối lớp đề thi:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"], index=3)
        with col_de2:
            thoi_gian_de = st.text_input("Thời gian làm bài thi:", value="45 phút")
            ty_le_de = st.text_input("Tỷ lệ ma trận mong muốn:", value="Nhận biết: 40% - Thông hiểu: 30% - Vận dụng thấp: 20% - Vận dụng cao: 10%")
            
        st.markdown("#### Cấu hình cấu trúc câu hỏi chi tiết")
        col_sl1, col_sl2 = st.columns(2)
        with col_sl1:
            tong_so_tn = st.slider("Tổng số câu trắc nghiệm khách quan:", min_value=0, max_value=40, value=16)
        with col_sl2:
            tong_so_tl = st.number_input("Tổng số câu hỏi tự luận tùy chọn:", min_value=0, value=2, step=1)
            
        st.markdown("##### Phân bổ chi tiết hình thức câu hỏi trắc nghiệm")
        col_hb1, col_sl_cs2, col_sl_cs3, col_sl_cs4 = st.columns(4)
        with col_hb1: tn_1_dap_an = st.number_input("Trắc nghiệm 1 đáp án đúng:", min_value=0, value=10, step=1)
        with col_sl_cs2: tn_dung_sai = st.number_input("Trắc nghiệm Đúng/Sai:", min_value=0, value=2, step=1)
        with col_sl_cs3: tn_dien_khuyen = st.number_input("Trắc nghiệm điền khuyết:", min_value=0, value=2, step=1)
        with col_sl_cs4: tn_tra_loi_ngan = st.number_input("Trắc nghiệm trả lời ngắn:", min_value=0, value=2, step=1)
            
        tong_phan_bo_thuc_te = tn_1_dap_an + tn_dung_sai + tn_dien_khuyen + tn_tra_loi_ngan
        if tong_phan_bo_thuc_te == tong_so_tn:
            st.success(f"✅ Đồng bộ thành công! Tổng phân bổ ({tong_phan_bo_thuc_te} câu) khớp với cấu hình đề thi.")
        else:
            st.warning(f"⚠️ Cảnh báo: Tổng phân bổ chi tiết ({tong_phan_bo_thuc_te} câu) chưa khớp với Tổng số câu ({tong_so_tn} câu).")
            
        uploaded_excel_de = st.file_uploader("Tải lên file giới hạn kiến thức hoặc bộ đề thi làm căn cứ (.pdf, .docx):", type=["pdf", "docx"], key="up_file_de_thi")
        content_de_nguon = ""
        if uploaded_excel_de:
            content_de_nguon = read_uploaded_docx(uploaded_excel_de) if uploaded_excel_de.name.endswith('.docx') else read_uploaded_pdf(uploaded_excel_de)
            st.success("✅ Đã nạp thành công tài liệu kiến thức nguồn vào bộ nhớ AI.")
            
        uu_tien_de = st.checkbox("🎯 BÁM SÁT 100% TÀI LIỆU/KIẾN THỨC NGUỒN VÀ BẮT BUỘC TẠO MA TRẬN", value=True)
        
        if st.button("🛠️ Thiết lập Đề thi + Ma trận & Đặc tả"):
            if not api_key_input: st.error("Thầy cần cấu hình Gemini API Key tại thanh bên!")
            else:
                with st.spinner("AI đang xử lý Ma trận, Đặc tả và Đề thi..."):
                    try:
                        prompt_de = f"""Đóng vai một chuyên gia khảo thí xuất sắc. Hãy thiết kế Đề kiểm tra đánh giá định kỳ môn {mon_de} lớp {khoi_de}.
Thời gian làm bài: {thoi_gian_de}.
Cấu trúc yêu cầu:
- Trắc nghiệm: {tong_so_tn} câu ({tn_1_dap_an} câu 1 đáp án, {tn_dung_sai} câu đúng/sai, {tn_dien_khuyen} câu điền khuyết, {tn_tra_loi_ngan} câu trả lời ngắn).
- Tự luận: {tong_so_tl} câu.
- Tỷ lệ các mức độ nhận thức: {ty_le_de}.
"""
                        if uu_tien_de and content_de_nguon:
                            prompt_de += f"""\n\nBẮT BUỘC BÁM SÁT TUYỆT ĐỐI 100% NỘI DUNG TÀI LIỆU NGUỒN SAU ĐÂY MỚI ĐƯỢC RA ĐỀ:
--- BẮT ĐẦU TÀI LIỆU NGUỒN ---
{content_de_nguon}
--- KẾT THÚC TÀI LIỆU NGUỒN ---"""

                        prompt_de += """\n
BẠN PHẢI TRÌNH BÀY ĐẦY ĐỦ VÀ RÕ RÀNG THEO 3 PHẦN DƯỚI ĐÂY:
PHẦN 1. MA TRẬN ĐỀ KIỂM TRA (Sử dụng bảng định dạng Markdown để kẻ khung)
PHẦN 2. BẢNG ĐẶC TẢ CHI TIẾT (Sử dụng bảng định dạng Markdown để kẻ khung)
PHẦN 3. ĐỀ KIỂM TRA VÀ ĐÁP ÁN CHI TIẾT (Chia rõ Trắc nghiệm và Tự luận)"""

                        result_text, _ = run_ai_prompt_safe_func(prompt_de, api_key_input)
                        st.markdown(result_text)
                        st.session_state["db_de_kiem_tra"].append({"ten_de": ten_de if ten_de else "Đề tự động dựng", "mon": mon_de, "khoi": khoi_de, "noi_dung": result_text})
                        st.success("✅ Đã lưu bộ đề thi thành công vào Thư mục lưu trữ!")
                    except Exception as error_ai: st.error(f"Lỗi hệ thống AI: {error_ai}")
                        
    with tab_kho_luu_tru:
        st.subheader("📁 Thư mục lưu trữ đề kiểm tra nội bộ đã dựng")
        if not st.session_state["db_de_kiem_tra"]:
            st.info("💡 Chưa có đề kiểm tra nào được tạo. Hãy thiết kế đề thi mới ở tab bên cạnh.")
        else:
            for idx, item in enumerate(st.session_state["db_de_kiem_tra"]):
                with st.expander(f"📋 {item['ten_de']} - Môn: {item['mon']} ({item['khoi']})"):
                    st.markdown(item["noi_dung"])
                    col_bt1, col_bt2 = st.columns([1, 1])
                    with col_bt1:
                        st.download_button(label="📥 Tải File Word (Kèm ma trận & kẻ bảng)", data=export_to_docx_vietnam_standard(item["noi_dung"], item["ten_de"]), file_name=f"{item['ten_de']}_{idx}.docx", key=f"dl_de_thi_{idx}")
                    with col_bt2:
                        if st.button("🗑️ Xóa đề kiểm tra này", key=f"del_de_thi_{idx}"):
                            st.session_state["db_de_kiem_tra"].pop(idx)
                            st.rerun()