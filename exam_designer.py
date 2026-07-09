# exam_designer.py
import streamlit as st
import io
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import matplotlib.pyplot as plt
import numpy as np

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

def generate_plot_stream(eq_str):
    fig, ax = plt.subplots(figsize=(5, 3.5))
    x = np.linspace(-10, 10, 400)
    safe_dict = {"x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan, "sqrt": np.sqrt}
    try:
        eq_str_py = eq_str.replace('^', '**')
        y = eval(eq_str_py, {"__builtins__": {}}, safe_dict)
        if isinstance(y, (int, float)): y = np.full_like(x, y)
        ax.plot(x, y, color='#1E40AF', linewidth=2.5)
        ax.axhline(0, color='black', linewidth=1.2)
        ax.axvline(0, color='black', linewidth=1.2)
        ax.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)
        ax.set_ylim([-10, 10])
        ax.set_title(f"Đồ thị: y = {eq_str}", fontsize=10, pad=10)
    except:
        ax.text(0.5, 0.5, f"[Không thể vẽ đồ thị: Sai cú pháp toán học]", ha='center', va='center', color='red')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
    buf.seek(0)
    plt.close(fig)
    return buf

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
    admin_table.columns.width = Inches(3.2)
    admin_table.columns.width = Inches(3.8)
    
    cell_l = admin_table.rows.cells
    p_left = cell_l.paragraphs
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.add_run(f"{school_name.upper()}\n").bold = True
    p_left.add_run(f"{group_name.upper()}\n").bold = True
    p_left.add_run("Số: ..... /BB-TCM").font.size = Pt(11)
    
    cell_r = admin_table.rows.cells
    p_right = cell_r.paragraphs
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
        cols = len(table_data)
        table = doc.add_table(rows=len(table_data), cols=cols)
        table.style = 'Table Grid'
        for r_idx, row in enumerate(table_data):
            for c_idx, cell_val in enumerate(row):
                if c_idx < cols:
                    cell = table.cell(r_idx, c_idx)
                    p = cell.paragraphs
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
        if '[GRAPH:' in cleaned_line:
            match = re.search(r'\[GRAPH:\s*(.+?)\]', cleaned_line)
            if match:
                eq = match.group(1)
                img_stream = generate_plot_stream(eq)
                doc.add_picture(img_stream, width=Inches(3.5))
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                continue 
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
    # CSS thiết kế các khối nhãn và tiêu đề ngay ngắn, tự nhiên, không chồng chữ
    st.markdown("""
    <style>
    .title-pink-header { background-color: #FCE4EC; color: #C2185B; padding: 10px; text-align: center; font-weight: bold; font-size: 17px; border-radius: 4px; border: 1px solid #F87171; margin-bottom: 15px;}
    .title-green-header { background-color: #E8F5E9; color: #1B5E20; padding: 10px; text-align: center; font-weight: bold; font-size: 17px; border-radius: 4px; border: 1px solid #4ADE80; margin-bottom: 15px;}
    
    /* Ô hiển thị tổng số câu và tổng điểm tự động cập nhật */
    .row-summary-pink { display: flex; justify-content: space-between; align-items: center; border: 1px dashed #F87171; padding: 8px 12px; border-radius: 4px; font-weight: bold; color: #B91C1C; background-color: #FFF5F5; margin-bottom: 20px; }
    .row-summary-green { display: flex; justify-content: space-between; align-items: center; border: 1px solid #4ADE80; padding: 8px 12px; border-radius: 4px; font-weight: bold; color: #166534; background-color: #F0FDF4; margin-bottom: 20px; }
    .val-cell-box { background: white; padding: 4px 18px; border: 1px solid #D1D5DB; border-radius: 4px; margin-left: 8px; font-family: monospace; font-size: 14px; color: black; display: inline-block;}
    
    .text-bold-label { font-size: 16px; font-weight: bold; color: black; margin-top: 6px; }
    .text-italic-unit { font-size: 15px; font-style: italic; color: black; margin-top: 8px; }
    .footer-red { color: #D32F2F; font-weight: bold; font-style: italic; font-size: 14px; text-align: center; margin-top: 30px; padding-top: 10px; border-top: 1px solid #ccc;}
    
    /* Khử nhãn mặc định Streamlit */
    div[data-testid="stNumberInput"] label { display: none !important; }
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stTextInput"] label { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    if "db_de_kiem_tra" not in st.session_state: st.session_state["db_de_kiem_tra"] = []
    if "current_exam_designer_output" not in st.session_state: st.session_state["current_exam_designer_output"] = ""

    tab_thiet_ke, tab_kho_luu_tru = st.tabs(["📝 CHỨC NĂNG: TẠO ĐỀ KIỂM TRA AI", "📂 THƯ MỤC ĐỀ ĐÃ XÂY DỰNG"])
    
    with tab_thiet_ke:
        # Thanh điều khiển kéo thả tài liệu ở phía trên
        col_top_lbl, col_top_btn1, col_top_btn2 = st.columns([2.5, 1.3, 1.5])
        hinh_thuc = col_top_lbl.selectbox("Hình thức đề thi cấu hình:", ["Trắc nghiệm kết hợp tự luận", "100% Trắc nghiệm", "100% Tự luận"])
        col_top_btn1.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
        col_top_btn1.button("TÁI KHUNG MA TRẬN MẪU", type="secondary", use_container_width=True)
        
        uploaded_files = col_top_btn2.file_uploader("Tải tài liệu nền", type=["docx", "pdf", "xlsx", "xls"], accept_multiple_files=True, label_visibility="collapsed", key="uploader_multi_exam_v9")
        if uploaded_files:
            st.markdown(f"**🔹 Đã kết nối ({len(uploaded_files)}) tệp tài liệu nền tham khảo.**")
        else:
            st.markdown("<p style='color: gray; font-size: 12px; font-style: italic; margin-top:-10px;'>🌐 Chưa có tài liệu nào được tải lên hệ thống.</p>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_main1, col_main2 = st.columns(2)
        
        # --- 💥 CỘT LỆCH TRÁI: PHẦN TRẮC NGHIỆM ---
        with col_main1:
            st.markdown("<div class='title-pink-header'>PHẦN TRẮC NGHIỆM</div>", unsafe_allow_html=True)
            
            # Khởi tạo các ô nhập số liệu thành phần của phần trắc nghiệm
            r1_l, r1_i, r1_u, r1_lbl, r1_sc = st.columns([1.5, 0.5, 0.4, 0.5, 0.6])
            r1_l.markdown("<div class='text-bold-label'>Câu nhiều lựa chọn:</div>", unsafe_allow_html=True)
            c1 = r1_i.number_input("c1", min_value=0, max_value=50, value=12, step=1)
            r1_u.markdown("<div class='text-bold-label'>câu.</div>", unsafe_allow_html=True)
            r1_lbl.markdown("<div class='text-bold-label' style='text-align:right;'>Điểm</div>", unsafe_allow_html=True)
            p1 = r1_sc.number_input("p1", min_value=0.0, max_value=10.0, value=3.0, step=0.25)
            
            r2_l, r2_i, r2_u, r2_lbl, r2_sc = st.columns([1.5, 0.5, 0.4, 0.5, 0.6])
            r2_l.markdown("<div class='text-bold-label'>Câu đúng sai:</div>", unsafe_allow_html=True)
            c2 = r2_i.number_input("c2", min_value=0, max_value=50, value=2, step=1)
            r2_u.markdown("<div class='text-bold-label'>câu.</div>", unsafe_allow_html=True)
            r2_lbl.markdown("<div class='text-bold-label' style='text-align:right;'>Điểm</div>", unsafe_allow_html=True)
            p2 = r2_sc.number_input("p2", min_value=0.0, max_value=10.0, value=0.5, step=0.25)
            
            r3_l, r3_i, r3_u, r3_lbl, r3_sc = st.columns([1.5, 0.5, 0.4, 0.5, 0.6])
            r3_l.markdown("<div class='text-bold-label'>Câu điền khuyết:</div>", unsafe_allow_html=True)
            c3 = r3_i.number_input("c3", min_value=0, max_value=50, value=1, step=1)
            r3_u.markdown("<div class='text-bold-label'>câu.</div>", unsafe_allow_html=True)
            r3_lbl.markdown("<div class='text-bold-label' style='text-align:right;'>Điểm</div>", unsafe_allow_html=True)
            p3 = r3_sc.number_input("p3", min_value=0.0, max_value=10.0, value=0.25, step=0.25)
            
            r4_l, r4_i, r4_u, r4_lbl, r4_sc = st.columns([1.5, 0.5, 0.4, 0.5, 0.6])
            r4_l.markdown("<div class='text-bold-label'>Câu trả lời ngắn:</div>", unsafe_allow_html=True)
            c4 = r4_i.number_input("c4", min_value=0, max_value=50, value=1, step=1)
            r4_u.markdown("<div class='text-bold-label'>câu.</div>", unsafe_allow_html=True)
            r4_lbl.markdown("<div class='text-bold-label' style='text-align:right;'>Điểm</div>", unsafe_allow_html=True)
            p4 = r4_sc.number_input("p4", min_value=0.0, max_value=10.0, value=0.25, step=0.25)
            
            # Tính toán tổng câu và tổng điểm trắc nghiệm thời gian thực
            tong_cau_tn_real = c1 + c2 + c3 + c4
            tong_diem_tn_real = p1 + p2 + p3 + p4
            
            # Khối hiển thị tổng số câu đứng cố định an toàn ngay đầu cột hồng
            st.markdown(
                f"<div class='row-summary-pink'>"
                f"<div>Tổng số câu TNKQ: <span class='val-cell-box'>{tong_cau_tn_real}</span></div>"
                f"<div>Tổng điểm TN: <span class='val-cell-box'>{tong_diem_tn_real:.2f}</span></div>"
                f"</div>"
                f"<div class='text-bold-label' style='text-decoration: underline; margin-bottom:15px;'>Trong đó:</div>", unsafe_allow_html=True
            )

        # --- 💥 CỘT LỆCH PHẢI: PHẦN TỰ LUẬN (MỞ KHÓA TOÀN DIỆN TỰ SINH VÀ ĐẾM ĐIỂM) ---
        with col_main2:
            st.markdown("<div class='title-green-header'>PHẦN TỰ LUẬN</div>", unsafe_allow_html=True)
            
            # 🌟 MỞ KHÓA: Sử dụng st.number_input tự do thay đổi số câu tự luận, mặc định gợi ý là 4 câu
            num_tl_input = st.number_input("TỔNG SỐ CÂU TỰ LUẬN:", min_value=0, max_value=20, value=4, step=1, key="num_tl_v10")
            
            tl_scores = []
            if num_tl_input > 0:
                for i in range(num_tl_input):
                    rl_l, rl_i, rl_u = st.columns([1.2, 0.8, 1.7])
                    rl_l.markdown(f"<div class='text-bold-label' style='text-align:right; padding-right:20px;'>Câu {i+1}:</div>", unsafe_allow_html=True)
                    # Điểm gợi ý ban đầu cho 4 câu đầu tiên theo phôi của thầy, các câu sinh thêm từ câu thứ 5 gán mặc định 1.0đ
                    init_score = 1.5 if i==0 or i==1 else (2.0 if i==2 else 1.0)
                    score_cell = rl_i.number_input(f"s_tl_{i}", min_value=0.0, max_value=10.0, value=init_score if i < 4 else 1.0, step=0.5)
                    rl_u.markdown("<div class='text-italic-unit'>điểm</div>", unsafe_allow_html=True)
                    tl_scores.append(score_cell)
                    
            # Thuật toán cộng dồn động: Tính tổng điểm tự luận từ mảng điểm thực tế thầy nhập
            tong_diem_tl_real = sum(tl_scores)
            
            # 🌟 ĐỒNG BỘ ĐỘNG: Đẩy dữ liệu tính toán thời gian thực lên ô tổng kết quả (Khử lỗi khóa ô cứng)
            st.markdown(
                f"<div class='row-summary-green' style='margin-top:-{num_tl_input*44 + 115}px; margin-bottom:{num_tl_input*44 + 75}px; position:relative; z-index:10; pointer-events:none;'>"
                f"<div>TỔNG SỐ CÂU TỰ LUẬN: <span class='val-cell-box'>{num_tl_input}</span></div>"
                f"<div>ĐIỂM: <span class='val-cell-box'>{tong_diem_tl_real:.1f}</span></div>"
                f"</div>", unsafe_allow_html=True
            )
            
        st.markdown("<br>", unsafe_allow_html=True)

        # --- 🌟 VÁ LỖI CÚ PHÁP: Truyền chính xác số lượng cột phân bổ (2) vào hàm .columns ---
        st.markdown("**Tỷ lệ mức độ nhận thức (%):**")
        c_m1, c_m2, c_m3, c_m4 = st.columns(4)
        
        cm1_l, cm1_i = c_m1.columns(2) # Đã vá lỗi truyền tham số số 2
        cm1_l.markdown("<div class='text-bold-label'>Nhận biết:</div>", unsafe_allow_html=True)
        mz_nb = cm1_i.number_input("m1", min_value=0, max_value=100, value=40, step=5)
        
        cm2_l, cm2_i = c_m2.columns([1.5, 1])
        cm2_l.markdown("<div class='text-bold-label'>Thông hiểu:</div>", unsafe_allow_html=True)
        mz_th = cm2_i.number_input("m2", min_value=0, max_value=100, value=30, step=5)
        
        cm3_l, cm3_i = c_m3.columns([1.8, 1])
        cm3_l.markdown("<div class='text-bold-label'>Vận dụng thấp:</div>", unsafe_allow_html=True)
        mz_vd = cm3_i.number_input("m3", min_value=0, max_value=100, value=20, step=5)
        
        cm4_l, cm4_i = c_m4.columns([1.8, 1])
        cm4_l.markdown("<div class='text-bold-label'>Vận dụng cao:</div>", unsafe_allow_html=True)
        mz_vdc = cm4_i.number_input("m4", min_value=0, max_value=100, value=10, step=5)

        st.markdown("<br>", unsafe_allow_html=True)
        col_btn_zone, col_clear_file_zone, col_check_zone = st.columns([2, 1.5, 2.5])
        run_exam_ai = col_btn_zone.button("🚀 Tự động tạo ma trận & đề thi", type="primary", use_container_width=True, key="btn_run_exam_ultimate")
        
        if uploaded_files:
            if col_clear_file_zone.button("🗑️ Xóa file vừa tải", type="secondary", use_container_width=True, key="btn_clear_uploaded_files"):
                st.session_state["uploader_multi_exam_v9"] = []
                st.success("💥 Đã dọn sạch các tệp tài liệu tham khảo!")
                st.rerun()
        yeu_cau_bam_sat = col_check_zone.checkbox("Yêu cầu bám sát kiến thức trong tài liệu tải lên", value=True, key="chk_bam_sat_ultimate")

        note_de = st.text_area("Nhập yêu cầu cụ thể (Tùy chọn)", placeholder="Nhập yêu cầu khác....", label_visibility="collapsed", key="note_de_area_u")

        if run_exam_ai:
            combined_context_text = ""
            if uploaded_files:
                for file_node in uploaded_files:
                    if file_node.name.endswith(".docx"): combined_context_text += read_uploaded_docx(file_node) + "\n\n"
                    elif file_node.name.endswith(".pdf"): combined_context_text += read_uploaded_pdf(file_node) + "\n\n"

            prompt_exam = (
                f"Hãy thiết kế một Ma trận đề thi và Đề kiểm tra chi tiết (kèm Đáp án) cho môn: Khoa học tự nhiên, hình thức: {hinh_thuc}, thời gian: 45 phút.\n"
                f"CẤU TRÚC ĐỀ BẮT BUỘC ĐỒNG BỘ ĐỊNH MỨC:\n"
                f"- Trắc nghiệm tổng cộng {tong_cau_tn_real} câu với {tong_diem_tn_real} điểm, chi tiết gồm: Nhiều lựa chọn: {c1} câu, Đúng/sai: {c2} câu, Điền khuyết: {c3} câu, Trả lời ngắn: {c4} câu.\n"
                f"- Tự luận tổng cộng {num_tl_input} câu lớn với tổng điểm là {tong_diem_tl_real} điểm.\n"
                f"MA TRẬN MỨC ĐỘ NHẬN THỨC: Nhận biết {mz_nb}%, Thông hiểu {mz_th}%, Vận dụng {mz_vd}%, Vận dụng cao {mz_vdc}%.\n"
                f"Nhiệm vụ cốt lõi: Bạn bắt buộc phải tổng hợp nội dung, phân tích kiến thức bám sát theo tất cả các tài liệu đính kèm sau đây:\n"
                f"[DỮ LIỆU THAM KHẢO TÀI LIỆU TẢI LÊN]:\n{combined_context_text}\n\n"
                f"Yêu cầu nội dung kiến thức bổ sung: {note_de}.\n"
                f"Trả về văn bản định dạng Markdown đẹp mắt, ma trận dựng bảng dạng '|'."
            )

            if run_ai_prompt_safe_func is not None:
                with st.spinner("🚀 Trợ lý AI đang liên kết dữ liệu đa tài liệu và thiết kế đề thi..."):
                    res_text, status = run_ai_prompt_safe_func(prompt_exam)
                    if status == "error": st.error(f"❌ Lỗi kết nối: {res_text}")
                    else:
                        st.session_state["current_exam_designer_output"] = res_text
                        st.session_state["db_de_kiem_tra"].append({
                            "id": f"De_{len(st.session_state['db_de_kiem_tra'])+1}", "mon": "Khoa học tự nhiên", "hinh_thuc": hinh_thuc, "data": res_text
                        })
                        st.success(f"🎉 Khởi tạo đề kiểm tra thành công bằng mô hình {status}!")
                        st.rerun()
            else:
                st.error("❌ Lỗi luồng: Chưa kết nối được trình điều khiển AI tổng từ file app.py.")

        if st.session_state["current_exam_designer_output"]:
            st.markdown("---")
            st.markdown(st.session_state["current_exam_designer_output"])
            word_exam_data = export_to_docx_vietnam_standard(st.session_state["current_exam_designer_output"], "ĐỀ KIỂM TRA MÔN KHOA HỌC TỰ NHIÊN")
            st.download_button(label="📥 Tải file Word (.docx) Đề kiểm tra chuẩn Quốc hiệu", data=word_exam_data, file_name="De_Kiem_Tra_KHTN.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    with tab_kho_luu_tru:
        st.markdown("#### 📂 Các đề kiểm tra đã được AI tự động sinh và lưu trữ")
        if not st.session_state["db_de_kiem_tra"]:
            st.caption("✨ Chưa có đề kiểm tra nào được tạo.")
        else:
            indices_to_delete = []
            for idx, de_info in enumerate(st.session_state["db_de_kiem_tra"]):
                with st.expander(f"📋 Đề môn {de_info['mon']} - {de_info['hinh_thuc']}"):
                    st.markdown(de_info['data'])
                    h_col1, h_col2 = st.columns(2)
                    word_hist_data = export_to_docx_vietnam_standard(de_info['data'], f"ĐỀ KIỂM TRA MÔN {de_info['mon'].upper()}")
                    h_col1.download_button(label="📥 Tải file Word", data=word_hist_data, file_name=f"De_{idx+1}.docx", key=f"dl_ex_w_{idx}", use_container_width=True)
                    if h_col2.button("❌ Xóa đề thi", key=f"del_ex_btn_{idx}", use_container_width=True): indices_to_delete.append(idx)
            if indices_to_delete:
                for index in sorted(indices_to_delete, reverse=True): st.session_state["db_de_kiem_tra"].pop(index)
                st.success("🗑️ Đã xóa đề thi!")
                st.rerun()

    st.markdown("<div class='footer-red'>© Bản quyền thuộc về Tác giả: Lê Hồng Dưỡng | Đơn vị: Trường THCS Nguyễn Chí Thanh - phường Tân Lập - tỉnh Đắk Lắk</div>", unsafe_allow_html=True)
