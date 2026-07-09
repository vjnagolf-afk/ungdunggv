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
    
    safe_dict = {
        "x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan, "sqrt": np.sqrt
    }
    try:
        eq_str_py = eq_str.replace('^', '**')
        y = eval(eq_str_py, {"__builtins__": {}}, safe_dict)
        
        if isinstance(y, (int, float)):
            y = np.full_like(x, y)

        ax.plot(x, y, color='#1E40AF', linewidth=2.5)
        ax.axhline(0, color='black', linewidth=1.2)
        ax.axvline(0, color='black', linewidth=1.2)
        ax.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)
        
        ax.set_ylim([-10, 10])
        ax.set_title(f"Đồ thị: y = {eq_str}", fontsize=10, pad=10)
    except Exception as e:
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
# --- GIAO DIỆN CHÍNH PHÂN HỆ THIẾT KẾ ĐỀ ---
def render_exam_designer_section(run_ai_prompt_safe_func=None):
    st.markdown("""
    <style>
    .header-pink { background-color: #FCE4EC; color: #880E4F; padding: 10px; text-align: center; font-weight: bold; font-size: 16px; border-radius: 4px; margin-bottom: 15px;}
    .header-green { background-color: #E8F5E9; color: #1B5E20; padding: 10px; text-align: center; font-weight: bold; font-size: 16px; border-radius: 4px; margin-bottom: 15px;}
    .footer-red { color: #D32F2F; font-weight: bold; font-style: italic; font-size: 14px; text-align: center; margin-top: 30px; padding-top: 10px; border-top: 1px solid #ccc;}
    div[data-testid="stNumberInput"] label { display: none !important; } 
    div[data-testid="stTextInput"] label { display: none !important; } 
    div[data-testid="stSelectbox"] label { display: none !important; }
    div[data-testid="stTabs"] button { font-size: 22px !important; font-weight: 800 !important; color: #1E3A8A !important; }
    div[data-testid="stTabs"] button[aria-selected="true"] { color: #E11D48 !important; border-bottom-color: #E11D48 !important; }
    </style>
    """, unsafe_allow_html=True)

    if "db_de_kiem_tra" not in st.session_state:
        st.session_state["db_de_kiem_tra"] = []
    if "current_exam_designer_output" not in st.session_state:
        st.session_state["current_exam_designer_output"] = ""

    tab_thiet_ke, tab_kho_luu_tru = st.tabs(["📝 CHỨC NĂNG: TẠO ĐỀ KIỂM TRA AI", "📂 THƯ MỤC ĐỀ ĐÃ XÂY DỰNG"])
    
    with tab_thiet_ke:
        col_top1, col_top2 = st.columns([1, 1])
        with col_top1:
            c_lbl, c_sel = st.columns([1, 2])
            c_lbl.markdown("<div style='margin-top: 8px;'>Hình thức đề:</div>", unsafe_allow_html=True)
            hinh_thuc = c_sel.selectbox("Hinh_thuc", ["Trắc nghiệm kết hợp tự luận", "100% Trắc nghiệm", "100% Tự luận"])
            
            c_lbl2, c_txt2 = st.columns([1, 2])
            c_lbl2.markdown("<div style='margin-top: 8px;'>Môn học:</div>", unsafe_allow_html=True)
            mon_de = c_txt2.text_input("Mon", value="Khoa học tự nhiên")
            
            c_lbl3, c_sel3 = st.columns([1, 2])
            c_lbl3.markdown("<div style='margin-top: 8px;'>Khối lớp:</div>", unsafe_allow_html=True)
            khoi_de = c_sel3.selectbox("Khoi", ["Khối 6", "Khối 7", "Khối 8", "Khối 9"])
            
            c_lbl4, c_txt4 = st.columns([1, 2])
            c_lbl4.markdown("<div style='margin-top: 8px;'>Thời gian làm bài:</div>", unsafe_allow_html=True)
            tg_de = c_txt4.text_input("Thoi_gian", value="45 phút")

            st.markdown("**📂 Đính kèm tài liệu phân phối chương trình hoặc ma trận mẫu (Tùy chọn):**")
            tai_lieu_dinh_kem = st.file_uploader("Tải tài liệu nền (.docx, .pdf)", type=["docx", "pdf"], key="file_de_tai_lieu")

        with col_top2:
            st.markdown("**📊 Cấu trúc số câu hỏi và ma trận điểm dòng:**")
            col_sc1, col_sc2 = st.columns(2)
            num_tn = col_sc1.number_input("Số câu trắc nghiệm:", min_value=0, max_value=50, value=12, step=1, key="num_tn")
            pt_tn = col_sc2.number_input("Tổng điểm trắc nghiệm:", min_value=0.0, max_value=10.0, value=3.0, step=0.25, key="pt_tn")
            
            num_tl = col_sc1.number_input("Số câu tự luận:", min_value=0, max_value=20, value=3, step=1, key="num_tl")
            pt_tl = col_sc2.number_input("Tổng điểm tự luận:", min_value=0.0, max_value=10.0, value=4.0, step=0.25, key="pt_tl")
            
            num_khuyet = col_sc1.number_input("Số câu điền khuyết:", min_value=0, max_value=20, value=4, step=1, key="num_khuyet")
            pt_khuyet = col_sc2.number_input("Tổng điểm điền khuyết:", min_value=0.0, max_value=10.0, value=1.5, step=0.25, key="pt_khuyet")
            
            num_ngan = col_sc1.number_input("Số câu trả lời ngắn:", min_value=0, max_value=20, value=4, step=1, key="num_ngan")
            pt_ngan = col_sc2.number_input("Tổng điểm trả lời ngắn:", min_value=0.0, max_value=10.0, value=1.5, step=0.25, key="pt_ngan")

        st.markdown("---")
        st.markdown("**🎯 Cấu hình tỷ lệ mức độ nhận thức thi cử (%):**")
        col_mz1, col_mz2, col_mz3, col_mz4 = st.columns(4)
        mz_nb = col_mz1.number_input("Nhận biết (%):", min_value=0, max_value=100, value=40, step=5, key="mz_nb")
        mz_th = col_mz2.number_input("Thông hiểu (%):", min_value=0, max_value=100, value=30, step=5, key="mz_th")
        mz_vd = col_mz3.number_input("Vận dụng (%):", min_value=0, max_value=100, value=20, step=5, key="mz_vd")
        mz_vdc = col_mz4.number_input("Vận dụng cao (%):", min_value=0, max_value=100, value=10, step=5, key="mz_vdc")

        st.markdown("**💬 Nhập yêu cầu cụ thể cho nội dung kiến thức đề thi:**")
        note_de = st.text_area("Yêu cầu bổ sung cho AI:", placeholder="Ví dụ: Đề thi giữa học kỳ II môn Vật lý 7, tập trung vào chương Ánh sáng và Tốc độ...", label_visibility="collapsed", key="note_de_area")

        run_exam_ai = st.button("✨ Tự động tạo ma trận & đề thi bằng AI", type="primary", use_container_width=True)
        if run_exam_ai:
            # Thu thập văn bản nền từ file upload đính kèm nếu có
            context_text = ""
            if tai_lieu_dinh_kem is not None:
                if tai_lieu_dinh_kem.name.endswith(".docx"):
                    context_text = read_uploaded_docx(tai_lieu_dinh_kem)
                elif tai_lieu_dinh_kem.name.endswith(".pdf"):
                    context_text = read_uploaded_pdf(tai_lieu_dinh_kem)

            # Thiết lập câu lệnh chi tiết ép AI sinh Ma trận và Đề thi chuẩn chỉ
            prompt_exam = (
                f"Hãy thiết kế một Ma trận đề thi và Đề kiểm tra chi tiết (kèm Đáp án) cho môn: {mon_de}, {khoi_de}, hình thức: {hinh_thuc}, thời gian: {tg_de}.\n"
                f"CẤU TRÚC ĐỀ BẮT BUỘC:\n"
                f"- Trắc nghiệm: {num_tn} câu ({pt_tn} điểm)\n- Tự luận: {num_tl} câu ({pt_tl} điểm)\n"
                f"- Điền khuyết: {num_khuyet} câu ({pt_khuyet} điểm)\n- Trả lời ngắn: {num_ngan} câu ({pt_ngan} điểm)\n"
                f"MA TRẬN MỨC ĐỘ NHẬN THỨC: Nhận biết {mz_nb}%, Thông hiểu {mz_th}%, Vận dụng {mz_vd}%, Vận dụng cao {mz_vdc}%.\n"
                f"Yêu cầu nội dung kiến thức bổ sung: {note_de}.\n"
                f"Tài liệu tham khảo đính kèm (nếu có):\n{context_text}\n\n"
                f"Văn bản trả về trình bày đẹp mắt bằng Markdown. Đối với các bảng ma trận, sử dụng cấu trúc bảng kẻ dọc dạng '|'. "
                f"Nếu bài toán hình học hoặc đồ thị cần vẽ hình, hãy chèn thẻ mẫu dạng '[GRAPH: biểu_thức_toán_học]' (Ví dụ: [GRAPH: sin(x)] hoặc [GRAPH: x^2 - 4]) để hệ thống tự động vẽ đồ thị."
            )

            if run_ai_prompt_safe_func is not None:
                with st.spinner("🚀 Trợ lý AI đang tổng hợp kiến thức và thiết kế ma trận, đề thi..."):
                    res_text, status = run_ai_prompt_safe_func(prompt_exam)
                    
                    if status == "error":
                        st.error(f"❌ Máy chủ AI hoặc API Key phản hồi sự cố: {res_text}")
                    else:
                        st.session_state["current_exam_designer_output"] = res_text
                        
                        # Tự động đóng gói lưu trữ lưu lại lịch sử đề thi của trường
                        exam_id = f"De_{mon_de.replace(' ', '_')}_{khoi_de.replace(' ', '_')}_{len(st.session_state['db_de_kiem_tra'])+1}"
                        st.session_state["db_de_kiem_tra"].append({
                            "id": exam_id,
                            "mon": mon_de,
                            "khoi": khoi_de,
                            "hinh_thuc": hinh_thuc,
                            "data": res_text
                        })
                        st.success(f"🎉 Khởi tạo đề kiểm tra thành công bằng mô hình {status}!")
            else:
                st.error("❌ Lỗi luồng: Chưa kết nối được trình điều khiển AI tổng từ file app.py.")

        # --- KHU VỰC HIỂN THỊ ĐỀ THI VÀ TẢI VỀ ---
        if st.session_state["current_exam_designer_output"]:
            st.markdown("---")
            st.markdown("### 📝 Nội dung Ma trận & Đề kiểm tra vừa sinh:")
            st.markdown(st.session_state["current_exam_designer_output"])
            
            # Xuất phôi Word lưới kẻ bảng dày dặn
            word_exam_data = export_to_docx_vietnam_standard(
                st.session_state["current_exam_designer_output"], 
                f"ĐỀ KIỂM TRA MÔN {mon_de.upper()} - {khoi_de.upper()}"
            )
            st.download_button(
                label="📥 Tải file Word (.docx) Đề kiểm tra chuẩn Quốc hiệu về máy",
                data=word_exam_data,
                file_name=f"De_Kiem_Tra_{mon_de.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # --- KHU VỰC THƯ MỤC ĐỀ ĐÃ XÂY DỰNG (LỊCH SỬ KHO LƯU TRỮ) ---
    with tab_kho_luu_tru:
        st.markdown("#### 📂 Danh mục lưu trữ đề kiểm tra toàn trường")
        if not st.session_state["db_de_kiem_tra"]:
            st.caption("ℹ️ Chưa có đề kiểm tra nào được tạo. Thầy hãy thiết kế đề thi mới ở tab bên cạnh.")
        else:
            indices_to_delete = []
            for idx, de_info in enumerate(st.session_state["db_de_kiem_tra"]):
                with st.expander(f"📋 Đề môn {de_info['mon']} - {de_info['khoi']} ({de_info['hinh_thuc']})"):
                    st.markdown(de_info['data'])
                    
                    h_col1, h_col2 = st.columns(2)
                    word_hist_data = export_to_docx_vietnam_standard(de_info['data'], f"ĐỀ KIỂM TRA MÔN {de_info['mon'].upper()} - {de_info['khoi'].upper()}")
                    h_col1.download_button(
                        label="📥 Tải file Word",
                        data=word_hist_data,
                        file_name=f"De_{de_info['mon'].replace(' ', '_')}.docx",
                        key=f"dl_exam_word_{de_info['id']}",
                        use_container_width=True
                    )
                    
                    if h_col2.button("❌ Xóa đề thi khỏi kho lưu trữ", key=f"del_exam_btn_{de_info['id']}", use_container_width=True):
                        indices_to_delete.append(idx)
                        
            if indices_to_delete:
                for index in sorted(indices_to_delete, reverse=True):
                    st.session_state["db_de_kiem_tra"].pop(index)
                st.success("🗑️ Đã xóa đề thi khỏi danh mục kho lưu trữ!")
                st.rerun()
