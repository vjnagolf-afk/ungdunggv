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

# HÀM MỚI: TỰ ĐỘNG VẼ ĐỒ THỊ TỪ BIỂU THỨC TOÁN HỌC VÀ TRẢ VỀ FILE ẢNH
def generate_plot_stream(eq_str):
    fig, ax = plt.subplots(figsize=(5, 3.5))
    x = np.linspace(-10, 10, 400)
    
    # Xử lý các cú pháp toán học cơ bản để an toàn khi chạy eval
    safe_dict = {
        "x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan, "sqrt": np.sqrt
    }
    try:
        # Chuyển đổi dấu ^ thành ** theo chuẩn Python nếu AI lỡ viết sai
        eq_str_py = eq_str.replace('^', '**')
        y = eval(eq_str_py, {"__builtins__": {}}, safe_dict)
        
        # Nếu kết quả là 1 số (đường thẳng), biến nó thành mảng
        if isinstance(y, (int, float)):
            y = np.full_like(x, y)

        ax.plot(x, y, color='#1E40AF', linewidth=2.5)
        ax.axhline(0, color='black', linewidth=1.2)
        ax.axvline(0, color='black', linewidth=1.2)
        ax.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)
        
        # Giới hạn trục y để hình không bị kéo giãn quá mức
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

        # ========================================================
        # PHÁT HIỆN THẺ [GRAPH: ...] VÀ CHÈN HÌNH ẢNH VÀO WORD
        # ========================================================
        if '[GRAPH:' in cleaned_line:
            match = re.search(r'\[GRAPH:\s*(.+?)\]', cleaned_line)
            if match:
                eq = match.group(1)
                img_stream = generate_plot_stream(eq)
                doc.add_picture(img_stream, width=Inches(3.5))
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                continue # Bỏ qua dòng chữ [GRAPH: ...], thay bằng hình ảnh
            
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
            khoi_de = c_sel3.selectbox("Khoi", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"], index=3)
            
            c_lbl4, c_txt4 = st.columns([1, 2])
            c_lbl4.markdown("<div style='margin-top: 8px;'>Thời gian:</div>", unsafe_allow_html=True)
            thoi_gian_de = c_txt4.text_input("Thoi_gian", value="45 phút")

        with col_top2:
            st.markdown("<div style='margin-top: 8px;'>TẢI TÀI LIỆU LÊN (Giới hạn kiến thức/Đề cương):</div>", unsafe_allow_html=True)
            uploaded_files_de = st.file_uploader(
                "Up_Files", 
                type=["pdf", "docx"], 
                accept_multiple_files=True
            )
            if not uploaded_files_de:
                st.markdown("*Chưa có tài liệu nào được tải lên hệ thống.*", unsafe_allow_html=True)
            else:
                st.success(f"✅ Đã tải lên {len(uploaded_files_de)} tài liệu.")

        st.markdown("<hr style='margin: 10px 0px;'>", unsafe_allow_html=True)

        col_tn, spacer, col_tl = st.columns([10, 1, 10])
        
        with col_tn:
            st.markdown("<div class='header-pink'>PHẦN TRẮC NGHIỆM</div>", unsafe_allow_html=True)
            
            c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
            c1.markdown("<b style='color:#C62828; line-height:2.2;'>Tổng số câu TNKQ:</b>", unsafe_allow_html=True)
            ph_tong_so_tn = c2.empty() 
            c3.markdown("<b style='line-height:2.2;'>Tổng điểm TN:</b>", unsafe_allow_html=True)
            ph_tong_diem_tn = c4.empty()
            
            c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
            c1.markdown("<div style='line-height:2.2;'>Số câu nhiều lựa chọn:</div>", unsafe_allow_html=True)
            tn_1_dap_an = c2.number_input("TN_1_DA", min_value=0, value=12)
            c3.markdown("<div style='line-height:2.2;'>Tổng điểm dòng này:</div>", unsafe_allow_html=True)
            diem_tn_1 = c4.number_input("Diem_TN_1", min_value=0.0, value=3.0, step=0.25, format="%.2f")
            
            c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
            c1.markdown("<div style='line-height:2.2;'>Số câu đúng sai:</div>", unsafe_allow_html=True)
            tn_dung_sai = c2.number_input("TN_DS", min_value=0, value=2)
            c3.markdown("<div style='line-height:2.2;'>Tổng điểm dòng này:</div>", unsafe_allow_html=True)
            diem_tn_2 = c4.number_input("Diem_TN_2", min_value=0.0, value=1.0, step=0.25, format="%.2f")

            c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
            c1.markdown("<div style='line-height:2.2;'>Số câu điền khuyết:</div>", unsafe_allow_html=True)
            tn_dien_khuyen = c2.number_input("TN_DK", min_value=0, value=0)
            c3.markdown("<div style='line-height:2.2;'>Tổng điểm dòng này:</div>", unsafe_allow_html=True)
            diem_tn_3 = c4.number_input("Diem_TN_3", min_value=0.0, value=0.0, step=0.25, format="%.2f")

            c1, c2, c3, c4 = st.columns([3, 2, 3, 2])
            c1.markdown("<div style='line-height:2.2;'>Số câu trả lời ngắn:</div>", unsafe_allow_html=True)
            tn_tra_loi_ngan = c2.number_input("TN_TLN", min_value=0, value=0)
            c3.markdown("<div style='line-height:2.2;'>Tổng điểm dòng này:</div>", unsafe_allow_html=True)
            diem_tn_4 = c4.number_input("Diem_TN_4", min_value=0.0, value=0.0, step=0.25, format="%.2f")

            tong_so_tn = tn_1_dap_an + tn_dung_sai + tn_dien_khuyen + tn_tra_loi_ngan
            tong_diem_tn = diem_tn_1 + diem_tn_2 + diem_tn_3 + diem_tn_4
            
            ph_tong_so_tn.text_input("Lock_TS_TN", value=str(tong_so_tn), disabled=True)
            ph_tong_diem_tn.text_input("Lock_TD_TN", value=f"{tong_diem_tn:.2f}", disabled=True)

        with col_tl:
            st.markdown("<div class='header-green'>PHẦN TỰ LUẬN</div>", unsafe_allow_html=True)
            
            c1, c2, c3, c4 = st.columns([3, 2, 2, 2])
            c1.markdown("<b style='color:#1565C0; line-height:2.2;'>TỔNG SỐ CÂU TỰ LUẬN:</b>", unsafe_allow_html=True)
            tong_so_tl = c2.number_input("Tong_TL", min_value=0, max_value=20, value=5)
            
            c3.markdown("<b style='line-height:2.2;'>ĐIỂM TỔNG:</b>", unsafe_allow_html=True)
            ph_tong_diem_tl = c4.empty()
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            diem_tl_list = []
            tong_diem_tl_auto = 0.0
            
            for i in range(int(tong_so_tl)):
                rc1, rc2, rc3, rc4 = st.columns([1, 2, 2, 2])
                rc2.markdown(f"<div style='line-height:2.2;'>Câu {i+1}</div>", unsafe_allow_html=True)
                diem_cau = rc3.number_input(f"Diem_Cau_{i+1}", min_value=0.0, value=1.0, step=0.25, format="%.2f", key=f"diem_tl_{i}")
                rc4.markdown("<div style='line-height:2.2;'>ĐIỂM</div>", unsafe_allow_html=True)
                diem_tl_list.append(diem_cau)
                tong_diem_tl_auto += diem_cau 
                
            ph_tong_diem_tl.text_input("Lock_TD_TL", value=f"{tong_diem_tl_auto:.2f}", disabled=True)

        st.markdown("<hr style='margin: 15px 0px;'>", unsafe_allow_html=True)

        c_btn, c_chk = st.columns([2, 5])
        btn_tao = c_btn.button("⚙ Tự động tạo ma trận & đề thi", type="primary", use_container_width=True)
        uu_tien_de = c_chk.checkbox("Yêu cầu bám sát kiến thức trong tài liệu tải lên", value=True)

        st.markdown("<b>Tỷ lệ mức độ nhận thức (%):</b>", unsafe_allow_html=True)
        c_nb1, c_nb2, c_th1, c_th2, c_vd1, c_vd2, c_vdc1, c_vdc2 = st.columns([1,1,1,1,1,1,1,1])
        c_nb1.markdown("<div style='line-height:2.2;'>Nhận biết:</div>", unsafe_allow_html=True)
        nb = c_nb2.number_input("NB", value=40)
        c_th1.markdown("<div style='line-height:2.2;'>Thông hiểu:</div>", unsafe_allow_html=True)
        th = c_th2.number_input("TH", value=30)
        c_vd1.markdown("<div style='line-height:2.2;'>Vận dụng:</div>", unsafe_allow_html=True)
        vd = c_vd2.number_input("VD", value=20)
        c_vdc1.markdown("<div style='line-height:2.2;'>Vận dụng cao:</div>", unsafe_allow_html=True)
        vdc = c_vdc2.number_input("VDC", value=10)

        st.markdown("<div style='margin-top: 8px;'>Nhập yêu cầu khác (Tùy chọn):</div>", unsafe_allow_html=True)
        yeu_cau_khac = st.text_area("Yeu_Cau_Khac", placeholder="Nhập yêu cầu khác ....")

        if btn_tao:
            if not api_key_input: 
                st.error("Thầy cần cấu hình Gemini API Key tại thanh bên!")
            else:
                with st.spinner("Hệ thống đang phân tích tài liệu và cấu trúc để sinh Ma trận, Đề thi & Đồ thị..."):
                    try:
                        content_de_nguon = ""
                        if uploaded_files_de:
                            for file in uploaded_files_de:
                                content_de_nguon += f"\n--- TÀI LIỆU: {file.name} ---\n"
                                if file.name.endswith('.docx'): content_de_nguon += read_uploaded_docx(file)
                                else: content_de_nguon += read_uploaded_pdf(file)
                        
                        diem_tl_str = ", ".join([f"Câu {i+1} ({diem_tl_list[i]} điểm)" for i in range(int(tong_so_tl))])

                        prompt_de = f"""Đóng vai một chuyên gia khảo thí. Hãy thiết kế Đề kiểm tra môn {mon_de} {khoi_de}. Hình thức: {hinh_thuc}.
Thời gian: {thoi_gian_de}.
Cấu trúc điểm (Tỷ lệ {nb}-{th}-{vd}-{vdc}):
- TRẮC NGHIỆM ({tong_so_tn} câu - {tong_diem_tn} điểm):
  + {tn_1_dap_an} câu nhiều lựa chọn ({diem_tn_1} điểm)
  + {tn_dung_sai} câu đúng/sai ({diem_tn_2} điểm)
  + {tn_dien_khuyen} câu điền khuyết ({diem_tn_3} điểm)
  + {tn_tra_loi_ngan} câu trả lời ngắn ({diem_tn_4} điểm)
- TỰ LUẬN ({tong_so_tl} câu - {tong_diem_tl_auto} điểm). Điểm chi tiết: {diem_tl_str}.
Yêu cầu khác: {yeu_cau_khac}
"""
                        if uu_tien_de and content_de_nguon:
                            prompt_de += f"\n\nBẮT BUỘC BÁM SÁT 100% KIẾN THỨC TÀI LIỆU SAU ĐÂY:\n{content_de_nguon}"

                        prompt_de += """\n
LƯU Ý ĐỊNH DẠNG CÔNG THỨC TOÁN/LÝ/HÓA ĐỂ XUẤT RA WORD KHÔNG BỊ LỖI:
1. TUYỆT ĐỐI KHÔNG DÙNG ký hiệu LaTeX ($ hay $$).
2. Sử dụng các ký tự Toán học Unicode chuẩn (VD: √, ½, ², ³, Δ, π, α, β, ➔). Viết phân số dưới dạng ngang (Ví dụ: (2h)/g hoặc g/(2v_0^2)).
3. Sử dụng thẻ <sub>...</sub> và <sup>...</sup> cho chỉ số (VD: H<sub>2</sub>O, x<sup>2</sup>, v<sub>0</sub>).
4. QUAN TRỌNG: NẾU CẦN VẼ ĐỒ THỊ (HÀM SỐ, QUỸ ĐẠO...), bạn BẮT BUỘC phải chèn một dòng duy nhất theo cú pháp: [GRAPH: biểu_thức_python]
   Ví dụ: Quỹ đạo parabol y = x^2 / 80, hãy chèn dòng này vào bài: [GRAPH: x**2 / 80]
   Ví dụ: Hàm số bậc nhất y = 2x + 1, hãy chèn dòng này: [GRAPH: 2*x + 1]
   (Viết chuẩn cú pháp toán học Python, dùng biến x).
   Về Hình học/Vector: AI không thể vẽ ảnh, hãy để lại dòng ghi chú [DÀNH KHOẢNG TRỐNG TẠI ĐÂY ĐỂ CHÈN HÌNH MINH HỌA].

TRÌNH BÀY ĐẦY ĐỦ (Dùng bảng Markdown |---|---| cho Ma trận và Đặc tả):
PHẦN 1. MA TRẬN ĐỀ KIỂM TRA
PHẦN 2. BẢNG ĐẶC TẢ CHI TIẾT
PHẦN 3. ĐỀ KIỂM TRA & ĐÁP ÁN"""

                        result_text, _ = run_ai_prompt_safe_func(prompt_de, api_key_input)
                        
                        st.session_state["db_de_kiem_tra"].append({"ten_de": f"Đề {mon_de} - {khoi_de} ({thoi_gian_de})", "mon": mon_de, "khoi": khoi_de, "noi_dung": result_text})
                        st.success("🎉 Đã tạo đề thi thành công! Thầy/Cô vui lòng chuyển sang thẻ 📂 THƯ MỤC ĐỀ ĐÃ XÂY DỰNG để xem, thu gọn đồ thị và tải file Word.")
                    except Exception as error_ai: 
                        st.error(f"Lỗi hệ thống AI: {error_ai}")

        st.markdown("<div class='footer-red'>© Bản quyền thuộc về Tác giả: Lê Hồng Dưỡng | Đơn vị: Trường THCS Nguyễn Chí Thanh – phường Tân Lập - tỉnh Đắk Lắk</div>", unsafe_allow_html=True)

    with tab_kho_luu_tru:
        st.subheader("📂 Các đề kiểm tra đã được AI tự động sinh và lưu trữ")
        if not st.session_state["db_de_kiem_tra"]:
            st.info("💡 Chưa có đề kiểm tra nào được tạo. Thầy hãy thiết kế đề thi mới ở tab bên cạnh.")
        else:
            for idx, item in enumerate(reversed(st.session_state["db_de_kiem_tra"])):
                real_idx = len(st.session_state["db_de_kiem_tra"]) - 1 - idx
                with st.expander(f"📋 {item['ten_de']} (Bấm để xem/thu gọn nội dung)"):
                    # Render trực quan ảnh đồ thị ngay trên web trước khi tải về Word
                    hien_thi_web = item["noi_dung"]
                    # Xóa chữ [GRAPH: ...] trên màn hình web để không bị thô (do ảnh chỉ chèn vào Word)
                    # Nếu muốn hiển thị cả trên web, Streamlit cần parse riêng. Hiện tại, ta báo cho GV biết.
                    hien_thi_web = re.sub(r'\[GRAPH:\s*(.+?)\]', r'*[Hệ thống sẽ tự động vẽ đồ thị \1 vào file Word khi tải về]*', hien_thi_web)
                    st.markdown(hien_thi_web, unsafe_allow_html=True)
                    
                    st.markdown("<hr>", unsafe_allow_html=True)
                    col_bt1, col_bt2 = st.columns([1, 1])
                    with col_bt1:
                        st.download_button(
                            label="📥 Tải File Word (Word Docx Kèm Đồ Thị)", 
                            data=export_to_docx_vietnam_standard(item["noi_dung"], item["ten_de"]), 
                            file_name=f"{item['ten_de']}_{real_idx}.docx", 
                            key=f"dl_de_thi_{real_idx}", 
                            use_container_width=True,
                            type="primary"
                        )
                    with col_bt2:
                        if st.button("🗑️ Xóa đề kiểm tra này khỏi kho lưu trữ", key=f"del_de_thi_{real_idx}", use_container_width=True):
                            st.session_state["db_de_kiem_tra"].pop(real_idx)
                            st.rerun()
