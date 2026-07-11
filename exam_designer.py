# =====================================================================
# PHẦN 1: CẤU HÌNH THƯ VIỆN & GOOGLE SHEETS TRÊN EXAM_DESIGNER.PY
# =====================================================================
import streamlit as st
import gspread
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard

# Khai báo cấu hình Google Sheets của hệ thống
SPREADSHEET_ID = "1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY"
TAB_NAME = "DE_KT"

def get_exam_sheet():
    """Hàm kết nối và xác thực tài khoản dịch vụ Google Sheets an toàn"""
    try:
        creds_dict = None
        priority_keys = ["gspread_credentials", "GSPREAD_CREDENTIALS", "google_sheet_creds", "google_creds", "GOOGLE_KEY"]
        
        # Tìm thông tin xác thực ưu tiên trong Streamlit secrets
        for key in priority_keys:
            if key in st.secrets:
                creds_dict = st.secrets[key]
                break
                
        # Tìm kiếm dự phòng nếu không khớp từ khóa ưu tiên
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
        sh = gc.open_by_key(SPREADSHEET_ID)
        return sh.worksheet(TAB_NAME)
    except Exception:
        return None

def sync_exam_to_google_sheet(ten_de, mon, khoi, thoi_gian, noi_dung):
    """Đồng bộ nội dung đề thi mới tạo lên bảng tính"""
    sheet = get_exam_sheet()
    if sheet:
        try:
            if len(sheet.get_all_values()) == 0:
                sheet.append_row(["Tên Đề", "Môn Học", "Khối Lớp", "Thời Gian", "Nội Dung Đề Thi"])
            sheet.append_row([ten_de, mon, khoi, thoi_gian, noi_dung])
            return True
        except Exception:
            return False
    return False

def get_all_exams_from_sheet():
    """Tải toàn bộ cơ sở dữ liệu danh sách đề thi đã lưu trữ"""
    sheet = get_exam_sheet()
    return sheet.get_all_records() if sheet else []

def delete_exam_from_sheet(row_index):
    """Xóa hàng đề thi được chọn dựa trên chỉ số index thực tế"""
    sheet = get_exam_sheet()
    if sheet:
        try:
            # +2 vì dòng 1 là tiêu đề (header) và index Google Sheet bắt đầu từ 1
            sheet.delete_rows(row_index + 2)
            return True
        except Exception:
            return False
    return False
# =====================================================================
# PHẦN 2: GIAO DIỆN CHÍNH & THIẾT LẬP THÔNG SỐ CẤU CẤU TRÚC ĐỀ THI
# =====================================================================

def render_exam_designer_section(run_ai_prompt_safe_func):
    """Hàm render toàn bộ giao diện thiết kế và quản lý đề thi"""
    
    # Giao diện CSS tùy biến nâng cao cho Streamlit
    st.markdown("""
        <style>
            .block-container {
                max-width: 80% !important;
                padding-top: 1rem !important;
            }
            div[data-testid="stTabs"] button {
                font-size: 18px !important;
                color: red !important;
                font-weight: bold !important;
                text-transform: uppercase !important;
            }
            .stNumberInput input {
                border: 2px solid #FFD700 !important;
                font-weight: bold !important;
                font-size: 14px !important;
                text-align: center !important;
            }
            div[data-testid="stCheckbox"] label span {
                color: red !important;
                font-style: italic !important;
                font-weight: bold !important;
                font-size: 15px !important;
            }
            button[kind="primary"] {
                background-color: #ff3b3b !important;
                border: 1px solid #cc0000 !important;
            }
            button[kind="primary"] * {
                color: #ffffff !important;
                font-weight: bold !important;
                font-size: 16px !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    # Khởi tạo các trạng thái session_state hệ thống an toàn
    if "ket_qua_de" not in st.session_state: st.session_state["ket_qua_de"] = ""
    if "model_dung" not in st.session_state: st.session_state["model_dung"] = ""
    
    if "save_ten_de" not in st.session_state: st.session_state["save_ten_de"] = "Kiểm tra đánh giá giữa kì I"
    if "save_school" not in st.session_state: st.session_state["save_school"] = "TRƯỜNG THCS NGUYỄN CHÍ THANH"
    if "save_mon_hoc" not in st.session_state: st.session_state["save_mon_hoc"] = "Khoa học tự nhiên"
    if "save_khoi_lop" not in st.session_state: st.session_state["save_khoi_lop"] = "Lớp 8"
    if "save_thoi_gian" not in st.session_state: st.session_state["save_thoi_gian"] = "45 phút"

    # Phân chia hai tab chức năng chính
    tab_tao_de, tab_thu_muc = st.tabs(["CHỨC NĂNG TẠO ĐỀ KIỂM TRA", "THƯ MỤC LƯU ĐỀ ĐÃ XD"])
    
    with tab_tao_de:
        # Hàng 1: Menu thông tin chung của đề thi
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("<h6 style='color:blue; font-weight:bold; text-align:center;'>MENU MÔN HỌC</h6>", unsafe_allow_html=True)
            danh_sach_mon = [
                "Ngữ văn", "Toán", "Tiếng Anh", "Giáo dục công dân", "Lịch sử và Địa lí", 
                "Khoa học tự nhiên", "KHTN (Lý)", "KHTN (Hóa)", "KHTN (Sinh)", 
                "Công nghệ", "Tin học", "Nghệ thuật", "Giáo dục địa phương", 
                "Hoạt động trải nghiệm, hướng nghiệp"
            ]
            idx_mon = danh_sach_mon.index(st.session_state["save_mon_hoc"]) if st.session_state["save_mon_hoc"] in danh_sach_mon else 5
            st.session_state["save_mon_hoc"] = st.selectbox("Môn học:", danh_sach_mon, index=idx_mon, label_visibility="collapsed")
        
        with col2:
            st.markdown("<h6 style='color:blue; font-weight:bold; text-align:center;'>MENU LỚP</h6>", unsafe_allow_html=True)
            st.session_state["save_khoi_lop"] = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Khối 10", "Khối 11", "Khối 12"], index=2, label_visibility="collapsed")
        
        with col3:
            st.markdown("<h6 style='color:blue; font-weight:bold; text-align:center;'>HÌNH THỨC KT</h6>", unsafe_allow_html=True)
            hinh_thuc = st.selectbox("Hình thức đề:", ["Trắc nghiệm kết hợp tự luận", "100% Trắc nghiệm", "100% Tự luận"], label_visibility="collapsed")
        
        with col4:
            st.markdown("<h6 style='color:blue; font-weight:bold; text-align:center;'>THỜI GIAN</h6>", unsafe_allow_html=True)
            st.session_state["save_thoi_gian"] = st.text_input("Thời gian:", value=st.session_state["save_thoi_gian"], label_visibility="collapsed")

        # Hàng 2: Tỷ lệ nhận thức ma trận điểm
        st.markdown("<h6 style='color:red; font-weight:bold; margin-top:15px;'>Tỷ lệ mức độ nhận thức (%):</h6>", unsafe_allow_html=True)
        col_nb, col_th, col_vd, col_vdc = st.columns(4)
        with col_nb:
            c_l, c_r = st.columns([1, 1])
            c_l.markdown("<p style='font-weight:bold; font-size:15px; text-align:right; margin-top:5px;'>Nhận biết:</p>", unsafe_allow_html=True)
            tl_nhan_biet = c_r.number_input("nb", min_value=0, max_value=100, value=40, step=10, label_visibility="collapsed")
        with col_th:
            c_l, c_r = st.columns([1, 1])
            c_l.markdown("<p style='font-weight:bold; font-size:15px; text-align:right; margin-top:5px;'>Thông hiểu:</p>", unsafe_allow_html=True)
            tl_thong_hieu = c_r.number_input("th", min_value=0, max_value=100, value=30, step=10, label_visibility="collapsed")
        with col_vd:
            c_l, c_r = st.columns([1, 1])
            c_l.markdown("<p style='font-weight:bold; font-size:15px; text-align:right; margin-top:5px;'>Vận dụng:</p>", unsafe_allow_html=True)
            tl_van_dung = c_r.number_input("vd", min_value=0, max_value=100, value=20, step=10, label_visibility="collapsed")
        with col_vdc:
            c_l, c_r = st.columns([1.2, 1])
            c_l.markdown("<p style='font-weight:bold; font-size:15px; text-align:right; margin-top:5px;'>Vận dụng cao:</p>", unsafe_allow_html=True)
            tl_vd_cao = c_r.number_input("vdc", min_value=0, max_value=100, value=10, step=10, label_visibility="collapsed")

        st.markdown("<br>", unsafe_allow_html=True)

        # Hàng 3: Đặt tên đề thi và Tải tài liệu tham chiếu cấu trúc
        col_name, col_f1, col_f2 = st.columns([2, 1.2, 1.2])
        with col_name:
            st.markdown("<h6 style='color:red; font-weight:bold;'>Tên bài kiểm tra / Đề số:</h6>", unsafe_allow_html=True)
            st.session_state["save_ten_de"] = st.text_input("tende", value=st.session_state["save_ten_de"], label_visibility="collapsed")
    
        with col_f1:
            st.markdown("<p style='color:red; font-style:italic; font-weight:bold; text-align:center; font-size:14px;'>Tải Đề Cương (.docx, .pdf):</p>", unsafe_allow_html=True)
            file_de_cuong = st.file_uploader("f1", type=["docx", "pdf"], key="exam_upload_de_cuong_2026", label_visibility="collapsed")
        with col_f2:
            st.markdown("<p style='color:red; font-style:italic; font-weight:bold; text-align:center; font-size:14px;'>Tải Đề mẫu ma trận (.docx, .pdf):</p>", unsafe_allow_html=True)
            file_ma_tran = st.file_uploader("f2", type=["docx", "pdf"], key="exam_upload_ma_tran_2026", label_visibility="collapsed")

        # Logic đọc dữ liệu từ file tải lên
        noi_dung_de_cuong, noi_dung_ma_tran = "", ""
        if file_de_cuong:
            noi_dung_de_cuong = read_uploaded_docx(file_de_cuong) if file_de_cuong.name.endswith(".docx") else read_uploaded_pdf(file_de_cuong)
            st.success(f" Đã nạp đề cương: {file_de_cuong.name}")
        if file_ma_tran:
            noi_dung_ma_tran = read_uploaded_docx(file_ma_tran) if file_ma_tran.name.endswith(".docx") else read_uploaded_pdf(file_ma_tran)
            st.success(f" Đã nạp ma trận: {file_ma_tran.name}")
# =====================================================================
# PHẦN 3: THIẾT LẬP ĐỊNH LƯỢNG SỐ CÂU HỎI & ĐIỂM CHI TIẾT
# =====================================================================

        st.markdown("<br>", unsafe_allow_html=True)
        col_tn, col_space, col_tl = st.columns([1.1, 0.1, 1.1])
        
        # Cấu hình điểm chi tiết phần Trắc nghiệm
        with col_tn:
            header_tn_placeholder = st.empty()
            
            # Câu trắc nghiệm nhiều lựa chọn
            c1, c2, c3, c4 = st.columns([2.2, 1.2, 1.2, 0.6])
            c1.markdown("<p style='font-size:14px; font-weight:bold; margin-top:8px;'>Số câu nhiều lựa chọn:</p>", unsafe_allow_html=True)
            sc_nhieu_lua_chon = c2.number_input("sc_nlc", min_value=0, value=12, label_visibility="collapsed")
            d_nhieu_lua_chon = c3.number_input("d_nlc", min_value=0.0, value=3.25, step=0.25, label_visibility="collapsed")
            c4.markdown("<p style='font-style:italic; font-size:14px; margin-top:8px;'>điểm</p>", unsafe_allow_html=True)
            
            # Câu trắc nghiệm đúng/sai
            c1, c2, c3, c4 = st.columns([2.2, 1.2, 1.2, 0.6])
            c1.markdown("<p style='font-size:14px; font-weight:bold; margin-top:8px;'>Số câu đúng/sai:</p>", unsafe_allow_html=True)
            sc_dung_sai = c2.number_input("sc_ds", min_value=0, value=1, label_visibility="collapsed")
            d_dung_sai = c3.number_input("d_ds", min_value=0.0, value=0.25, step=0.25, label_visibility="collapsed")
            c4.markdown("<p style='font-style:italic; font-size:14px; margin-top:8px;'>điểm</p>", unsafe_allow_html=True)
            
            # Câu điền khuyết
            c1, c2, c3, c4 = st.columns([2.2, 1.2, 1.2, 0.6])
            c1.markdown("<p style='font-size:14px; font-weight:bold; margin-top:8px;'>Số câu điền khuyết:</p>", unsafe_allow_html=True)
            sc_dien_khuyet = c2.number_input("sc_dk", min_value=0, value=1, label_visibility="collapsed")
            d_dien_khuyet = c3.number_input("d_dk", min_value=0.0, value=0.25, step=0.25, label_visibility="collapsed")
            c4.markdown("<p style='font-style:italic; font-size:14px; margin-top:8px;'>điểm</p>", unsafe_allow_html=True)
            
            # Câu trả lời ngắn
            c1, c2, c3, c4 = st.columns([2.2, 1.2, 1.2, 0.6])
            c1.markdown("<p style='font-size:14px; font-weight:bold; margin-top:8px;'>Số câu trả lời ngắn:</p>", unsafe_allow_html=True)
            sc_tra_loi_ngan = c2.number_input("sc_tln", min_value=0, value=2, label_visibility="collapsed")
            d_tra_loi_ngan = c3.number_input("d_tln", min_value=0.0, value=0.5, step=0.25, label_visibility="collapsed")
            c4.markdown("<p style='font-style:italic; font-size:14px; margin-top:8px;'>điểm</p>", unsafe_allow_html=True)
            
            # Tính toán tổng lượng câu và điểm trắc nghiệm (đã chuẩn hóa tên biến không dấu)
            tong_cau_tn = sc_nhieu_lua_chon + sc_dung_sai + sc_dien_khuyet + sc_tra_loi_ngan
            tong_diem_tn = d_nhieu_lua_chon + d_dung_sai + d_dien_khuyet + d_tra_loi_ngan
            hien_thi_tn = str(round(tong_diem_tn, 2))
            
            header_tn_placeholder.markdown(f"""
                <div style="background-color: #FFF2CC; padding: 10px; display: flex; align-items: center; justify-content: center; gap: 15px;">
                    <span style="color: blue; font-size: 18px; font-weight: bold;">TRẮC NGHIỆM</span>
                    <span style="background-color: white; border: 1px solid #FFD700; color: blue; padding: 4px 45px; font-weight: bold; font-size: 18px;">{hien_thi_tn}</span>
                    <span style="color: blue; font-size: 18px; font-weight: bold;">Điểm</span>
                </div><br>
            """, unsafe_allow_html=True)

        # Cấu hình điểm chi tiết phần Tự luận
        with col_tl:
            h1, h2, h3, h4 = st.columns([1.5, 0.8, 1, 1])
            with h1: st.markdown("<div style='background-color:#E2EEDB; height:45px; line-height:45px; text-align:center; color:blue; font-weight:bold; font-size:18px;'>TỰ LUẬN</div>", unsafe_allow_html=True)
            with h2: sc_tu_luan = st.number_input("sc_tl", min_value=1, max_value=20, value=4, label_visibility="collapsed")
            with h3: ph_tong_diem_tl = st.empty()
            with h4: st.markdown("<div style='background-color:#E2EEDB; height:45px; line-height:45px; text-align:center; color:blue; font-weight:bold; font-size:18px;'>Điểm</div>", unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            tong_diem_tl = 0.0
            diem_chi_tiet_tl = []
            
            # Vòng lặp phân rã phân bổ điểm cho từng câu tự luận độc lập
            for i in range(int(sc_tu_luan)):
                c1, c2, c3 = st.columns([2.5, 1.2, 1.3])
                c1.markdown(f"<p style='font-size:14px; font-weight:bold; text-align:center; margin-top:8px;'>Câu {i+1}.</p>", unsafe_allow_html=True)
                d_c = c2.number_input(f"tl_{i}", min_value=0.0, value=2.75 if i==0 else (1.5 if i==1 else 1.0), step=0.25, label_visibility="collapsed")
                c3.markdown("<p style='font-style:italic; font-size:14px; margin-top:8px;'>điểm</p>", unsafe_allow_html=True)
                tong_diem_tl += d_c
                diem_chi_tiet_tl.append((i+1, d_c))
            
            hien_thi_tl = str(round(tong_diem_tl, 2))
            ph_tong_diem_tl.markdown(f"<div style='background-color:white; border:2px solid #FFD700; height:45px; line-height:40px; text-align:center; color:red; font-weight:bold; font-size:18px;'>{hien_thi_tl}</div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        c_yc, c_chk = st.columns([1, 1])
        with c_yc:
            st.markdown("<p style='color:red; font-weight:bold; text-decoration:underline; font-style:italic; margin-bottom:5px;'>Yêu cầu khác:</p>", unsafe_allow_html=True)
        with c_chk:
            bam_sat = st.checkbox("Bám sát nội dung đề cương tải lên:", value=True)
            
        yeu_cau_khac = st.text_area("Yêu cầu khác", placeholder="Ví dụ: ....", label_visibility="collapsed")
# =====================================================================
# PHẦN 4: PHÁT SINH CẤU TRÚC PROMPT AI & ĐIỀU PHỐI MÔ HÌNH XỬ LÝ
# =====================================================================

        col_btn_l, col_btn_r = st.columns([1.5, 1.0])
        with col_btn_l:
            nut_sinh_de = st.button(" 🚀 TỰ ĐỘNG KHỞI TẠO MA TRẬN VÀ ĐỀ THI", type="primary", use_container_width=True)
        with col_btn_r:
            mo_hinh_uu_tien = st.selectbox("Mô hình xử lý đề thi:", ["3.1 Flash-Lite", "3.5 Flash", "3.1 Pro", "Tư duy mở rộng"], index=0, label_visibility="collapsed")

        # Khởi chạy quy trình AI sinh đề
        if nut_sinh_de:
            if int(tl_nhan_biet + tl_thong_hieu + tl_van_dung + tl_vd_cao) != 100:
                st.error(" ⚠️ Tổng tỷ lệ mức độ nhận thức phải bằng 100%!")
            else:
                with st.spinner(" 🧠 Hệ thống đang lập bảng Ma trận, Bản đặc tả và dựng đề thi..."):
                    chuoi_diem_tl_clean = ", ".join([f"Câu {c_id}. ({d} điểm)" for c_id, d in diem_chi_tiet_tl])
                    lenh_bam_sat = "\n- YÊU CẦU QUAN TRỌNG: Bám sát 100% nội dung đề cương tài liệu tải lên." if bam_sat else ""
                    
                    # Prompt chuyên sâu tối ưu cấu trúc tài liệu xuất ra
                    prompt_vi_mo = f"""
Bạn là Chuyên gia Khảo thí và Kiểm định Chất lượng Giáo dục phổ thông tại Việt Nam.
Hãy thiết kế một tài liệu kiểm tra hoàn chỉnh cho môn {st.session_state['save_mon_hoc']} - {st.session_state['save_khoi_lop']} (Thời gian làm bài: {st.session_state['save_thoi_gian']}).

CẤU TRÚC ĐẦU RA BẮT BUỘC THEO THỨ TỰ (KHÔNG ĐƯỢC BỎ SÓT):

1. KHUNG MA TRẬN ĐỀ KIỂM TRA: Thiết lập dưới dạng bảng Markdown bằng ký tự '|'. Các cột bao gồm: Nội dung kiến thức, Đơn vị kiến thức, Số câu NB, Số câu TH, Số câu VD, Số câu VDC, Tổng số câu, Tỷ lệ (%).
2. BẢNG ĐẶC TẢ KỸ THUẬT ĐỀ KIỂM TRA: Thiết lập dưới dạng bảng Markdown bằng ký tự '|'. Các cột bao gồm: Nội dung kiến thức, Đơn vị kiến thức, Mức độ đánh giá, Số câu hỏi theo ma trận.
3. NỘI DUNG ĐỀ KIỂM TRA CHI TIẾT:
- Hình thức đề: {hinh_thuc}
- Phần Trắc nghiệm ({tong_diem_tn} điểm, {tong_cau_tn} câu): {sc_nhieu_lua_chon} câu nhiều lựa chọn, {sc_dung_sai} câu đúng sai.
- Phần Tự luận ({tong_diem_tl} điểm): Gồm các câu {chuoi_diem_tl_clean}.
LƯU Ý ĐẶC BIỆT: TUYỆT ĐỐI KHÔNG ghi chú mức độ nhận thức (NB, TH, VD) hay tên bài học vào cạnh tên câu hỏi. Trình bày chính xác theo định dạng: "Câu X. (Y điểm): [Nội dung câu hỏi]". Ví dụ: "Câu 1. (2.5 điểm): "
- Tỷ lệ nhận thức: Nhận biết {tl_nhan_biet}%, Thông hiểu {tl_thong_hieu}%, Vận dụng {tl_van_dung}%, Vận dụng cao {tl_vd_cao}%.
- Yêu cầu bổ sung: {yeu_cau_khac} {lenh_bam_sat}

[FILE ĐỀ CƯƠNG KIẾN THỨC GỐC]:
{noi_dung_de_cuong if noi_dung_de_cuong else "Theo phân phối chương trình học chuẩn."}

[FILE MA TRẬN ĐỂ BẮT CHƯỚC (NẾU CÓ)]:
{noi_dung_ma_tran if noi_dung_ma_tran else "Tự sinh dựa theo số lượng câu trên giao diện."}

QUY ĐỊNH ĐỊNH DẠNG TỐI QUAN TRỌNG KHI XUẤT WORD:
- CÔNG THỨC TOÁN, LÝ, HÓA: BẮT BUỘC dùng cú pháp LaTeX chuẩn và bọc trong $...$ (inline) hoặc $$...$$ (display block). Ví dụ: $v = \\frac{{s}}{{t}}$ hoặc $16.67 \\, \\text{{m/s}}$. KHÔNG dùng ký tự văn bản thường cho phân số hay công thức.
- ĐỒ THỊ VÀ BẢNG BIỂU: Nếu đề thi yêu cầu có đồ thị (quãng đường - thời gian, v.v.), hãy sử dụng mã LaTeX (TikZ/pgfplots) đặt trong $$...$$ để hệ thống vẽ đồ thị tự động vào Word, HOẶC thiết lập Bảng số liệu rõ ràng thay vì chỉ viết "(Hình minh họa: ...)". Tuyệt đối không để trống hình vẽ.
- CÁC ĐÁP ÁN TRẮC NGHIỆM: Bắt buộc viết tách dòng độc lập cho A, B, C, D.
"""
                    # Gọi hàm thực thi AI an toàn
                    ket_qua, model_thuc_te = run_ai_prompt_safe_func(prompt_vi_mo, mo_hinh_uu_tien)
                    st.session_state["ket_qua_de"] = ket_qua
                    st.session_state["model_dung"] = model_thuc_te
                    
                    # Tự động đồng bộ lên Google Sheets nếu tạo mới đề thành công
                    if ket_qua and "Lỗi" not in ket_qua:
                        success = sync_exam_to_google_sheet(
                            st.session_state["save_ten_de"],
                            st.session_state["save_mon_hoc"],
                            st.session_state["save_khoi_lop"],
                            st.session_state["save_thoi_gian"],
                            ket_qua
                        )
                        if success:
                            st.toast(" Đã tự động lưu trữ lên Google Sheet!", icon="🚀")
