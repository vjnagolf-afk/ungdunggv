import streamlit as st
import gspread
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard

# --- (Giữ nguyên toàn bộ hàm sync_exam_to_google_sheet) ---
SPREADSHEET_ID = "1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY"
TAB_NAME = "DE_KT"

def sync_exam_to_google_sheet(ten_de, mon, khoi, thoi_gian, noi_dung):
    try:
        creds_dict = None
        priority_keys = ["gspread_credentials", "GSPREAD_CREDENTIALS", "google_sheet_creds", "google_creds", "GOOGLE_KEY"]
        for key in priority_keys:
            if key in st.secrets: creds_dict = st.secrets[key]; break
        if creds_dict is None:
            for key in st.secrets.keys():
                node = st.secrets[key]
                if hasattr(node, "get") or isinstance(node, dict):
                    if node.get("type") == "service_account" or "private_key" in node: creds_dict = node; break
        if creds_dict is None: return False
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_key(SPREADSHEET_ID)
        worksheet = sh.worksheet(TAB_NAME)
        if len(worksheet.get_all_values()) == 0:
            worksheet.append_row(["Tên Đề", "Môn Học", "Khối Lớp", "Thời Gian", "Nội Dung Đề Thi"])
        worksheet.append_row([ten_de, mon, khoi, thoi_gian, noi_dung])
        return True
    except: return False

def render_exam_designer_section(run_ai_prompt_safe_func):
    if "ket_qua_de" not in st.session_state: st.session_state["ket_qua_de"] = ""
    if "model_dung" not in st.session_state: st.session_state["model_dung"] = ""
    
    # Init states
    if "save_ten_de" not in st.session_state: st.session_state["save_ten_de"] = "Đề kiểm tra giữa học kỳ I"
    if "save_school" not in st.session_state: st.session_state["save_school"] = "TRƯỜNG THCS NGUYỄN CHÍ THANH"
    if "save_mon_hoc" not in st.session_state: st.session_state["save_mon_hoc"] = "Khoa học tự nhiên"
    if "save_khoi_lop" not in st.session_state: st.session_state["save_khoi_lop"] = "Lớp 8"
    if "save_thoi_gian" not in st.session_state: st.session_state["save_thoi_gian"] = "45 phút"

    tab_tao_de, tab_thu_muc = st.tabs(["🔴 CHỨC NĂNG TẠO ĐỀ KIỂM TRA AI", "🗂️ THƯ MỤC ĐỀ ĐÃ XÂY DỰNG"])
    
    with tab_tao_de:
        # Bố cục hàng 1: Thông tin cấu hình và Tài liệu
        col_cfg, col_files = st.columns([1, 1])
        with col_cfg:
            hinh_thuc = st.selectbox("Hình thức đề:", ["Trắc nghiệm kết hợp tự luận", "100% Trắc nghiệm", "100% Tự luận"])
            st.session_state["save_mon_hoc"] = st.text_input("Môn học:", value=st.session_state["save_mon_hoc"])
            st.session_state["save_khoi_lop"] = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Khối 10", "Khối 11", "Khối 12"], index=2)
            st.session_state["save_thoi_gian"] = st.text_input("Thời gian:", value=st.session_state["save_thoi_gian"])
            st.session_state["save_school"] = st.text_input("Tên trường:", value=st.session_state["save_school"])
        with col_files:
            file_de_cuong = st.file_uploader("1. Tải Đề Cương (.docx, .pdf):", type=["docx", "pdf"], key="exam_upload_de_cuong_2026")
            file_ma_tran = st.file_uploader("2. Tải Khung Ma Trận / Đặc tả:", type=["docx", "pdf"], key="exam_upload_ma_tran_2026")
            
        # Xử lý file (giữ nguyên logic)
        noi_dung_de_cuong = read_uploaded_docx(file_de_cuong) if file_de_cuong and file_de_cuong.name.endswith(".docx") else (read_uploaded_pdf(file_de_cuong) if file_de_cuong else "")
        noi_dung_ma_tran = read_uploaded_docx(file_ma_tran) if file_ma_tran and file_ma_tran.name.endswith(".docx") else (read_uploaded_pdf(file_ma_tran) if file_ma_tran else "")

        st.divider()
        
        # Bố cục phần Trắc nghiệm & Tự luận
        c_tn, c_tl = st.columns(2)
        with c_tn:
            with st.container(border=True):
                st.markdown("#### 🔴 PHẦN TRẮC NGHIỆM")
                sc_nlc = st.number_input("Số câu nhiều lựa chọn:", 0, 50, 12)
                sc_ds = st.number_input("Số câu đúng sai:", 0, 50, 2)
                sc_dk = st.number_input("Số câu điền khuyết:", 0, 50, 0)
                sc_tln = st.number_input("Số câu trả lời ngắn:", 0, 50, 0)
                tong_diem_tn = st.number_input("Tổng điểm TN:", value=4.00, step=0.25)
        with c_tl:
            with st.container(border=True):
                st.markdown("#### 🟢 PHẦN TỰ LUẬN")
                sc_tl = st.number_input("TỔNG SỐ CÂU TỰ LUẬN:", 0, 20, 3)
                tong_diem_tl = st.number_input("ĐIỂM TỔNG TỰ LUẬN:", value=6.00, step=0.25)
                with st.expander("Chỉ định điểm chi tiết từng câu"):
                    diem_ct = [st.number_input(f"Câu {i+1}", value=2.0) for i in range(sc_tl)]

        # Nhận thức & Yêu cầu cuối
        st.markdown("**Tỷ lệ mức độ nhận thức (%):**")
        cols_tyle = st.columns(4)
        tl_nb = cols_tyle[0].number_input("Nhận biết:", 0, 100, 40, 10)
        tl_th = cols_tyle[1].number_input("Thông hiểu:", 0, 100, 30, 10)
        tl_vd = cols_tyle[2].number_input("Vận dụng:", 0, 100, 20, 10)
        tl_vdc = cols_tyle[3].number_input("Vận dụng cao:", 0, 100, 10, 10)

        st.session_state["save_ten_de"] = st.text_input("Tên bài kiểm tra / Đề số:", value=st.session_state["save_ten_de"])
        yeu_cau_khac = st.text_area("Yêu cầu khác:")
        
        # Nút bấm và logic gọi AI (Giữ nguyên phần gọi `run_ai_prompt_safe_func` và logic tạo Word như cũ)
        if st.button("🔴 Tự động khởi tạo ma trận và đề thi", type="primary", use_container_width=True):
             # [Dán logic xử lý AI của bạn vào đây]
             pass

    with tab_thu_muc:
        st.write("📂 Danh sách các đề kiểm tra đã đồng bộ:")
        st.markdown(f"🔗 [Bấm vào đây để mở trực tiếp Google Sheets](https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit)")
