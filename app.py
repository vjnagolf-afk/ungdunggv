import streamlit as st
import io
import json
import os
import pandas as pd
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from pypdf import PdfReader
# Cấu hình kết nối Google Sheets
@st.cache_resource # Lệnh này giúp ứng dụng không phải đăng nhập lại mỗi khi thầy nhấn nút
def get_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('key_google.json', scope)
    return gspread.authorize(creds)

client = get_client()
# ==============================================================================
# 1. CẤU HÌNH TRANG VÀ THIẾT LẬP SIÊU CSS ĐỊNH DẠNG THANH BÊN THEO BIỂU MẪU CHUẨN
# ==============================================================================
st.set_page_config(
    page_title="EduAssist AI 2.0 - Hệ Sinh Thái Giáo Viên Số",
    page_icon="🚀",
    layout="wide"
)
# Giao diện tiêu đề chính
st.title("📚 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")
st.caption("Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk")
st.markdown("---")
st.markdown("""
<style>

/* Đẩy nội dung thanh bên lên sát đỉnh đầu một cách vừa vặn */
div[data-testid="stSidebarUserContent"] {
    padding-top: 1.0rem !important;
    padding-bottom: 0px !important;
}
/* Rút bớt lề của đường kẻ phân cách để hai khối nội dung sát lại nhau hơn */
div[data-testid="stSidebarUserContent"] hr {
    margin-top: 0.6rem !important;
    margin-bottom: 0.6rem !important;
}
/* Thiết lập khoảng cách dọc giữa các Widget ở mức vừa phải, đẹp mắt */
div[data-testid="stSidebarUserContent"] div[data-testid="stVerticalBlock"] > div {
    margin-bottom: 0px !important;
}
/* Ép nhãn tiêu đề CHỌN PHÂN HỆ TÁC NGHIỆP in đậm, căn giữa 100%, màu đỏ rực rỡ */
div[data-testid="stSidebarUserContent"] div[data-testid="stRadio"] > label {
    font-weight: 900 !important;
    color: #E11D48 !important;
    text-align: center !important;
    display: block !important;
    width: 100% !important;
    font-size: 26px !important;
    letter-spacing: 0.5px !important;
    border-bottom: 2px dashed #FDA4AF !important;
    padding-bottom: 6px !important;
    margin-bottom: 12px !important;
}
/* Làm nổi bật tuyệt đối 2 dòng chữ tùy chọn phân hệ: Chữ siêu đậm, tăng size */
div[data-testid="stSidebarUserContent"] div[data-testid="stRadio"] div[data-baseweb="radio"] span {
    font-weight: 900 !important;
    font-size: 26px !important;
}
/* Gán màu sắc đặc trưng nổi bật riêng cho từng mục tác nghiệp chuyên sâu */
div[data-testid="stSidebarUserContent"] div[data-testid="stRadio"] div[data-baseweb="radio"]:nth-child(1) span {
    color: #1D4ED8 !important; /* Xanh Navy đậm nổi bật */
}
div[data-testid="stSidebarUserContent"] div[data-testid="stRadio"] div[data-baseweb="radio"]:nth-child(2) span {
    color: #EA580C !important; /* Cam hổ phách rực rỡ tương phản */
}
/* Hiệu ứng rê chuột mượt mà cho tất cả các nút bấm */
div.stButton > button, div.stDownloadButton > button {
    transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
    border-radius: 8px !important;
}
div.stButton > button:hover, div.stDownloadButton > button:hover {
    background-color: #1E40AF !important; 
    color: #FFFFFF !important; 
    border-color: #3B82F6 !important; 
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 15px rgba(59, 130, 246, 0.4) !important;
}
</style>
""", unsafe_allow_html=True)
# ==============================================================================
# 2. BỘ ENGINE AI & KẾ XUẤT VĂN BẢN CHUẨN ĐỊNH DẠNG HÀNH CHÍNH (TIMES NEW ROMAN)
# ==============================================================================
def run_ai_prompt_safe(prompt_text, api_key):
    models_to_try = ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.5-pro"]
    client = genai.Client(api_key=api_key)
    last_error = None
    for model in range(len(models_to_try)):
        for attempt in range(2):
            try:
                response = client.models.generate_content(model=models_to_try[model], contents=prompt_text)
                return response.text, models_to_try[model]
            except Exception as e:
                last_error = e
                time.sleep(2)
    raise Exception(f"Tất cả các mô hình AI đều đang gặp lỗi: {last_error}")

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
    
    # SỬA TRIỆT ĐỂ LỖI THEO ẢNH TRACEBACK: Định vị chính xác qua hàng đầu tiên rows[0].cells [1]
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
    
    for line in text_content.split('\n'):
        cleaned_line = line.strip()
        if not cleaned_line: continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if cleaned_line.startswith('#'):
            p.add_run(cleaned_line.lstrip('#').strip()).bold = True
        else:
            p.add_run(cleaned_line)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def convert_df_to_excel_bytes(df, sheet_name='Sheet1'):
    val_output = io.BytesIO()
    with pd.ExcelWriter(val_output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return val_output.getvalue()
def generate_dynamic_5512_docx():
    doc = Document()
    title = doc.add_paragraph()
    title.add_run("KHUNG KẾ HOẠCH BÀI DẠY MẪU CHUẨN 5512").bold = True
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

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
# ==============================================================================
# HÀM PHỤ TRỢ: LƯU TRỮ VÀ ĐỌC HỒ SƠ TỪ TỆP TIN CỤC BỘ ĐỂ CHỐNG RESET KHI TẢI LẠI TRANG [1]
# ==============================================================================
def save_data_to_local_file(file_name, data):
    try:
        with open(file_name, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        pass

def load_data_from_local_file(file_name, default_data):
    if os.path.exists(file_name):
        try:
            with open(file_name, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            return default_data
    return default_data

FILE_THANH_VIEN = "db_thanh_vien.json"
FILE_THANH_TICH = "db_thanh_tich_da_nam.json"
DANH_SACH_NAM_HOC = [f"{nam} - {nam+1}" for nam in range(2020, 2036)]

MAU_THANH_VIEN = [
    {"STT": 1, "Họ và tên": "Lê Hồng Dưỡng", "Chức vụ": "Tổ trưởng", "Phân môn chính": "KHTN (Lý) - CN", "Email": "vjnagolf@gmail.com", "Số điện thoại": "0984331178"},
    {"STT": 2, "Họ và tên": "Nguyễn Thị Huyền Trang", "Chức vụ": "Tổ viên", "Phân môn chính": "KHTN (Lý) - CN", "Email": "binhnt@gmail.com", "Số điện thoại": "0987654321"},
    {"STT": 3, "Họ và tên": "Lê Hùng Cường", "Chức vụ": "Tổ viên", "Phân môn chính": "KHTN (Lý) - CN", "Email": "viettt@gmail.com", "Số điện thoại": "0905123456"}
]

MAU_THANH_TICH = {
    "2025 - 2026": [
        {"Năm học": "2025 - 2026", "Họ và tên": "Lê Hồng Dưỡng", "Đánh giá viên chức": "Hoàn thành xuất sắc (HTSX)", "Kết quả BD HSG": "1 giải Nhất môn Vật lý", "Kết quả NCKH": "Đạt giải Nhì", "Kết quả STEM": "02 bài học xuất sắc", "Kết quả Sáng kiến": "Giải B", "Thi GVDG": "Cấp Tỉnh", "Thi GVCNG": "Không", "Kết quả TDTT": "Giải Nhất bóng bàn", "Kết quả VN": "Giải xuất sắc", "Hiến máu nhân đạo": "01 lần", "Hoạt động khác": "Tích cực Chuyển đổi số"},
        {"Năm học": "2025 - 2026", "Họ và tên": "Nguyễn Thị Huyền Trang", "Đánh giá viên chức": "Hoàn thành tốt (HTT)", "Kết quả BD HSG": "2 giải Ba môn Sinh", "Kết quả NCKH": "Cấp Trường", "Kết quả STEM": "01 chuyên đề Tốt", "Kết quả Sáng kiến": "Cấp Huyện", "Thi GVDG": "Cấp Huyện", "Thi GVCNG": "Không", "Kết quả TDTT": "Đội trưởng Bóng chuyền", "Kết quả VN": "Biên đạo giải Nhất", "Hiến máu nhân đạo": "01 lần", "Hoạt động khác": "Công đoàn viên xuất sắc"}
    ]
}

if "db_thanh_vien" not in st.session_state:
    st.session_state["db_thanh_vien"] = load_data_from_local_file(FILE_THANH_VIEN, MAU_THANH_VIEN)

if "db_thanh_tich_da_nam" not in st.session_state:
    st.session_state["db_thanh_tich_da_nam"] = load_data_from_local_file(FILE_THANH_TICH, MAU_THANH_TICH)

# Hàm đồng bộ an toàn chặn lỗi TypeError ép kiểu trống [1]
def sync_sub_databases():
    current_pc = st.session_state.get("db_phan_cong_hien_tai", [])
    pc_mapped = {x["Họ và tên"]: x for x in current_pc}
    new_pc = []
    
    for idx, tv in enumerate(st.session_state["db_thanh_vien"]):
        name = tv.get("Họ và tên", "Giáo viên mới")
        try:
            stt_val = int(tv["STT"]) if (pd.notna(tv.get("STT")) and str(tv.get("STT")).strip() != "") else (idx + 1)
        except:
            stt_val = idx + 1
            
        if name in pc_mapped:
            old_item = pc_mapped[name]
            new_pc.append({
                "STT": stt_val, "Họ và tên": name, "Phân môn chính": tv.get("Phân môn chính", ""),
                "Lớp phân công": old_item.get("Lớp phân công", "9A1, 9A2"), "Số tiết phân công": old_item.get("Số tiết phân công", 8),
                "Nhiệm vụ kiêm nhiệm": tv.get("Chức vụ", "")
            })
        else:
            new_pc.append({
                "STT": stt_val, "Họ và tên": name, "Phân môn chính": tv.get("Phân môn chính", ""),
                "Lớp phân công": "", "Số tiết phân công": 0, "Nhiệm vụ kiêm nhiệm": tv.get("Chức vụ", "")
            })
    st.session_state["db_phan_cong_hien_tai"] = new_pc
    # Khóa logic chuẩn: Tab thành tích lịch sử thi đua (Tab 3) biệt lập 100%, không bị đồng bộ đè [1]

if "db_phan_cong_hien_tai" not in st.session_state:
    sync_sub_databases()

if "db_de_kiem_tra" not in st.session_state:
    st.session_state["db_de_kiem_tra"] = [
        {"ten_de": "Đề giữa kỳ I - KHTN 9", "mon": "Khoa học tự nhiên", "khoi": "Lớp 9", "noi_dung": "### MA TRẬN ĐỀ THI ĐÃ DỰNG"}
    ]

has_secrets = "GEMINI_API_KEY" in st.secrets if hasattr(st, "secrets") else False
api_key_input = st.secrets["GEMINI_API_KEY"] if has_secrets else ""
st.sidebar.title("MENU HỆ THỐNG")
phan_he_lam_viec = st.sidebar.radio("CHỌN PHÂN HỆ TÁC NGHIỆP", [" Trợ lý Giảng dạy (Giáo viên)", " Trợ lý Quản lý (Tổ chuyên môn)"], index=0)

if phan_he_lam_viec == " Trợ lý Giảng dạy (Giáo viên)":
    chuc_nang_chinh = st.sidebar.selectbox("CHỌN NỘI DUNG THỰC HIỆN", ["1. Thiết kế KHBD thông minh", "2. Thiết kế Đề KTĐG (Ma trận - Đặc tả)", "3. Đánh giá học sinh"])
else:
    chuc_nang_chinh = st.sidebar.selectbox(
        "QUẢN LÝ TỔ CHUYÊN MÔN", 
        [
            "1. Hệ thống Quản lý và Phân công chuyên môn giảng dạy", 
            "2. Xây dựng Biên bản sinh hoạt tổ chuyên môn định kỳ",
            "3. Xây dựng Kế hoạch Giáo dục cá nhân (Phụ lục III - Công văn 5512)",
            "4. Thống kê số liệu tổ"
        ]
    )

st.sidebar.markdown("---")
st.sidebar.title("🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
user_role = st.sidebar.selectbox("VỚI TƯ CÁCH LÀ:", ["Giáo viên / tổ viên", "Tổ trưởng chuyên môn (Admin)"], index=1)
is_admin_choice = (user_role == "Tổ trưởng chuyên môn (Admin)")

if "is_admin_verified" not in st.session_state:
    st.session_state["is_admin_verified"] = False

if is_admin_choice:
    admin_password = st.sidebar.text_input("Nhập mã pin quản trị Admin:", type="password", key="pwd_admin_root")
    if admin_password == "021476":
        st.session_state["is_admin_verified"] = True
        st.sidebar.success("✅ Xác thực thành công! Quyền Admin đã mở.")
    else:
        st.session_state["is_admin_verified"] = False
        if admin_password: st.sidebar.error("❌ Sai mã pin bảo mật!")
        else: st.sidebar.info("💡 Điền mã pin để kích hoạt quyền sửa.")
else:
    st.session_state["is_admin_verified"] = False
    st.sidebar.info("💡 Bạn được xem hồ sơ tổ")

if not has_secrets:
    api_key_input = st.sidebar.text_input("Nhập khóa Gemini API Key nếu cần dùng AI:", type="password", value=api_key_input)

st.sidebar.markdown("---")
tac_gia = st.sidebar.text_input("Họ và tên tác giả:", value="Lê Hồng Dưỡng")
don_vi = st.sidebar.text_input("Đơn vị công tác:", value="Trường THCS Nguyễn Chí Thanh")
# Giao diện chính phân hệ Giáo viên
if phan_he_lam_viec == " Trợ lý Giảng dạy (Giáo viên)":
    if chuc_nang_chinh == "1. Thiết kế KHBD thông minh":
        st.header("📋 THIẾT KẾ KẾ HOẠCH BÀI DẠY (KHBD) TÍCH HỢP NĂNG LỰC SỐ")
        st.markdown("##### Kho tài liệu hướng dẫn xây dựng KHBD")
        st.info("Thầy cô có thể tải trực tiếp tài liệu hướng dẫn cấu trúc khung kế hoạch bài dạy mẫu 5512 thế hệ mới:")
        
        col_w, col_t = st.columns(2)
        with col_w: st.download_button(label="📥 Tải Khung KHBD chuẩn (.docx)", data=generate_dynamic_5512_docx(), file_name="Khung_Ke_hoach_Bai_day_5512.docx")
        with col_t: st.download_button(label="📥 Tải Cẩm nang sư phạm (.txt)", data="CẨM NANG SƯ PHẠM 5512 NÂNG CAO".encode('utf-8'), file_name="Cam_nang_Huong_dan_5512.txt")
        
        st.markdown("---")
        col_input1, col_input2 = st.columns(2)
        with col_input1:
            ten_bai = st.text_input("Tên bài học học tập:", placeholder="Ví dụ: Bài 1: Qua Đèo Ngang")
            lop = st.selectbox("Khối lớp giảng dạy:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9"], index=3)
        with col_input2:
            mon_hoc = st.text_input("Môn học chuyên môn:", value="Khoa học tự nhiên")
            thoi_luong = st.text_input("Thời lượng thực hiện:", value="1 tiết")
            
        uploaded_khbd_file = st.file_uploader("Tải file học liệu gốc/Nội dung SGK nguồn (.pdf, .docx):", type=["pdf", "docx"], key="khbd_up")
        khbd_content = ""
        if uploaded_khbd_file:
            khbd_content = read_uploaded_docx(uploaded_khbd_file) if uploaded_khbd_file.name.endswith('.docx') else read_uploaded_pdf(uploaded_khbd_file)
            st.success("✅ Đã nạp thành công học liệu gốc thành công vào bộ nhớ AI.")
            
        uu_tien_tai_lieu = st.checkbox("🎯 ƯU TIÊN BÁM SÁT 100% TÀI LIỆU TẢI LÊN", value=True)
        if st.button("🚀 XD KHBD thông minh"):
            with st.spinner("AI đang dựng giáo án..."):
                try:
                    res, _ = run_ai_prompt_safe(f"Soạn giáo án chuẩn 5512 cho bài {ten_bai}", api_key_input)
                    st.markdown(res)
                except Exception as e: st.error(f"Lỗi: {e}")
    elif chuc_nang_chinh == "2. Thiết kế Đề KTĐG (Ma trận - Đặc tả)":
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
                tong_so_tn = st.slider("Tổng số câu trắc nghiệm khách quan (Tối đa 32 câu):", min_value=0, max_value=32, value=16)
            with col_sl2:
                tong_so_tl = st.number_input("Tổng số câu hỏi tự luận tùy chọn (Không giới hạn):", min_value=0, value=2, step=1)
                
            st.markdown("##### Phân bổ chi tiết hình thức câu hỏi trắc nghiệm")
            col_hb1, col_sl_cs2, col_sl_cs3, col_sl_cs4 = st.columns(4)
            with col_hb1:
                tn_1_dap_an = st.number_input("Trắc nghiệm 1 đáp án đúng:", min_value=0, value=10, step=1)
            with col_sl_cs2:
                tn_dung_sai = st.number_input("Trắc nghiệm Đúng/Sai:", min_value=0, value=2, step=1)
            with col_sl_cs3:
                tn_dien_khuyen = st.number_input("Trắc nghiệm điền khuyết:", min_value=0, value=2, step=1)
            with col_sl_cs4:
                tn_tra_loi_ngan = st.number_input("Trắc nghiệm trả lời ngắn:", min_value=0, value=2, step=1)
                
            tong_phan_bo_thuc_te = tn_1_dap_an + tn_dung_sai + tn_dien_khuyen + tn_tra_loi_ngan
            if tong_phan_bo_thuc_te == tong_so_tn:
                st.success(f"✅ Đồng bộ thành công! Tổng phân bổ ({tong_phan_bo_thuc_te} câu) khớp với cấu hình đề thi.")
            else:
                st.warning(f"⚠️ Chưa đồng bộ! Tổng phân bổ chi tiết ({tong_phan_bo_thuc_te} câu) chưa khớp với Tổng số câu trên thanh trượt ({tong_so_tn} câu).")
                
            uploaded_excel_de = st.file_uploader("Tải lên file giới hạn kiến thức hoặc bộ đề thi làm căn cứ (.pdf, .docx):", type=["pdf", "docx"], key="up_file_de_thi")
            content_de_nguon = ""
            if uploaded_excel_de:
                content_de_nguon = read_uploaded_docx(uploaded_excel_de) if uploaded_excel_de.name.endswith('.docx') else read_uploaded_pdf(uploaded_excel_de)
                st.success("✅ Đã nạp thành công tài liệu kiến thức nguồn vào bộ nhớ AI.")
                
            uu_tien_de = st.checkbox("🎯 ƯU TIÊN BÁM SÁT 100% TÀI LIỆU/ĐỀ MẪU TẢI LÊN", value=True)
            
            if st.button("🛠️ Thiết lập Đề thi + Ma trận & Đặc tả"):
                if not api_key_input: st.error("Thầy cần cấu hình Gemini API Key tại thanh bên!")
                else:
                    with st.spinner("AI đang dựng đề..."):
                        try:
                            prompt_de = f"Thiết kế đề thi môn {mon_de} lớp {khoi_de}. Số câu trắc nghiệm: {tong_so_tn}, tự luận: {tong_so_tl}."
                            result_text, _ = run_ai_prompt_safe(prompt_de, api_key_input)
                            st.markdown(result_text)
                            st.session_state["db_de_kiem_tra"].append({"ten_de": ten_de if ten_de else "Đề tự động dựng", "mon": mon_de, "khoi": khoi_de, "noi_dung": result_text})
                            st.success("✅ Đã lưu bộ đề thi thành công!")
                        except Exception as error_ai: st.error(f"Lỗi hệ thống AI: {error_ai}")
                            
        with tab_kho_luu_tru:
            st.subheader("📁 Thư mục lưu trữ đề kiểm tra nội bộ đã dựng")
            for idx, item in enumerate(st.session_state["db_de_kiem_tra"]):
                with st.expander(f"📋 {item['ten_de']} - Môn: {item['mon']} ({item['khoi']})"):
                    st.markdown(item["noi_dung"])
                    st.download_button(label="📥 Tải lại File Word", data=export_to_docx_vietnam_standard(item["noi_dung"], item["ten_de"]), file_name=f"{item['ten_de']}_{idx}.docx", key=f"dl_de_thi_{idx}")

    elif chuc_nang_chinh == "3. Đánh giá học sinh":
        st.header("📋 BỘ CÔNG CỤ ĐÁNH GIÁ, NHẬN XÉT VÀ HỖ TRỢ HỌC SINH")
        chu_de = st.text_input("Nội dung đánh giá:", value="Thầy/Cô điền tên bài học/Chủ đề cần xây dựng tiêu chí đánh giá vào đây")
        if st.button("🚀 Sinh học liệu"):
            with st.spinner("AI đang tạo..."):
                try:
                    res, _ = run_ai_prompt_safe(f"Tạo tiêu chí rubric cho chủ đề {chu_de}", api_key_input)
                    st.markdown(res)
                except Exception as error_rb: st.error(f"Lỗi: {error_rb}")
else:
    # PHÂN HỆ: QUẢN LÝ TỔ CHUYÊN MÔN
    if chuc_nang_chinh == "1. Hệ thống Quản lý và Phân công chuyên môn giảng dạy":
        st.subheader("📋 HỆ THỐNG QUẢN LÝ VÀ PHÂN CÔNG CHUYÊN MÔN GIẢNG DẠY")
        tab_thanh_vien, tab_phan_cong, tab_thanh_tich = st.tabs(["👥 DANH SÁCH THÀNH VIÊN & CHUYÊN MÔN", "📅 THEO DÕI PHÂN CÔNG CHUYÊN MÔN", "📈 GHI CHÉP THÀNH TÍCH CỦA GIÁO VIÊN"])
        
        with tab_thanh_vien:
            st.markdown("##### Giao diện quản lý danh sách thành viên tổ")
            df_thanh_vien = pd.DataFrame(st.session_state["db_thanh_vien"])
            df_tv_edited = st.data_editor(df_thanh_vien, num_rows="dynamic" if st.session_state["is_admin_verified"] else "fixed", disabled=not st.session_state["is_admin_verified"], use_container_width=True, key="editor_thanh_vien")
            
            if st.session_state["is_admin_verified"] and st.button("💾 Lưu cập nhật Danh sách & Tự động Đồng bộ các Tab"):
                st.session_state["db_thanh_vien"] = df_tv_edited.to_dict('records')
                save_data_to_local_file(FILE_THANH_VIEN, st.session_state["db_thanh_vien"])
                sync_sub_databases() 
                st.success("✅ Đã cập nhật danh sách tổ viên hiện hành thành công!")
                st.rerun()

        with tab_phan_cong:
            st.markdown("##### Bảng phân công chuyên môn hiện hành (Tự động bộ danh sách nhân sự)")
            df_pc_current = pd.DataFrame(st.session_state["db_phan_cong_hien_tai"])
            df_pc_edited = st.data_editor(df_pc_current, num_rows="fixed", disabled=not st.session_state["is_admin_verified"], use_container_width=True, key="editor_phan_cong_active")
            if st.session_state["is_admin_verified"] and st.button("💾 Lưu dữ liệu Phân công giảng dạy"):
                st.session_state["db_phan_cong_hien_tai"] = df_pc_edited.to_dict('records')
                st.success("✅ Đã lưu phân công giảng dạy thành công!")
                st.rerun()

        with tab_thanh_tich:
            st.markdown("##### Bảng theo dõi Thành tích Giáo viên qua các niên khóa")
            nam_hoc_selected = st.selectbox("📅 CHỌN NĂM HỌC CẦN QUẢN LÝ VÀ LƯU TRỮ HỒ SƠ THI ĐUA:", DANH_SACH_NAM_HOC, index=5)
            st.markdown(f"###### Dữ liệu thi đua năm học đang chọn: <span style='color:#1E40AF; font-weight:bold;'>{nam_hoc_selected}</span>", unsafe_allow_html=True)
            
            if nam_hoc_selected not in st.session_state["db_thanh_tich_da_nam"] or not st.session_state["db_thanh_tich_da_nam"][nam_hoc_selected]:
                st.session_state["db_thanh_tich_da_nam"][nam_hoc_selected] = [
                    {
                        "Năm học": nam_hoc_selected, "Họ và tên": tv["Họ và tên"], "Đánh giá viên chức": "Hoàn thành tốt (HTT)",
                        "Kết quả BD HSG": "Không", "Kết quả NCKH": "Không", "Kết quả STEM": "Không", "Kết quả Sáng kiến": "Không",
                        "Thi GVDG": "Không", "Thi GVCNG": "Không", "Kết quả TDTT": "Không", "Kết quả VN": "Không", "Hiến máu nhân đạo": "Không", "Hoạt động khác": "Không"
                    } for tv in st.session_state["db_thanh_vien"]
                ]
            
            list_tt_data = st.session_state["db_thanh_tich_da_nam"][nam_hoc_selected]
            for row in list_tt_data:
                row["Năm học"] = nam_hoc_selected
                
            df_tt_current = pd.DataFrame(list_tt_data)
            df_tt_edited = st.data_editor(df_tt_current, num_rows="dynamic" if st.session_state["is_admin_verified"] else "fixed", disabled=not st.session_state["is_admin_verified"], use_container_width=True, key=f"editor_thanh_tich_{nam_hoc_selected}")
            
            if st.session_state["is_admin_verified"]:
                col_save_tt, col_ex_tt = st.columns(2)
                with col_save_tt:
                    if st.button(f"💾 Lưu dữ liệu Thi đua riêng cho năm {nam_hoc_selected}"):
                        saved_records = df_tt_edited.to_dict('records')
                        for record in saved_records:
                            record["Năm học"] = nam_hoc_selected
                        st.session_state["db_thanh_tich_da_nam"][nam_hoc_selected] = saved_records
                        save_data_to_local_file(FILE_THANH_TICH, st.session_state["db_thanh_tich_da_nam"]) 
                        st.success(f"✅ Hệ thống đã lưu khóa vĩnh viễn dữ liệu thi đua năm học {nam_hoc_selected} biệt lập!")
                        st.rerun()
                with col_ex_tt:
                    st.download_button(label=f"📥 Tải Bảng thi đua {nam_hoc_selected} về máy (.xlsx)", data=convert_df_to_excel_bytes(df_tt_current), file_name=f"Thi_Dua_Nam_Hoc_{nam_hoc_selected.replace(' - ', '_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.markdown("---")
                uploaded_excel_tt = st.file_uploader(f"Tải tệp Excel kết quả thi đua của năm học {nam_hoc_selected} (.xlsx):", type=["xlsx"], key=f"up_excel_tt_{nam_hoc_selected}")
                if uploaded_excel_tt:
                    try:
                        df_uploaded_tt = pd.read_excel(uploaded_excel_tt)
                        required_tt_cols = ["Năm học", "Họ và tên", "Đánh giá viên chức", "Kết quả BD HSG", "Kết quả NCKH", "Kết quả STEM", "Kết quả Sáng kiến", "Thi GVDG", "Thi GVCNG", "Kết quả TDTT", "Kết quả VN", "Hiến máu nhân đạo", "Hoạt động khác"]
                        for col in required_tt_cols:
                            if col not in df_uploaded_tt.columns: df_uploaded_tt[col] = "Không"
                        df_uploaded_tt["Năm học"] = nam_hoc_selected
                        st.session_state["db_thanh_tich_da_nam"][nam_hoc_selected] = df_uploaded_tt[required_tt_cols].to_dict('records')
                        save_data_to_local_file(FILE_THANH_TICH, st.session_state["db_thanh_tich_da_nam"])
                        st.success(f"🎉 Đã nạp và lưu vĩnh viễn tệp Excel thi đua niên khóa {nam_hoc_selected}!")
                        st.rerun()
                    except Exception as e: st.error(f"Lỗi file Excel: {e}")
            else:
                st.warning(f"🔒 Chế độ xem: Bạn chỉ được xem tổng hợp thi đua năm {nam_hoc_selected}. Chỉ Tổ trưởng mở khóa mã pin mới có quyền sửa đổi.")
    elif chuc_nang_chinh == "2. Xây dựng Biên bản sinh hoạt tổ chuyên môn định kỳ":
        st.subheader("📝 XÂY DỰNG BIÊN BẢN SINH HOẠT TỔ CHUYÊN MÔN")
        st.markdown("#### I. THỜI GIAN, ĐỊA ĐIỂM, THÀNH PHẦN HOẠT ĐỘNG")
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            ky_hop_so = st.text_input("Kỳ họp thứ:", value="05")
            thang_hop_so = st.text_input("Biên bản họp tháng:", value="01")
            nam_hoc_chuoi = st.text_input("Năm học trường đăng ký:", value="2025 - 2026")
        with col_t2:
            gio_phut_hop = st.text_input("Vào lúc mấy giờ mấy phút:", value="13 giờ 30 phút")
            ngay_thang_hop = st.text_input("Ngày/Tháng/Năm cuộc họp diễn ra:", value="17/01/2026")
            phong_hop = st.text_input("Địa điểm tại phòng học:", value="Hội trường")
        with col_t3:
            chu_tri_cuoc = st.text_input("Chủ trì cuộc họp (Chủ tọa):", value="Thầy Lê Hồng Dưỡng – Tổ trưởng")
            thu_ky_cuoc = st.text_input("Thư ký lập biên bản:", value="Cô Nguyễn Thị Bình")
            thanh_phan_co = st.text_input("Thành phần (Có mặt / Vắng mặt):", value="Có mặt: 10/10; Vắng mặt: 0")

        st.markdown("---")
        st.markdown("#### II. NỘI DUNG CUỘC HỌP & NẠP KẾ HOẠCH NGUỒN")
        uploaded_kh_thang = st.file_uploader("📥 Tải file Kế hoạch tháng làm căn cứ biên bản (.pdf, .docx):", type=["pdf", "docx"], key="up_kh_thang")
        content_kh_thang = ""
        if uploaded_kh_thang:
            content_kh_thang = read_uploaded_docx(uploaded_kh_thang) if uploaded_kh_thang.name.endswith('.docx') else read_uploaded_pdf(uploaded_kh_thang)
            st.success("✅ Đã nạp thành công kế hoạch tháng vào bộ nhớ trợ lý ảo.")

        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            bo_sung_11 = st.text_area("Ghi chú bổ sung mục 1.1 (Công tác trọng tâm):", placeholder="Nhập thêm lưu ý nếu có...")
        with col_m2:
            bo_sung_12 = st.text_area("Ghi chú bổ sung mục 1.2 (Thảo luận chuyên môn):", placeholder="Nhập lưu ý phương pháp, chuyên đề STEM...")
        with col_m3:
            bo_sung_13 = st.text_area("Ghi chú bổ sung mục 1.3 (Công tác khác):", placeholder="Nhập việc phong trào, công đoàn...")

        y_kien_dong_gop = st.text_area("Mục 2. Ý kiến đóng góp của các thành viên tổ chuyên môn:", value="- Các thành viên thảo luận sôi nổi và hoàn toàn nhất trí với phương hướng phân bổ chuyên môn tháng này.")

        if st.button("🤖 AI Tự động soạn thảo Biên bản"):
            if not uploaded_kh_thang: st.warning("⚠️ Thầy cần đính kèm file Kế hoạch tháng lên trước!")
            else:
                with st.spinner("AI đang dựng biên bản hành chính..."):
                    try:
                        prompt_bb = f"Hãy đóng vai thư ký tổ chuyên môn trường THCS Nguyễn Chí Thanh. Dựa vào nội dung file kế hoạch: {content_kh_thang}, soạn thảo nội dung phần II. NỘI DUNG CUỘC HỌP bóc tách lắp vào mục 1.1, 1.2, 1.3 phát triển chi tiết."
                        res_ai, _ = run_ai_prompt_safe(prompt_bb, api_key_input)
                        full_bien_ban_text = f"BIÊN BẢN SINH HOẠT TỔ CHUYÊN MÔN\\nKỲ HỌP THỨ {ky_hop_so} THÁNG {thang_hop_so}\\n\\nI. THỜI GIAN, ĐỊA ĐIỂM\\nLúc {gio_phut_hop}, ngày {ngay_thang_hop} tại {phong_hop}\\nChủ trì: {chu_tri_cuoc}, Thư ký: {thu_ky_cuoc}, Thành phần: {thanh_phan_co}\\n\\nII. NỘI DUNG CUỘC HỌP\\n{res_ai}\\n\\n2. Ý KIẾN ĐÓNG GÓP\\n{y_kien_dong_gop}"
                        st.markdown(full_bien_ban_text)
                        st.download_button("📥 Xuất văn bản biên bản Word (.docx)", data=export_to_docx_vietnam_standard(full_bien_ban_text, f"Bien_ban_thang_{thang_hop_so}"), file_name=f"Bien_ban_thang_{thang_hop_so}.docx")
                    except Exception as e: st.error(f"Lỗi: {e}")

    elif chuc_nang_chinh == "3. Xây dựng Kế hoạch Giáo dục cá nhân (Phụ lục III - Công văn 5512)":
        st.header("📋 KẾ HOẠCH GIÁO DỤC CÁ NHÂN CỦA GIÁO VIÊN (Phụ lục III)")
        gv_selected = st.selectbox("Chọn Giáo viên cần lập kế hoạch:", [x["Họ và tên"] for x in st.session_state["db_thanh_vien"]])
        mon_tu_dong = next((x["Phân môn chính"] for x in st.session_state["db_thanh_vien"] if x["Họ và tên"] == gv_selected), "")
        st.info(f"Phân môn giảng dạy: **{mon_tu_dong}**")
        if st.button("🪄 AI Thiết kế Phụ lục III chuẩn 5512"):
            with st.spinner("AI đang xử lý..."):
                try:
                    res, _ = run_ai_prompt_safe(f"Thiết kế khung kế hoạch giáo dục cá nhân phụ lục III 5512 cho môn {mon_tu_dong}", api_key_input)
                    st.markdown(res)
                    st.download_button("📥 Tải Phụ lục III Word (.docx)", data=export_to_docx_vietnam_standard(res, f"Phu_luc_III_{gv_selected}"), file_name=f"Phu_luc_III_{gv_selected}.docx")
                except Exception as e: st.error(f"Lỗi: {e}")

    elif chuc_nang_chinh == "4. Thống kê số liệu tổ":
        st.header("📊 THỐNG KÊ GIÁO VIÊN TỔ")
        tong_nhan_su = len(st.session_state["db_thanh_vien"])
        st.metric("Tổng số thành viên tổ", f"{tong_nhan_su} Giáo viên")
        df_tv_current = pd.DataFrame(st.session_state["db_thanh_vien"])
        if "Phân môn chính" in df_tv_current.columns:
            st.bar_chart(df_tv_current["Phân môn chính"].value_counts())
