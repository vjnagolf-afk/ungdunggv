import streamlit as st
import io
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import matplotlib.pyplot as plt
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ================= CẤU HÌNH GOOGLE SHEETS =================
SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY'

def get_dekt_sheet():
    creds_dict = dict(st.secrets["GOOGLE_KEY"])
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet("DE_KT")

# --- CÁC HÀM XỬ LÝ (GIỮ NGUYÊN) ---
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
        ax.text(0.5, 0.5, f"[Không thể vẽ đồ thị]", ha='center', va='center', color='red')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
    buf.seek(0)
    plt.close(fig)
    return buf

def export_to_docx_vietnam_standard(text_content, title_name, school_name="TRƯỜNG THCS NGUYỄN CHÍ THANH", group_name="TỔ KHOA HỌC TỰ NHIÊN - GDTC"):
    # (Hàm giữ nguyên của thầy)
    doc = Document()
    # ... [Giữ nguyên logic của thầy] ...
    # Để tiết kiệm không gian chat, thầy vẫn giữ nguyên đoạn này trong file của thầy nhé
    return doc.save(io.BytesIO()) # Chỉ là minh họa, code của thầy vẫn dùng đúng hàm này

def render_exam_designer_section(api_key_input, run_ai_prompt_safe_func):
    # GIỮ NGUYÊN STYLE CSS CỦA THẦY
    st.markdown("""<style> ... </style>""", unsafe_allow_html=True)

    if "db_de_kiem_tra" not in st.session_state: st.session_state["db_de_kiem_tra"] = []
    if "cloud_data_dekt" not in st.session_state: st.session_state["cloud_data_dekt"] = []

    # TAB CỦA THẦY
    tab_thiet_ke, tab_kho_luu_tru = st.tabs(["📝 CHỨC NĂNG: TẠO ĐỀ KIỂM TRA AI", "☁️ KHO ĐÁM MÂY (GOOGLE SHEETS)"])
    
    with tab_thiet_ke:
        # ... [Giữ nguyên bố cục cột của thầy] ...
        # Sau khi chạy AI thành công và có kết quả:
        if run_exam_ai:
            # ... (Logic AI của thầy) ...
            # THÊM NÚT LƯU SAU KHI TẠO XONG:
            if st.button("☁️ Lưu Đồng Bộ Lên Google Sheets", type="primary", use_container_width=True):
                try:
                    sheet = get_dekt_sheet()
                    ten_de = f"Đề {mon_de} - {khoi_de}"
                    ngay_luu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    sheet.append_row([ten_de, mon_de, hinh_thuc, str(st.session_state["current_exam_designer_output"]), ngay_luu])
                    st.success("✅ Đã đồng bộ lên Google Sheets!")
                except Exception as e: st.error(f"Lỗi: {e}")

    with tab_kho_luu_tru:
        st.subheader("📂 KHO ĐÁM MÂY (GOOGLE SHEETS)")
        if st.button("🔄 Tải dữ liệu từ Google Sheets"):
            st.session_state["cloud_data_dekt"] = get_dekt_sheet().get_all_values()
            
        for idx, row in enumerate(st.session_state.get("cloud_data_dekt", [])):
            if len(row) >= 4:
                with st.expander(f"📋 {row[0]}"):
                    st.markdown(row[3])
                    if st.button("❌ Xóa", key=f"del_{idx}"):
                        get_dekt_sheet().delete_row(idx + 1)
                        st.rerun()
