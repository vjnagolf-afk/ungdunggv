import streamlit as st
import io
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
from google import genai
from datetime import datetime

# =========================================================
# CẤU HÌNH GOOGLE SHEETS
# =========================================================
SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY' # <--- DÁN ID VÀO ĐÂY
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

def get_sheet():
    # Thay vì dùng from_json_keyfile_name (đọc file), ta dùng from_json_keyfile_dict (đọc từ secrets)
    creds_dict = st.secrets["GOOGLE_KEY"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet("STEM_Projects")

# =========================================================
# HÀM XỬ LÝ WORD & DỮ LIỆU
# =========================================================
def create_word_file(title, content):
    doc = Document()
    doc.add_heading(title, 0)
    # Loại bỏ dấu # và ** để file Word sạch sẽ
    clean_content = re.sub(r'^#+\s*', '', content, flags=re.MULTILINE)
    clean_content = clean_content.replace('**', '').replace('*', '')
    doc.add_paragraph(clean_content)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =========================================================
# GIAO DIỆN 3 THẺ (Logic đã tích hợp Sheets)
# =========================================================
def render_tab_1():
    st.info("💡 **THẺ 1:** Chọn tiêu chí để AI gợi ý dự án.")
    # (Giữ logic cũ của thầy ở đây)

def render_tab_2():
    st.success("🛠️ **THẺ 2:** Soạn KHBD.")
    ten_chu_de_t2 = st.text_input("Tên dự án:", key="ten_t2")
    
    if st.button("🚀 KÍCH HOẠT AI BIÊN SOẠN KHBD"):
        # ... (GỌI API GEMINI NHƯ CŨ) ...
        st.session_state.stem_generated_content = "Kết quả từ AI..." # Giả lập

    if st.session_state.stem_generated_content:
        if st.button("💾 LƯU VÀO GOOGLE SHEETS"):
            try:
                sheet = get_sheet()
                ngay_luu = datetime.now().strftime("%Y-%m-%d")
                sheet.append_row([ten_chu_de_t2, st.session_state.stem_generated_content, ngay_luu])
                st.toast("Đã lưu thành công lên Drive!", icon="✅")
            except Exception as e:
                st.error(f"Lỗi lưu Sheets: {e}")

def render_tab_3():
    st.warning("📁 **THẺ 3:** Đọc dữ liệu từ Google Sheets.")
    if st.button("🔄 Tải lại dữ liệu"):
        try:
            sheet = get_sheet()
            records = sheet.get_all_records()
            st.session_state.stem_saved_projects = {r['Ten_Du_An']: r['Noi_Dung'] for r in records}
        except Exception as e:
            st.error(f"Lỗi đọc Sheets: {e}")
    
    for ten, nd in st.session_state.stem_saved_projects.items():
        with st.expander(f"📌 {ten}"):
            st.markdown(nd)

def render_stem_section():
    # Thêm các dòng khởi tạo này vào ngay đầu hàm
    if "stem_saved_projects" not in st.session_state: 
        st.session_state.stem_saved_projects = {}
    if "stem_generated_content" not in st.session_state: 
        st.session_state.stem_generated_content = ""
    
    st.markdown("## 🚀 HỆ SINH THÁI GIÁO DỤC STEM")
    
    tab1, tab2, tab3 = st.tabs(["💡 1. SẢN PHẨM", "🛠️ 2. XÂY DỰNG KHBD", "📁 3. KHBD ĐÃ LƯU"])
    with tab1: render_tab_1()
    with tab2: render_tab_2()
    with tab3: render_tab_3()
