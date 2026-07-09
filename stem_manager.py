import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Cấu hình scope
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
SHEET_ID = '1QhX2fP520f9xXoP2p3W5i2123456789' # Thầy thay đúng ID của thầy vào đây

def get_sheet():
    creds_dict = dict(st.secrets["GOOGLE_KEY"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPE)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet("STEM_Projects")

def save_to_sheets(ten_du_an, noi_dung):
    try:
        sheet = get_sheet()
        ngay_luu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([str(ten_du_an), str(noi_dung), str(ngay_luu)])
        return True
    except Exception as e:
        st.error(f"Lỗi: {e}")
        return False

def render_tab_2():
    st.success("🛠️ THẺ 2: Soạn KHBD.")
    if "stem_generated_content" not in st.session_state:
        st.session_state.stem_generated_content = ""
    
    ten_chu_de_t2 = st.text_input("Tên dự án:", key="ten_t2")
    if st.button("🚀 KÍCH HOẠT AI"):
        st.session_state.stem_generated_content = "Nội dung dự án STEM của thầy..."
    
    if st.session_state.stem_generated_content:
        st.write(st.session_state.stem_generated_content)
        if st.button("💾 LƯU VÀO GOOGLE SHEETS"):
            save_to_sheets(ten_chu_de_t2, st.session_state.stem_generated_content)

def render_stem_section():
    st.markdown("## 🚀 HỆ SINH THÁI GIÁO DỤC STEM")
    tab1, tab2, tab3 = st.tabs(["💡 1. SẢN PHẨM", "🛠️ 2. XÂY DỰNG KHBD", "📁 3. KHBD ĐÃ LƯU"])
    with tab2: render_tab_2()
