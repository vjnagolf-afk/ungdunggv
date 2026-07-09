import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY' # Hãy đảm bảo ID này đúng

def get_sheet():
    # 1. Lấy dict trực tiếp từ secrets
    creds_dict = dict(st.secrets["GOOGLE_KEY"])
    
    # 2. Định nghĩa quyền truy cập
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # 3. Sử dụng Credentials hiện đại (Không cần file, không cần luồng stream)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    
    # 4. Ủy quyền cho gspread
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

# --- GIAO DIỆN CÁC THẺ ---
def render_tab_1():
    st.info("💡 Đây là nội dung của Thẻ 1 (Sản phẩm).")
    # Thầy dán lại logic cũ của Thẻ 1 vào đây

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
            if save_to_sheets(ten_chu_de_t2, st.session_state.stem_generated_content):
                st.toast("Đã lưu thành công!", icon="✅")

def render_tab_3():
    st.warning("📁 Đây là nội dung của Thẻ 3 (KHBD đã lưu).")
    # Thầy dán lại logic cũ của Thẻ 3 vào đây

# --- HÀM CHÍNH ---
def render_stem_section():
    st.markdown("## 🚀 HỆ SINH THÁI GIÁO DỤC STEM")
    tab1, tab2, tab3 = st.tabs(["💡 1. SẢN PHẨM", "🛠️ 2. XÂY DỰNG KHBD", "📁 3. KHBD ĐÃ LƯU"])
    with tab1: render_tab_1()
    with tab2: render_tab_2()
    with tab3: render_tab_3()
