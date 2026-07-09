import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Cấu hình
SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY' # Thầy nhớ kiểm tra lại ID này

import json
import tempfile

def get_sheet():
    # 1. Lấy dữ liệu từ secrets
    creds_dict = dict(st.secrets["GOOGLE_KEY"])
    
    # 2. Tạo một tệp tạm thời trong bộ nhớ để lưu JSON
    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as tmp:
        json.dump(creds_dict, tmp)
        tmp_path = tmp.name
    
    # 3. Sử dụng đường dẫn tệp tạm thời này để xác thực
    # Vì tệp này nằm trong bộ nhớ máy chủ, gspread sẽ đọc được mà không báo lỗi stream
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(tmp_path, scope)
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
