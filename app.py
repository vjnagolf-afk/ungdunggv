# app.py - Tích hợp bộ quét IP Mạng thông minh khóa cứng quyền Admin vĩnh viễn cho mọi Tab mới
import streamlit as st
import pandas as pd
import sqlite3
import os
import sys
import urllib.request

# THUẬT TOÁN ĐƯỜNG DẪN: Ép hệ thống tìm module trong cùng thư mục chạy ứng dụng
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from database_manager import inject_demo_data, DB_PATH
from ai_service import run_ai_prompt_safe
from exam_designer import render_exam_designer_section 
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_manager import render_khbd_section  
from danh_gia_manager import render_assessment_section
from org_manager import render_org_section
from bien_ban_manager import render_meeting_minutes
from ke_hoach_ca_nhan_manager import render_personal_plan 
from stem_manager import render_stem_section
from chu_nhiem_manager import render_chu_nhiem_section 
from hskt_plan import render_special_ed_section
from teaching_assistant.main import render_teaching_assistant_section
st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")
# --- THUẬT TOÁN WHITELIST IP: KHÓA CỨNG ĐỊA CHỈ IP CHÍNH CHỦ DỰ ÁN ---
def get_current_public_ip():
    """Tự động quét địa chỉ IP công cộng của thiết bị đang kết nối mạng"""
    try:
        # Gọi dịch vụ kiểm tra IP độc lập của Amazon
        return urllib.request.urlopen('https://amazonaws.com').read().decode('utf-8').strip()
    except:
        return "127.0.0.1"

# 🌟 BƯỚC THIẾT LẬP: Thầy lấy địa chỉ IP mạng máy tính của thầy dán vào danh sách dưới đây
# (Mẹo lấy IP: Thầy mở tab mới gõ "my ip" lên Google, copy dãy số dạng 113.161... dán đè vào đây)
ADMIN_IP_WHITELIST = [
    "113.161.220.105",  # Ví dụ địa chỉ IP mạng Internet trường THCS Nguyễn Chí Thanh
    "14.161.12.34"      # Thầy có thể thêm nhiều IP khác nhau (Ví dụ IP mạng Wifi nhà riêng của thầy)
]
current_device_ip = get_current_public_ip()

# 🚀 TỰ ĐỘNG KHÓA CỨNG QUYỀN ADMIN: Mở tab mới hay F5 chỉ cần trùng IP mạng là tự động nhận diện Admin
if current_device_ip in ADMIN_IP_WHITELIST:
    is_admin_owner = True
else:
    # Dự phòng: Nếu thầy đi công tác mạng khác, thầy vẫn gõ mật mã "123456" vào ô để mở Admin tạm thời
    MAT_MA_ADMIN_CỐ_ĐỊNH = "123456"
    if "current_entered_password" not in st.session_state:
        st.session_state["current_entered_password"] = ""
    url_params = st.query_params
    if url_params.get("admin") == MAT_MA_ADMIN_CỐ_ĐỊNH:
        st.session_state["current_entered_password"] = MAT_MA_ADMIN_CỐ_ĐỊNH
        
    is_admin_owner = (st.session_state["current_entered_password"] == MAT_MA_ADMIN_CỐ_ĐỊNH)

# 🚀 BỘ VÁ CSS: Khử padding/margin thừa giữa các phần tử trong Sidebar
st.markdown("""
    <style>
        [data-testid="stSidebarUserContent"] {
            padding-top: 1.0rem !important;
            padding-bottom: 1.0rem !important;
        }
        [data-testid="stSidebarUserContent"] .stMarkdown, 
        [data-testid="stSidebarUserContent"] .stRadio, 
        [data-testid="stSidebarUserContent"] .stSelectbox {
            margin-bottom: -0.2rem !important;
        }
        [data-testid="stSidebarUserContent"] hr {
            margin-top: 0.5rem !important;
            margin-bottom: 0.5rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# Tiêu đề giao diện chính
st.markdown("<h1 style='text-align: center; color: darkred; font-weight: bold;'>🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #0056b3; font-weight: bold; font-size: 16px;'>Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk</p>", unsafe_allow_html=True)
st.markdown("---")

# ==================================================================================
# --- THANH ĐIỀU HƯỚNG TỔNG TẠI SIDEBAR (VỊ TRÍ 1 - TRÊN CÙNG) ---
# ==================================================================================
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Hỗ trợ giảng dạy", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_root_key_2026_v9"
)
st.sidebar.markdown("---")
# ==================================================================================
# --- KHỐI ĐIỀU HƯỚNG TÁC NGHIỆP CHI TIẾT (3 PHÂN HỆ ĐỘC LẬP) ---
# ==================================================================================

# 1. Phân hệ Trợ lý Giảng dạy (Giáo viên)
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB","6. Thiết kế bài dạy STEM","7. Kế hoạch công tác chủ nhiệm lớp","8. Kế hoạch hỗ trợ học sinh khuyết tật"], label_visibility="collapsed", key="menu_gv_selectbox_v9")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "2. Thiết kế Đề KT": 
        # Thêm dòng này để kiểm tra xem nó có vào được hàm không
        # st.write("Đang tải giao diện thiết kế đề...") 
        from exam_designer import render_exam_designer_section
        from ai_service import run_ai_prompt_safe
        render_exam_designer_section(lambda p, m: run_ai_prompt_safe(p, m, is_admin_owner))
    elif menu == "3. Đánh giá HS": 
        render_assessment_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()
    elif menu == "6. Thiết kế bài dạy STEM":
        render_stem_section()
    elif menu == "7. Kế hoạch công tác chủ nhiệm lớp":
        render_chu_nhiem_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "8. Kế hoạch hỗ trợ học sinh khuyết tật":
        render_special_ed_section(lambda p: run_ai_prompt_safe(p))

# 2. Phân hệ Hỗ trợ giảng dạy (MỚI - ĐỘC LẬP)
elif phan_he == "Hỗ trợ giảng dạy":
    render_teaching_assistant_section()

# 3. Phân hệ Quản lý tổ chuyên môn
elif phan_he == "Trợ lý Quản lý (Tổ chuyên môn)":
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox("Nội dung quản lý", ["1. Quản lý & Phân công chuyên môn", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"], label_visibility="collapsed", key="menu_ql_selectbox_v9")
    
    if menu == "1. Quản lý & Phân công chuyên môn": 
        render_org_section()
    elif menu == "2. Biên bản sinh hoạt": 
        render_meeting_minutes(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "3. Kế hoạch cá nhân": 
        render_personal_plan(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "4. Thống kê số liệu": 
        # ... giữ nguyên phần logic thống kê cũ của bạn ở đây ...
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        # [Chèn phần code xử lý database/dataframe của tổ chuyên môn vào đây]
        
        df_tv = pd.DataFrame()
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                query = "SELECT m.fullname as [Họ và tên], m.main_subject as [Phân môn chính], a.total_periods as [Số tiết/Tuần] FROM org_members m LEFT JOIN org_assignments a ON m.fullname = a.fullname"
                df_tv = pd.read_sql_query(query, conn)
                conn.close()
            except Exception as e:
                st.error(f"⚠️ Lỗi kết nối cơ sở dữ liệu nội bộ: {str(e)}")
        
        thuc_te_co_du_lieu = not df_tv.empty and len(df_tv) > 0 and not (df_tv["Phân môn chính"] == "").all()

        if not thuc_te_co_du_lieu:
            st.warning("ℹ️ Hiện tại chưa có dữ liệu giáo viên nào được nhập từ phân hệ '1. Quản lý & Phân công chuyên môn'.")
            if st.button("💡 Nạp nhanh dữ liệu mẫu để thử nghiệm biểu đồ", type="primary", use_container_width=True):
                try:
                    inject_demo_data()
                    st.success("🎉 Đã nạp dữ liệu thử nghiệm trực tiếp vào SQLite!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Không thể nạp dữ liệu mẫu: {str(e)}")
        else:
            st.subheader("📋 Danh sách phân công hiện tại:")
            st.dataframe(df_tv, use_container_width=True)

st.sidebar.markdown("---")

# ==================================================================================
# --- KHỐI HIỂN THỊ TRẠNG THÁI TÀI KHOẢN (VỊ TRÍ 3 - DƯỚI CÙNG SIDEBAR) ---
# ==================================================================================
st.sidebar.markdown("### 🔑 TRẠNG THÁI TÀI KHOẢN")

if current_device_ip in ADMIN_IP_WHITELIST:
    st.sidebar.success("👑 Thiết bị: Chủ dự án (Admin)")
    st.sidebar.caption(f"Nhận diện tự động qua IP: `{current_device_ip}`. Đặc quyền chạy trực tiếp bằng Key hệ thống vĩnh viễn.")
    st.session_state["gv_api_key_input"] = ""
else:
    # Dự phòng nếu thầy ở vùng mạng khác thì hiện ô nhập mật mã/Key như cũ
    mat_ma_nhap = st.sidebar.text_input("Mật mã định danh Admin (Nếu có):", type="password", key="admin_password_input_field_v2")
    if mat_ma_nhap:
        st.session_state["current_entered_password"] = mat_ma_nhap

    if st.session_state.get("current_entered_password") == "123456":
        st.sidebar.success("👑 Thiết bị: Chủ dự án (Admin)")
        st.sidebar.caption("Kích hoạt đặc quyền qua mật mã thành công.")
        st.session_state["gv_api_key_input"] = ""
    else:
        st.sidebar.warning("🔒 Thiết bị: Thành viên/Giáo viên")
        st.sidebar.caption("Vui lòng dán API Key cá nhân từ Google AI Studio để mở khóa phân hệ.")
        st.sidebar.text_input("Nhập API Key Gemini của thầy/cô:", type="password", placeholder="AIzaSy...", key="gv_api_key_input")
        if st.session_state["gv_api_key_input"]:
            st.sidebar.success("🟢 Đã nhận diện Key cá nhân.")

# --- KHỐI THÔNG TIN TÁC GIẢ & ĐƠN VỊ ---
st.sidebar.markdown("<br>", unsafe_allow_html=True)
st.sidebar.markdown("""
    <div style='text-align: center; border-top: 1px solid #ddd; padding-top: 12px; margin-top: 5px;'>
        <p style='color: #991B1B; font-weight: bold; margin-bottom: 2px; font-size: 14px;'>Tác giả: Lê Hồng Dưỡng</p>
        <p style='color: #1E3A8A; font-weight: bold; margin-bottom: 0px; font-size: 14px;'>Đơn vị: THCS Nguyễn Chí Thanh</p>
    </div>
""", unsafe_allow_html=True)
