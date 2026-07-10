# app.py - Tích hợp bộ nhớ LocalStorage vĩnh viễn, chống bắt nhập lại mật mã khi F5
import streamlit as st
import pandas as pd
import sqlite3
import os
import sys
import streamlit.components.v1 as components

# THUẬT TOÁN ĐƯỜNG DẪN: Ép hệ thống tìm module trong cùng thư mục chạy ứng dụng
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from database_manager import inject_demo_data, DB_PATH
from ai_service import run_ai_prompt_safe

# Nhúng các phân hệ tác nghiệp vệ tinh của thầy/cô
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

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

# 🚀 BỘ VÁ CSS: Khử padding/margin thừa giữa các phần tử trong Sidebar
st.markdown("""
    <style>
        [data-testid="stSidebarUserContent"] {
            padding-top: 1.5rem !important;
            padding-bottom: 1.5rem !important;
        }
        [data-testid="stSidebarUserContent"] .stMarkdown, 
        [data-testid="stSidebarUserContent"] .stRadio, 
        [data-testid="stSidebarUserContent"] .stSelectbox {
            margin-bottom: -0.4rem !important;
        }
        [data-testid="stSidebarUserContent"] hr {
            margin-top: 0.6rem !important;
            margin-bottom: 0.6rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# Tiêu đề giao diện chính
st.markdown("<h1 style='text-align: center; color: darkred; font-weight: bold;'>🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #0056b3; font-weight: bold; font-size: 16px;'>Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk</p>", unsafe_allow_html=True)
st.markdown("---")

# --- MENU ĐIỀU HƯỚNG TỔNG TẠI SIDEBAR (VỊ TRÍ 1 - TRÊN CÙNG) ---
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_root_key_2026_v9"
)

st.sidebar.markdown("---")

# ==================================================================================
# --- 🚀 THUẬT TOÁN ĐỌC/GHI BỘ NHỚ TRÌNH DUYỆT (LOCAL STORAGE) QUA JAVASCRIPT ---
# ==================================================================================
MAT_MA_ADMIN_CỐ_ĐỊNH = "123456"

# Khởi tạo biến lưu trữ mật mã trong bộ nhớ Streamlit nếu chưa có
if "saved_admin_password" not in st.session_state:
    st.session_state["saved_admin_password"] = ""

# Đọc mật mã ẩn từ URL được chuyển hướng từ JavaScript
query_params = st.query_params
if "get_pwd" in query_params:
    st.session_state["saved_admin_password"] = query_params["get_pwd"]

# Nhúng đoạn mã HTML/JS chạy ngầm để đọc và tự động điền mật mã khi F5
components.html(f"""
    <script>
        // 1. Kiểm tra xem máy thầy đã từng lưu mật mã admin chưa
        var stored_pwd = localStorage.getItem("nch_admin_password_storage");
        var current_url = new URL(window.parent.location.href);
        
        // 2. Nếu tìm thấy mật mã trong máy và URL chưa được nạp, tự chuyển hướng nạp ngầm vào Streamlit
        if (stored_pwd && !current_url.searchParams.has("get_pwd")) {{
            current_url.searchParams.set("get_pwd", stored_pwd);
            window.parent.location.href = current_url.href;
        }}
    </script>
""", height=0, width=0)

st.sidebar.markdown("### 🔑 TRẠNG THÁI TÀI KHOẢN")

# Lấy mật mã đã lưu làm giá trị mặc định cho ô nhập liệu
default_pwd_value = st.session_state["saved_admin_password"]

mat_ma_nhap = st.sidebar.text_input(
    "Mật mã định danh Admin (Nếu có):", 
    value=default_pwd_value, 
    type="password", 
    key="admin_password_permanent_key"
)

# Nếu thầy gõ mật mã mới, lập tức ra lệnh cho JavaScript lưu chặt vào ổ cứng trình duyệt vĩnh viễn
if mat_ma_nhap:
    components.html(f"""
        <script>
            localStorage.setItem("nch_admin_password_storage", "{mat_ma_nhap}");
        </script>
    """, height=0, width=0)

# Kiểm tra điều kiện gán đặc quyền Admin dựa trên mật mã nạp tự động
if mat_ma_nhap == MAT_MA_ADMIN_CỐ_ĐỊNH:
    is_admin_owner = True
    st.sidebar.success("👑 Thiết bị: Chủ dự án (Admin)")
    st.sidebar.caption("Tự động kích hoạt quyền đặc quyền chạy trực tiếp bằng Key hệ thống.")
    st.session_state["gv_api_key_input"] = ""
else:
    is_admin_owner = False
    st.sidebar.warning("🔒 Thiết bị: Thành viên/Giáo viên")
    st.sidebar.caption("Vui lòng dán API Key cá nhân từ Google AI Studio để mở khóa phân hệ.")
    st.sidebar.text_input("Nhập API Key Gemini của thầy/cô:", type="password", placeholder="AIzaSy...", key="gv_api_key_input")
    if st.session_state["gv_api_key_input"]:
        st.sidebar.success("🟢 Đã nhận diện Key cá nhân.")

st.sidebar.markdown("---")

# --- KHỐI ĐIỀU HƯỚNG TÁC NGHIỆP CHI TIẾT (VỊ TRÍ 2 - Ở GIỮA) ---
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB","6. Thiết kế bài dạy STEM","7. Kế hoạch công tác chủ nhiệm lớp"], label_visibility="collapsed", key="menu_gv_selectbox_v9")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "2. Thiết kế Đề KT": 
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

else:  # Phân hệ Quản lý tổ chuyên môn
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox("Nội dung quản lý", ["1. Quản lý & Phân công chuyên môn", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"], label_visibility="collapsed", key="menu_ql_selectbox_v9")
    
    if menu == "1. Quản lý & Phân công chuyên môn": 
        render_org_section()
    elif menu == "2. Biên bản sinh hoạt": 
        render_meeting_minutes(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "3. Kế hoạch cá nhân": 
        render_personal_plan(lambda p: run_ai_prompt_safe(p, is_admin_owner=is_admin_owner))
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        
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

# --- KHỐI THÔNG TIN TÁC GIẢ & ĐƠN VỊ (XUYÊN SUỐT DƯỚI CHÂN TRANG SIDEBAR) ---
st.sidebar.markdown("<br>", unsafe_allow_html=True)
st.sidebar.markdown("""
    <div style='text-align: center; border-top: 1px solid #ddd; padding-top: 12px; margin-top: 5px;'>
        <p style='color: #991B1B; font-weight: bold; margin-bottom: 2px; font-size: 14px;'>Tác giả: Lê Hồng Dưỡng</p>
        <p style='color: #1E3A8A; font-weight: bold; margin-bottom: 0px; font-size: 14px;'>Đơn vị: THCS Nguyễn Chí Thanh</p>
    </div>
""", unsafe_allow_html=True)
