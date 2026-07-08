import streamlit as st
import pandas as pd
from google import genai

# --- 1. PHÂN LUỒNG IMPORT CÁC MODULE ĐỘC LẬP (ĐÃ KHỬ TRÙNG LẶP) ---
from exam_designer import render_exam_designer_section
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_manager import render_khbd_section
from danh_gia_manager import render_assessment_section

from org_manager import render_org_section
from bien_ban_manager import render_meeting_minutes
from ke_hoach_ca_nhan_manager import render_personal_plan 

# --- 2. CẤU HÌNH ĐỌC API KEY TỰ ĐỘNG TỪ TRONG SECRETS ---
# Thử đọc key hệ thống, nếu chưa có thì gán chuỗi rỗng
API_KEY_HE_THONG = st.secrets.get("GEMINI_API_KEY", "")

def run_ai_prompt_safe(prompt_text):
    # Ưu tiên lấy key từ mục Secrets của Manage App trước
    api_key = API_KEY_HE_THONG
    
    if not api_key:
        return "⚠️ Hệ thống chưa được cấu hình API Key trong mục Secrets. Vui lòng liên hệ Admin.", "error"
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt_text,
        )
        return response.text, "gemini-2.5-flash"
    except Exception as main_error:
        error_msg = str(main_error)
        if "503" in error_msg or "UNAVAILABLE" in error_msg or "high demand" in error_msg:
            try:
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model='gemini-1.5-flash',
                    contents=prompt_text,
                )
                return response.text, "gemini-1.5-flash (Kênh Dự phòng)"
            except Exception as backup_error:
                return f"Lỗi quá tải hệ thống trên diện rộng: {str(backup_error)}", "error"
        else:
            return f"Lỗi kết nối AI: {error_msg}", "error"

# --- 3. KHỞI TẠO BỘ NHỚ TẠM ---
if "db_thanh_vien" not in st.session_state: st.session_state["db_thanh_vien"] = []
if "db_phan_cong_hien_tai" not in st.session_state: st.session_state["db_phan_cong_hien_tai"] = []

st.set_page_config(page_title="HỆ SINH THÁI SỐ GIÁO VIÊN", layout="wide")

st.title("🔰 HỆ SINH THÁI SỐ - HỖ TRỢ GIÁO VIÊN")
st.caption("Sản phẩm tham gia Cuộc thi AI for Life năm 2026, trường THCS Nguyễn Chí Thanh - Phường Tân Lập tỉnh Đắk Lắk")
st.markdown("---")

## ==================================================================================
# --- THANH ĐIỀU HƯỚNG SỬA LỖI TRÙNG ID (DÀNH CHO FILE APP.PY) ---
# ==================================================================================
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

# 💥 VÁ LỖI TẬN GỐC: Ép khóa key tĩnh độc lập để ngăn chặn lỗi DuplicateWidgetID khi gọi AI ngầm
phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_key_v1" # <-- Khóa định danh độc nhất
)

# --- 5. XỬ LÝ ĐIỀU HƯỚNG ---
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB"], label_visibility="collapsed")
    
    # 💡 ĐÃ LOẠI BỎ Ô NHẬP API KEY THỦ CÔNG (Vì hệ thống đã chạy ngầm tự động bảo mật)
    st.sidebar.success("🔑 Đã kết nối API Key hệ thống từ Manage App.")
    
    if menu == "1. Thiết kế KHBD": 
        # Gọi hàm và truyền hàm AI ngắn gọn hơn (Không cần truyền biến api_key thủ công)
        render_khbd_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "2. Thiết kế Đề KT": 
        render_exam_designer_section("", lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Đánh giá HS": 
        # Gọi phân hệ thiết kế Rubric từ file danh_gia_manager.py
        render_assessment_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()

else:  # Phân hệ Quản lý tổ chuyên môn
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox("Nội dung quản lý", ["1. Quản lý & Phân công chuyên môn", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"], label_visibility="collapsed")
    
    if menu == "1. Quản lý & Phân công chuyên môn": render_org_section()
elif menu == "2. Biên bản sinh hoạt": 
    render_meeting_minutes(lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Kế hoạch cá nhân": render_personal_plan()
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        df_tv = pd.DataFrame(st.session_state["db_thanh_vien"])
        if not df_tv.empty and "Phân môn chính" in df_tv.columns:
            st.bar_chart(df_tv["Phân môn chính"].value_counts())
        else:
            st.warning("⚠️ Chưa có dữ liệu thành viên hoặc thiếu cột 'Phân môn chính' để lập biểu đồ.")
