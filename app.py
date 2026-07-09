import streamlit as st
import pandas as pd
from google import genai

# --- 1. PHÂN LUỒNG IMPORT CÁC MODULE ĐỘC LẬP (ĐÃ KHỬ TRÙNG LẶP) ---
from exam_designer import render_exam_designer_section
from grade_manager import render_grade_manager_section
from tkb_manager import render_tkb_manager  
from khbd_section_manager import render_khbd_section  # Cập nhật theo tên file thực tế của bạn
from danh_gia_manager import render_assessment_section

from org_manager import render_org_section
from bien_ban_manager import render_meeting_minutes
from ke_hoach_ca_nhan_manager import render_personal_plan 

# --- 2. CẤU HÌNH ĐỌC API KEY TỰ ĐỘNG TỪ TRONG SECRETS ---
API_KEY_HE_THONG = st.secrets.get("GEMINI_API_KEY", "")

def run_ai_prompt_safe(prompt_text):
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
# --- THANH ĐIỀU HƯỚNG TỔNG (ĐÃ VÁ LỖI TRÙNG KEY ĐỘC NHẤT) ---
## ==================================================================================
st.sidebar.markdown("### MENU HỆ THỐNG")
st.sidebar.caption("CHỌN PHÂN HỆ TÁC NGHIỆP")

phan_he = st.sidebar.radio(
    "Chọn phân hệ:",
    ["Trợ lý Giảng dạy (Giáo viên)", "Trợ lý Quản lý (Tổ chuyên môn)"],
    label_visibility="collapsed",
    key="app_main_sidebar_navigation_root_key_2026_v9"
)

# --- 5. XỬ LÝ ĐIỀU HƯỚNG ---
if phan_he == "Trợ lý Giảng dạy (Giáo viên)":
    st.sidebar.markdown("### 🛠️ CHỨC NĂNG GIÁO VIÊN")
    menu = st.sidebar.selectbox("Nội dung giảng dạy", ["1. Thiết kế KHBD", "2. Thiết kế Đề KT", "3. Đánh giá HS", "4. Quản lý điểm (SMAS)", "5. Quản lý TKB"], label_visibility="collapsed", key="menu_gv_selectbox_v9")
    
    st.sidebar.success("🔑 Đã kết nối API Key hệ thống từ Manage App.")
    
    if menu == "1. Thiết kế KHBD": 
        render_khbd_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "2. Thiết kế Đề KT": 
        render_exam_designer_section("", lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Đánh giá HS": 
        render_assessment_section(lambda p: run_ai_prompt_safe(p))
    elif menu == "4. Quản lý điểm (SMAS)": 
        render_grade_manager_section()
    elif menu == "5. Quản lý TKB": 
        render_tkb_manager()

else:  # Phân hệ Quản lý tổ chuyên môn
    st.sidebar.markdown("### 📂 QUẢN LÝ TỔ CHUYÊN MÔN")
    menu = st.sidebar.selectbox("Nội dung quản lý", ["1. Quản lý & Phân công chuyên môn", "2. Biên bản sinh hoạt", "3. Kế hoạch cá nhân", "4. Thống kê số liệu"], label_visibility="collapsed", key="menu_ql_selectbox_v9")
    
    if menu == "1. Quản lý & Phân công chuyên môn": 
        render_org_section()
    elif menu == "2. Biên bản sinh hoạt": 
        render_meeting_minutes(lambda p: run_ai_prompt_safe(p))
    elif menu == "3. Kế hoạch cá nhân": 
        render_personal_plan(lambda p: run_ai_prompt_safe(p))
    elif menu == "4. Thống kê số liệu": 
        st.header("📊 THỐNG KÊ SỐ LIỆU TỔ CHUYÊN MÔN")
        
        # Đọc dữ liệu từ bộ lưu trữ của mục 1
        raw_data = st.session_state.get("db_thanh_vien", [])
        df_tv = pd.DataFrame(raw_data)
        
        # Bộ quét chuẩn hóa tên cột thông minh (khử lệch Key giữa 2 file)
        if not df_tv.empty:
            for col in ["Họ tên", "Họ tên giáo viên", "Giáo viên thực hiện", "Họ và Tên", "Tên"]:
                if col in df_tv.columns and "Họ và tên" not in df_tv.columns:
                    df_tv = df_tv.rename(columns={col: "Họ và tên"})
            for col in ["Môn học / Phân môn phụ trách", "Môn giảng dạy", "Phân môn", "Môn", "Môn phụ trách"]:
                if col in df_tv.columns and "Phân môn chính" not in df_tv.columns:
                    df_tv = df_tv.rename(columns={col: "Phân môn chính"})
            for col in ["Số tiết", "Số tiết dạy", "Định mức tiết", "Số tiết / Tuần", "Số tiết/tuần"]:
                if col in df_tv.columns and "Số tiết/Tuần" not in df_tv.columns:
                    df_tv = df_tv.rename(columns={col: "Số tiết/Tuần"})
        
        # Luồng kiểm tra nếu rỗng thực sự thì hiện tùy chọn nạp demo thử nghiệm nhanh
        if df_tv.empty:
            st.warning("ℹ️ Hiện tại chưa có dữ liệu giáo viên nào được nhập từ phân hệ '1. Quản lý & Phân công chuyên môn'.")
            
            if st.button("💡 Nạp nhanh dữ liệu mẫu để thử nghiệm biểu đồ", type="primary", use_container_width=True):
                st.session_state["db_thanh_vien"] = [
                    {"Họ và tên": "Thầy Lê Hồng Dưỡng", "Phân môn chính": "Khoa học tự nhiên (Phân môn Vật lí)", "Số tiết/Tuần": 14},
                    {"Họ và tên": "Cô Nguyễn Thị Mai", "Phân môn chính": "Khoa học tự nhiên (Phân môn Hóa học)", "Số tiết/Tuần": 16},
                    {"Họ và tên": "Thầy Trần Văn Tâm", "Phân môn chính": "Khoa học tự nhiên (Phân môn Sinh học)", "Số tiết/Tuần": 12},
                    {"Họ và tên": "Cô Lê Thị Thúy", "Phân môn chính": "Lịch sử và Địa lí (Phân môn Lịch sử)", "Số tiết/Tuần": 15},
                    {"Họ và tên": "Thầy Phạm Minh Hoàng", "Phân môn chính": "Lịch sử và Địa lí (Phân môn Địa lý)", "Số tiết/Tuần": 14}
                ]
                st.success("🎉 Đã nạp dữ liệu thử nghiệm! Đang xây dựng sơ đồ...")
                st.rerun()
        else:
            # Xử lý các ô trống (NaN) nếu giáo viên quên nhập liệu ở mục 1
            if "Phân môn chính" not in df_tv.columns: df_tv["Phân môn chính"] = "Chưa phân môn"
            if "Họ và tên" not in df_tv.columns: df_tv["Họ và tên"] = "Giáo viên ẩn danh"
            df_tv["Phân môn chính"] = df_tv["Phân môn chính"].fillna("Chưa phân môn")
            df_tv["Họ và tên"] = df_tv["Họ và tên"].fillna("Giáo viên ẩn danh")
            
            # Chỉ số tổng quan (Metrics)
            st.markdown("### 📌 Chỉ số tổng quan tổ chuyên môn")
            m_col1, m_col2, m_col3 = st.columns(3)
            
            tong_so_gv = len(df_tv)
            m_col1.metric(label="👥 Tổng số Giáo viên trong tổ", value=f"{tong_so_gv} Thầy/Cô")
            
            so_phan_mon = df_tv["Phân môn chính"].nunique()
            m_col2.metric(label="📚 Số lượng Môn học/Phân môn", value=f"{so_phan_mon} Nhóm")
            
            if "Số tiết/Tuần" in df_tv.columns:
                df_tv["Số tiết/Tuần"] = pd.to_numeric(df_tv["Số tiết/Tuần"], errors='coerce').fillna(0)
                tong_tiet = int(df_tv["Số tiết/Tuần"].sum())
                m_col3.metric(label="⏱️ Tổng số tiết định mức / tuần", value=f"{tong_tiet} Tiết")
            else:
                m_col3.metric(label="⏱️ Tổng số tiết định mức / tuần", value="0 Tiết")
                
            st.markdown("---")
            
            # Khởi tạo các biểu đồ song song
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown("##### 📈 Số lượng giáo viên theo Phân môn")
                counts = df_tv["Phân môn chính"].value_counts()
                st.bar_chart(counts, color="#1E3A8A")
                
            with chart_col2:
                st.markdown("##### ⏳ Định mức Tiết dạy/Tuần của từng Giáo viên")
                df_chart_tiet = df_tv.set_index("Họ và tên")[["Số tiết/Tuần"]]
                st.bar_chart(df_chart_tiet, color="#EF4444")
                    
            st.markdown("---")
            st.markdown("### 🗂️ Danh sách trích xuất dữ liệu chi tiết")
            st.dataframe(df_tv, use_container_width=True, hide_index=True)
