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
    key="app_main_sidebar_navigation_root_key_2026_v9" # <-- Sửa đổi thành key độc nhất vô nhị
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
        
        # --- KIỂM TRA VÀ XỬ LÝ DỮ LIỆU TỪ MỤC QUẢN LÝ THÀNH VIÊN ---
        df_tv = pd.DataFrame(st.session_state["db_thanh_vien"])
        
        # Nếu chưa có ai nhập dữ liệu thực tế ở mục 1, giao diện hiển thị tùy chọn nạp dữ liệu demo để đi thi
        if df_tv.empty:
            st.warning("ℹ️ Hiện tại chưa có dữ liệu giáo viên nào được nhập từ phân hệ '1. Quản lý & Phân công chuyên môn'.")
            
            # Thiết kế nút bấm kích hoạt dữ liệu giả lập thông minh phục vụ chấm thi nhanh
            if st.button("💡 Nạp nhanh dữ liệu mẫu để thử nghiệm biểu đồ", type="primary", use_container_width=True):
                st.session_state["db_thanh_vien"] = [
                    {"Họ và tên": "Thầy Lê Hồng Dưỡng", "Phân môn chính": "Khoa học tự nhiên (Phân môn Vật lí)", "Số tiết/Tuần": 14},
                    {"Họ và tên": "Cô Nguyễn Thị Mai", "Phân môn chính": "Khoa học tự nhiên (Phân môn Hóa học)", "Số tiết/Tuần": 16},
                    {"Hên và tên": "Thầy Trần Văn Tâm", "Phân môn chính": "Khoa học tự nhiên (Phân môn Sinh học)", "Số tiết/Tuần": 12},
                    {"Họ và tên": "Cô Lê Thị Thúy", "Phân môn chính": "Lịch sử và Địa lí (Phân môn Lịch sử)", "Số tiết/Tuần": 15},
                    {"Họ và tên": "Thầy Phạm Minh Hoàng", "Phân môn chính": "Lịch sử và Địa lí (Phân môn Địa lý)", "Số tiết/Tuần": 14}
                ]
                st.success("🎉 Đã nạp dữ liệu thử nghiệm! Hệ thống đang xử lý sơ đồ biểu đồ...")
                st.rerun()
                
        # --- LUỒNG XỬ LÝ ĐỒNG BỘ DỮ LIỆU THỰC TẾ ---
        else:
            # Chuẩn hóa tên cột để đảm bảo tính toán không bị lỗi KeyError nếu file bên kia gõ lệch chữ
            if "Phân môn chính" not in df_tv.columns and "Môn học / Phân môn phụ trách" in df_tv.columns:
                df_tv = df_tv.rename(columns={"Môn học / Phân môn phụ trách": "Phân môn chính"})
            if "Số tiết/Tuần" not in df_tv.columns and "Số tiết" in df_tv.columns:
                df_tv = df_tv.rename(columns={"Số tiết": "Số tiết/Tuần"})
                
            # 1. Hộp số liệu tổng quan tự động tính toán (Metrics)
            st.markdown("### 📌 Chỉ số tổng quan tổ chuyên môn")
            m_col1, m_col2, m_col3 = st.columns(3)
            
            tong_so_gv = len(df_tv)
            m_col1.metric(label="👥 Tổng số Giáo viên trong tổ", value=f"{tong_so_gv} Thầy/Cô")
            
            # Hàm nunique() tự động đếm xem có bao nhiêu nhóm phân môn khác nhau được nhập ở mục 1
            so_phan_mon = df_tv["Phân môn chính"].nunique() if "Phân môn chính" in df_tv.columns else 0
            m_col2.metric(label="📚 Số lượng Môn học/Phân môn", value=f"{so_phan_mon} Nhóm")
            
            # Hàm sum() tự động tính toán lũy tiến tổng số định mức tiết dựa theo dữ liệu thực tế
            if "Số tiết/Tuần" in df_tv.columns:
                # Ép kiểu số để tránh lỗi nếu người dùng nhập chữ
                df_tv["Số tiết/Tuần"] = pd.to_numeric(df_tv["Số tiết/Tuần"], errors='coerce').fillna(0)
                tong_tiet = int(df_tv["Số tiết/Tuần"].sum())
                m_col3.metric(label="⏱️ Tổng số tiết định mức / tuần", value=f"{tong_tiet} Tiết")
            else:
                m_col3.metric(label="⏱️ Tổng số tiết định mức / tuần", value="0 Tiết")
                
            st.markdown("---")
            
            # 2. Khu vực xây dựng các biểu đồ so sánh trực quan độc lập
            chart_col1, chart_col2 = st.columns(2)
            
            # Biểu đồ 1: Số lượng giáo viên phân bổ theo từng Phân môn chính
            with chart_col1:
                st.markdown("##### 📈 Số lượng giáo viên theo Phân môn")
                if "Phân môn chính" in df_tv.columns:
                    counts = df_tv["Phân môn chính"].value_counts()
                    st.bar_chart(counts, color="#1E3A8A")
                else:
                    st.caption("Thiếu dữ liệu cột 'Phân môn chính' để lập sơ đồ.")
                
            # Biểu đồ 2: Định mức phân bổ số tiết dạy của từng giáo viên trong tuần
            with chart_col2:
                st.markdown("##### ⏳ Định mức Tiết dạy/Tuần của từng Giáo viên")
                # Đồng bộ tên cột họ tên nếu file bên phân công đặt tên là 'Họ tên' hoặc 'Giáo viên thực hiện'
                col_name_check = "Họ và tên" if "Họ và tên" in df_tv.columns else (df_tv.columns[0] if len(df_tv.columns) > 0 else "")
                
                if "Số tiết/Tuần" in df_tv.columns and col_name_check:
                    df_chart_tiet = df_tv.set_index(col_name_check)[["Số tiết/Tuần"]]
                    st.bar_chart(df_chart_tiet, color="#EF4444")
                else:
                    st.caption("Chưa có trường số liệu định mức số tiết dạy cụ thể của giáo viên.")
                    
            # 3. Bảng danh sách chi tiết lưu trữ dưới nền hệ thống
            st.markdown("---")
            st.markdown("### 🗂️ Danh sách trích xuất dữ liệu chi tiết")
            st.dataframe(df_tv, use_container_width=True, hide_index=True)

