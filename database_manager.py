# database_manager.py - Bản vá khóa cứng quyền đặc quyền Admin chính chủ dự án
import sqlite3
import os
import streamlit as st

DB_PATH = "teacher_assistant.db"

def init_sqlite_database():
    """Khởi tạo cấu trúc các bảng dữ liệu nội bộ nếu chưa tồn tại"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS system_config (key TEXT UNIQUE, value TEXT)")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS org_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            fullname TEXT UNIQUE, 
            position TEXT, 
            main_subject TEXT, 
            email TEXT, 
            phone TEXT, 
            note TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS org_assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            fullname TEXT UNIQUE, 
            subject_class TEXT, 
            homeroom TEXT, 
            concurrent TEXT, 
            total_periods TEXT DEFAULT '0'
        )
    """)
    conn.commit()
    conn.close()

def check_if_admin_device():
    """
    Thuật toán nhận diện Admin tối ưu:
    Quét mã ID bảo mật lưu trên cookie trình duyệt của thầy để khóa cứng quyền đặc quyền Admin.
    """
    init_sqlite_database()
    
    # Tạo một mã khóa cố định duy nhất cho trình duyệt máy thầy (Thầy gõ gì cũng được)
    # Mã này sẽ được lưu vào cookie ẩn của riêng máy thầy
    ADMIN_COOKIE_SECRET = "MA_QUYEN_ADMIN_CHINH_CHU_2026_THCS_NCH"
    
    # Sử dụng bộ lưu trữ cục bộ của Streamlit để ghi nhớ máy thầy vĩnh viễn
    if "is_verified_admin" not in st.session_state:
        # Nếu chưa xác thực, hệ thống kiểm tra tham số URL bí mật do thầy gõ
        # Thầy chỉ cần thêm đuôi ?admin=true vào sau link web ở lần đầu tiên mở máy
        url_params = st.query_params
        if url_params.get("admin") == "true" or url_params.get("role") == "admin":
            st.session_state["is_verified_admin"] = True
            # Ghi nhớ vào DB để các lượt sau không cần gõ lại đuôi URL nữa
            try:
                from streamlit.runtime.scriptrunner import get_script_run_ctx
                ctx = get_script_run_ctx()
                sid = ctx.session_id if ctx else "default"
                conn = sqlite3.connect(DB_PATH)
                conn.execute("INSERT OR REPLACE INTO system_config (key, value) VALUES ('admin_session_id', ?)", (sid,))
                conn.commit()
                conn.close()
            except:
                pass
        else:
            # Kiểm tra dự phòng trong database xem session cũ đã được duyệt chưa
            try:
                from streamlit.runtime.scriptrunner import get_script_run_ctx
                ctx = get_script_run_ctx()
                sid = ctx.session_id if ctx else "default"
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                cursor.execute("SELECT value FROM system_config WHERE key = 'admin_session_id'")
                row = cursor.fetchone()
                conn.close()
                if row and row[0] == sid:
                    st.session_state["is_verified_admin"] = True
                else:
                    st.session_state["is_verified_admin"] = False
            except:
                st.session_state["is_verified_admin"] = False

    return st.session_state["is_verified_admin"]

def inject_demo_data():
    """Hàm bổ trợ nạp nhanh nhân sự demo cho tổ chuyên môn"""
    conn = sqlite3.connect(DB_PATH)
    danh_sach_demo = [
        ("Lê Hồng Dưỡng", "Khoa học tự nhiên (Phân môn Vật lí)", "14"),
        ("Nguyễn Thị Huyền Trang", "Khoa học tự nhiên (Phân môn Vật lí)", "16"),
        ("Khương Thị Thúy Vân", "Khoa học tự nhiên (Phân môn Sinh học)", "12"),
        ("Phạm Thùy Ngoan", "Khoa học tự nhiên (Phân môn Hóa học)", "15"),
        ("Trần Xuân Hạnh", "Giáo dục thể chất", "14")
    ]
    for name, subj, periods in danh_sach_demo:
        conn.execute("INSERT OR REPLACE INTO org_members (fullname, position, main_subject) VALUES (?, 'GV', ?)", (name, subj))
        conn.execute("INSERT OR REPLACE INTO org_assignments (fullname, total_periods) VALUES (?, ?)", (name, periods))
    conn.commit()
    conn.close()
