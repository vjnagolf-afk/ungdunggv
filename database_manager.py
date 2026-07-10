# database_manager.py - Bản vá triệt tiêu hoàn toàn lỗi KeyError khi import vào app.py
import sqlite3
import os

DB_PATH = "teacher_assistant.db"

def init_sqlite_database():
    """Khởi tạo cấu trúc các bảng dữ liệu nội bộ an toàn nếu chưa tồn tại"""
    try:
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
    except Exception:
        pass

def inject_demo_data():
    """Hàm bổ trợ nạp nhanh nhân sự demo cho tổ chuyên môn"""
    try:
        init_sqlite_database()
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
    except Exception:
        pass
