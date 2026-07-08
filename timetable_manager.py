# timetable_manager.py
import streamlit as st
import sqlite3
import pandas as pd

DB_PATH = "teacher_assistant.db"

def setup_timetable_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS timetable (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_name TEXT,
            day_of_week TEXT,
            period INTEGER,
            class_name TEXT,
            subject TEXT
        )
    """)
    conn.commit()
    conn.close()

def render_timetable_section():
    setup_timetable_db()
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU")
    
    # Upload TKB
    uploaded_tkb = st.file_uploader("📥 Nhập file TKB (.xlsx)", type=["xlsx"])
    if uploaded_tkb:
        if st.button("🚀 Đồng bộ TKB"):
            df = pd.read_excel(uploaded_tkb)
            # Giả định file Excel có các cột: Teacher, Day, Period, Class, Subject
            conn = sqlite3.connect(DB_PATH)
            df.to_sql("timetable", conn, if_exists="replace", index=False)
            conn.close()
            st.success("✅ Đã cập nhật TKB thành công!")
            st.rerun()

    # Hiển thị TKB
    conn = sqlite3.connect(DB_PATH)
    df_tkb = pd.read_sql_query("SELECT * FROM timetable", conn)
    conn.close()
    
    if not df_tkb.empty:
        st.dataframe(df_tkb, use_container_width=True)
    else:
        st.info("💡 Chưa có dữ liệu TKB. Vui lòng nhập file.")