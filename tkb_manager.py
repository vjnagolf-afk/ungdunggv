import streamlit as st
import pandas as pd

def render_tkb_manager():
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU")
    uploaded_tkb = st.file_uploader("Tải lên file TKB (.xlsx)", type=["xlsx"])
    if uploaded_tkb:
        try:
            df = pd.read_excel(uploaded_tkb, sheet_name="TKB_GV_S", header=3)
            # Lọc tên cột: bỏ qua các cột 'Unnamed', THỨ, TIẾT
            teachers = [c for c in df.columns if isinstance(c, str) and "Unnamed" not in c and c not in ["THỨ", "TIẾT"]]
            
            if not teachers:
                st.error("Không tìm thấy giáo viên! Kiểm tra file Excel xem tên GV có nằm ở dòng header không.")
            else:
                selected_teacher = st.selectbox("👤 Chọn giáo viên:", sorted(teachers))
                tkb_gv = df[["THỨ", "TIẾT", selected_teacher]].copy()
                tkb_gv.columns = ["Thứ", "Tiết", "Lịch dạy"]
                tkb_gv["Thứ"] = tkb_gv["Thứ"].ffill()
                st.dataframe(tkb_gv.dropna(subset=["Lịch dạy"]), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")
