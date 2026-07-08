import streamlit as st
import pandas as pd

def render_timetable_section():
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU GIÁO VIÊN")
    
    # 1. Tải file TKB
    uploaded_tkb = st.file_uploader("📥 Nhập file TKB (.xlsx)", type=["xlsx"])
    
    if uploaded_tkb:
        # Đọc sheet TKB_GV_S (bỏ qua 3 dòng đầu không chứa dữ liệu)
        df = pd.read_excel(uploaded_tkb, sheet_name="TKB_GV_S", header=3)
        
        # Làm sạch cột: chỉ giữ lại Thứ, Tiết và tên GV
        # Các cột từ vị trí thứ 3 trở đi là tên giáo viên
        teacher_names = df.columns[2:].tolist()
        
        # 2. Chọn giáo viên
        selected_teacher = st.selectbox("👤 Chọn giáo viên:", teacher_names)
        
        if selected_teacher:
            # Lọc dữ liệu: Chỉ lấy Thứ, Tiết và cột của giáo viên đó
            tkb_gv = df[["THỨ", "TIẾT", selected_teacher]].copy()
            
            # Đổi tên cột cho đẹp
            tkb_gv.columns = ["Thứ", "Tiết", "Lịch dạy (Lớp - Môn)"]
            
            # Điền "Thứ" cho các tiết trống (vì trong excel chỉ có thứ ở dòng đầu)
            tkb_gv["Thứ"] = tkb_gv["Thứ"].ffill()
            
            # 3. Hiển thị
            st.subheader(f"Lịch dạy của: {selected_teacher}")
            st.dataframe(tkb_gv, use_container_width=True, hide_index=True)
            
            # Nút xuất riêng TKB giáo viên
            csv = tkb_gv.to_csv(index=False).encode('utf-8')
            st.download_button(f"📤 Tải TKB của {selected_teacher}", csv, f"TKB_{selected_teacher}.csv", "text/csv")
    else:
        st.info("💡 Vui lòng tải file Excel Thời khóa biểu lên.")
