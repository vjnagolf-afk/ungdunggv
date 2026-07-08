import streamlit as st
import pandas as pd

def render_timetable_section():
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU")
    
    uploaded_tkb = st.file_uploader("📥 Nhập file TKB (.xlsx)", type=["xlsx"])
    
    if uploaded_tkb:
        # Đọc dữ liệu
        df = pd.read_excel(uploaded_tkb, sheet_name="TKB_GV_S", header=3)
        
        # Lọc danh sách giáo viên chuẩn (loại bỏ cột Unnamed)
        teacher_names = [c for c in df.columns if "Unnamed" not in str(c) and c not in ["THỨ", "TIẾT"]]
        
        # Chia giao diện thành 2 tab
        tab1, tab2 = st.tabs(["📊 Thời khóa biểu chung", "👤 TKB theo giáo viên"])
        
        with tab1:
            st.subheader("Toàn bộ Thời khóa biểu")
            st.dataframe(df, use_container_width=True)
            
        with tab2:
            st.subheader("Tra cứu lịch dạy cá nhân")
            # Sử dụng selectbox với danh sách đã lọc, KHÔNG THỂ CHỌN CÁI KHÁC
            selected_teacher = st.selectbox("👤 Chọn giáo viên từ danh sách:", teacher_names)
            
            if selected_teacher:
                # Trích xuất dữ liệu của giáo viên đó
                tkb_gv = df[["THỨ", "TIẾT", selected_teacher]].copy()
                tkb_gv.columns = ["Thứ", "Tiết", "Lịch dạy (Lớp - Môn)"]
                
                # Điền thứ cho dễ nhìn
                tkb_gv["Thứ"] = tkb_gv["Thứ"].ffill()
                
                # Loại bỏ các dòng trống (tiết không dạy)
                tkb_gv = tkb_gv.dropna(subset=["Lịch dạy (Lớp - Môn)"])
                
                st.dataframe(tkb_gv, use_container_width=True, hide_index=True)
                
                # Nút xuất TKB cá nhân
                csv = tkb_gv.to_csv(index=False).encode('utf-8')
                st.download_button(f"📤 Tải TKB của {selected_teacher}", csv, f"TKB_{selected_teacher}.csv", "text/csv")
    else:
        st.info("💡 Vui lòng tải file Excel Thời khóa biểu lên để bắt đầu.")
