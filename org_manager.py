import streamlit as st
import pandas as pd

def render_org_section():
    st.subheader("📋 HỆ THỐNG QUẢN LÝ TỔ CHUYÊN MÔN")
    
    # Kiểm tra session_state để tránh lỗi key
    if "db_thanh_vien" not in st.session_state:
        st.session_state["db_thanh_vien"] = []
    if "db_phan_cong_hien_tai" not in st.session_state:
        st.session_state["db_phan_cong_hien_tai"] = []

    tab1, tab2, tab3 = st.tabs(["👥 Danh sách thành viên", "📊 Phân công giảng dạy", "🏆 Thành tích & Thi đua"])
    
    with tab1:
        st.write("Danh sách giáo viên trong tổ:")
        st.dataframe(pd.DataFrame(st.session_state["db_thanh_vien"]), use_container_width=True)
        
    with tab2:
        st.write("Phân công chuyên môn:")
        st.dataframe(pd.DataFrame(st.session_state["db_phan_cong_hien_tai"]), use_container_width=True)
        
    with tab3:
        st.write("Dữ liệu thành tích và thi đua:")
        st.json(st.session_state.get("db_thanh_tich_da_nam", {}))
