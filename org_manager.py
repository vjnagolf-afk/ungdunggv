import streamlit as st
import pandas as pd

def render_org_section():
    # --- 1. XỬ LÝ XÁC THỰC QUYỀN ADMIN TRÊN SIDEBAR ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
    st.sidebar.caption("VỚI TƯ CÁCH LÀ:")
    
    vai_tro = st.sidebar.selectbox(
        "Vai trò",
        ["Giáo viên bộ môn", "Tổ trưởng chuyên môn (Admin)", "Ban giám hiệu"],
        label_visibility="collapsed"
    )
    
    is_admin = False
    
    if vai_tro == "Tổ trưởng chuyên môn (Admin)":
        st.sidebar.caption("Nhập mã pin quản lý Admin:")
        ma_pin = st.sidebar.text_input("Mã PIN", type="password", label_visibility="collapsed")
        
        # Kiểm tra mã PIN (Thay '123456' bằng mã PIN bạn muốn đặt)
        if ma_pin == "123456":
            st.sidebar.success("✅ Xác thực thành công! Quyền Admin đã mở.")
            is_admin = True
        elif ma_pin != "":
            st.sidebar.error("❌ Mã PIN không chính xác.")

    # --- 2. HIỂN THỊ NỘI DUNG CHÍNH THEO TABS ---
    st.subheader("📋 HỆ THỐNG QUẢN LÝ VÀ PHÂN CÔNG CHUYÊN MÔN GIẢNG DẠY")
    tab1, tab2, tab3 = st.tabs(["👥 Danh sách thành viên", "📊 Phân công", "🏅 Thành tích"])
    
    with tab1:
        st.markdown("#### 👥 Danh sách thành viên tổ")
        if is_admin:
            st.info("💡 Chế độ Quản trị: Bạn có thể nhập dữ liệu thành viên.")
            # Form nhập liệu giả lập cho Admin
            with st.form("form_them_thành_vien"):
                ten = st.text_input("Họ và tên giáo viên:")
                mon = st.selectbox("Phân môn chính:", ["Toán", "Văn", "Anh", "KHTN", "Lịch sử & Địa lý"])
                if st.form_submit_button("Thêm thành viên"):
                    if ten:
                        st.session_state["db_thanh_vien"].append({"Họ và tên": ten, "Phân môn chính": mon})
                        st.success(f"Đã thêm giáo viên {ten}")
                        st.rerun()
        
        # Hiển thị bảng dữ liệu hiện có
        df_tv = pd.DataFrame(st.session_state.get("db_thanh_vien", []))
        if not df_tv.empty:
            st.dataframe(df_tv, use_container_width=True)
        else:
            st.caption("Chưa có dữ liệu thành viên tổ.")
            
    with tab2:
        st.markdown("#### 📊 Phân công giảng dạy")
        df_pc = pd.DataFrame(st.session_state.get("db_phan_cong_hien_tai", []))
        if not df_pc.empty:
            st.dataframe(df_pc, use_container_width=True)
        else:
            st.caption("Chưa có dữ liệu phân công giảng dạy.")
            
    with tab3:
        st.markdown("#### 🏅 Thành tích & Thi đua")
        st.json(st.session_state.get("db_thanh_tich_da_nam", {}))

# --- CÁC HÀM QUẢN LÝ KHÁC (Để app.py gọi tới) ---
def render_meeting_minutes():
    st.header("📝 BIÊN BẢN SINH HOẠT TỔ")
    st.write("Giao diện nhập biên bản họp chuyên môn định kỳ.")
    # Thêm code nhập liệu vào đây

def render_personal_plan():
    st.header("📅 KẾ HOẠCH GIÁO DỤC CÁ NHÂN")
    st.write("Giao diện lập kế hoạch cá nhân (Phụ lục III - 5512).")
    # Thêm code lập kế hoạch vào đây
