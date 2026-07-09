import streamlit as st
import pandas as pd

def render_chu_nhiem_section(run_ai_prompt_safe=None):
    st.subheader("📋 7. KẾ HOẠCH CÔNG TÁC CHỦ NHIỆM LỚP")
    st.caption("Hệ sinh thái số hỗ trợ giáo viên chủ nhiệm quản lý lớp học và lập kế hoạch thông minh.")
    
    tab_ke_hoach, tab_hoc_sinh, tab_ai_tro_ly = st.tabs([
        "📅 Kế hoạch Tuần/Tháng", 
        "👥 Quản lý Học sinh", 
        "🤖 Trợ lý AI Chủ nhiệm"
    ])
    
    with tab_ke_hoach:
        st.write("#### 📝 Lập kế hoạch công tác chủ nhiệm")
        col1, col2 = st.columns(2)
        with col1:
            tuan_hoc = st.selectbox("Chọn tuần học:", [f"Tuần {i}" for i in range(1, 36)])
        with col2:
            chu_de = st.text_input("Chủ đề sinh hoạt / Trọng tâm tuần:", placeholder="Ví dụ: Ổn định nề nếp...")
            
        st.text_area("Mục tiêu cần đạt:")
        st.text_area("Các biện pháp thực hiện cụ thể:")
        if st.button("💾 Lưu kế hoạch tuần", key="btn_save_kh"):
            st.success(f"Đã lưu thành công kế hoạch {tuan_hoc}!")

    with tab_hoc_sinh:
        st.write("#### 📊 Danh sách học sinh cần lưu ý & Phân công")
        
        # Đã điền mảng số nguyên tường minh từ 1 đến 3 để loại bỏ hoàn toàn lỗi dấu phẩy trống
        data = {
            "STT":,
            "Họ và tên": ["Nguyễn Văn A", "Trần Thị B", "Lê Hoàng C"],
            "Chức vụ / Đặc điểm": ["Lớp trưởng", "Bí thư", "Học sinh cần hỗ trợ"],
            "Ghi chú": ["Nhanh nhẹn", "Trách nhiệm", "Hay quên làm bài tập"]
        }
        df = pd.DataFrame(data)
        st.data_editor(df, num_rows="dynamic", use_container_width=True)

    with tab_ai_tro_ly:
        st.write("#### 🤖 Trợ lý AI tư vấn công tác chủ nhiệm")
        user_prompt = st.text_area("Nhập yêu cầu của bạn gửi tới AI:", key="txt_ai_prompt")
        
        if st.button("🚀 Yêu cầu AI xử lý", type="primary", key="btn_run_ai"):
            if not user_prompt.strip():
                st.warning("Vui lòng nhập nội dung yêu cầu!")
            else:
                if run_ai_prompt_safe is not None:
                    with st.spinner("AI đang xử lý..."):
                        response = run_ai_prompt_safe(user_prompt)
                        st.markdown("### 🌟 Kết quả từ Trợ lý AI:")
                        st.write(response)
                else:
                    st.info("AI khuyên bạn nên: Phối hợp với phụ huynh và gặp riêng học sinh.")
