import streamlit as st
import pandas as pd

def render_tab_7(run_ai_prompt_safe=None):
    st.subheader("📋 7. KẾ HOẠCH CÔNG TÁC CHỦ NHIỆM LỚP")
    st.caption("Hệ sinh thái số hỗ trợ giáo viên chủ nhiệm quản lý lớp học và lập kế hoạch thông minh.")
    
    # Tạo các tab nhỏ bên trong tính năng chủ nhiệm
    tab_ke_hoach, tab_hoc_sinh, tab_ai_tro_ly = st.tabs([
        "📅 Kế hoạch Tuần/Tháng", 
        "👥 Quản lý Học sinh", 
        "🤖 Trợ lý AI Chủ nhiệm"
    ])
    
    # --- TAB 1: KẾ HOẠCH TUẦN/THÁNG ---
    with tab_ke_hoach:
        st.write("#### 📝 Lập kế hoạch công tác chủ nhiệm")
        col1, col2 = st.columns(2)
        with col1:
            tuan_hoc = st.selectbox("Chọn tuần học:", [f"Tuần {i}" for i in range(1, 36)])
        with col2:
            chu_de = st.text_input("Chủ đề sinh hoạt / Trọng tâm tuần:", placeholder="Ví dụ: Ổn định nề nếp, phòng chống bạo lực học đường...")
            
        muc_tieu = st.text_area("Mục tiêu cần đạt:")
        bien_phap = st.text_area("Các biện pháp thực hiện cụ thể:")
        
        if st.button("💾 Lưu kế hoạch tuần", key="btn_save_kh"):
            st.success(f"Đã lưu thành công kế hoạch {tuan_hoc}!")

    # --- TAB 2: QUẢN LÝ HỌC SINH ---
    with tab_hoc_sink_hoac_lop:
        st.write("#### 📊 Danh sách học sinh cần lưu ý & Phân công")
        
        # Bảng danh dách mẫu đơn giản
        data = {
            "STT":,
            "Họ và tên": ["Nguyễn Văn A", "Trần Thị B", "Lê Hoàng C"],
            "Chức vụ / Đặc điểm": ["Lớp trưởng", "Bí thư", "Học sinh cần hỗ trợ học tập"],
            "Ghi chú": ["Nhanh nhẹn", "Trách nhiệm", "Hay quên làm bài tập"]
        }
        df = pd.DataFrame(data)
        st.data_editor(df, num_rows="dynamic", use_container_width=True)
        st.caption("💡 Bạn có thể bấm trực tiếp vào bảng trên để sửa thông tin hoặc thêm dòng mới.")

    # --- TAB 3: TRỢ LÝ AI CHỦ NHIỆM ---
    with tab_ai_tro_ly:
        st.write("#### 🤖 Trợ lý AI tư vấn công tác chủ nhiệm")
        st.info("Nhập tình huống sư phạm hoặc yêu cầu lập kế hoạch để AI của hệ thống hỗ trợ bạn.")
        
        # Gợi ý một số câu lệnh mẫu
        st.write("**Câu hỏi gợi ý cho giáo viên:**")
        cau_mau_1 = "Lập kế hoạch sinh hoạt lớp chủ đề 'Xây dựng tình bạn đẹp, nói không với bạo lực học đường' cho học sinh THCS."
        cau_mau_2 = "Gợi ý biện pháp giáo dục một học sinh cá biệt thường xuyên đi học muộn và không làm bài tập."
        
        if st.button("💡 Dùng mẫu 1", key="btn_mau1"):
            st.session_state["chu_nhiem_ai_input"] = cau_mau_1
        if st.button("💡 Dùng mẫu 2", key="btn_mau2"):
            st.session_state["chu_nhiem_ai_input"] = cau_mau_2

        # Ô nhập liệu
        user_prompt = st.text_area(
            "Nhập yêu cầu của bạn gửi tới AI:", 
            value=st.session_state.get("chu_nhiem_ai_input", ""),
            key="txt_ai_prompt"
        )
        
        if st.button("🚀 Yêu cầu AI xử lý", type="primary", key="btn_run_ai"):
            if not user_prompt.strip():
                st.warning("Vui lòng nhập nội dung yêu cầu trước khi nhấn gửi!")
            else:
                if run_ai_prompt_safe is not None:
                    with st.spinner("AI đang suy nghĩ và lập kế hoạch cho bạn..."):
                        # Gọi hàm xử lý AI dùng chung của toàn hệ thống
                        response = run_ai_prompt_safe(user_prompt)
                        st.markdown("### 🌟 Kết quả từ Trợ lý AI:")
                        st.write(response)
                else:
                    # Trường hợp chưa kết nối được hàm AI từ app.py, giả lập câu trả lời mẫu
                    st.warning("Hệ thống chưa truyền hàm kết nối AI. Dưới đây là phản hồi mẫu:")
                    st.info("AI khuyên bạn nên: 1. Gặp riêng học sinh lắng nghe tâm tư; 2. Phối hợp chặt chẽ với phụ huynh học sinh; 3. Tổ chức trò chơi gắn kết tập thể tại lớp.")
