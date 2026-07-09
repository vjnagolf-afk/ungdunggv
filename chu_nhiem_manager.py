import streamlit as st

def render_chu_nhiem_section(run_ai_prompt_safe=None):
    st.subheader("📋 7. KẾ HOẠCH CÔNG TÁC CHỦ NHIỆM LỚP")
    st.caption("Hệ sinh thái số hỗ trợ giáo viên chủ nhiệm quản lý lớp học và lập kế hoạch thông minh.")
    
    # Ô nhập liệu yêu cầu gửi tới AI
    user_prompt = st.text_area(
        "Nhập tình huống sư phạm hoặc yêu cầu lập kế hoạch chủ nhiệm:", 
        placeholder="Ví dụ: Gợi ý biện pháp giáo dục một học sinh thường xuyên đi học muộn...",
        key="txt_chu_nhiem_prompt"
    )
    
    if st.button("🚀 Trợ lý AI xử lý", type="primary", key="btn_chu_nhiem_ai"):
        if not user_prompt.strip():
            st.warning("Vui lòng nhập nội dung yêu cầu trước khi nhấn nút!")
        else:
            if run_ai_prompt_safe is not None:
                with st.spinner("AI đang suy nghĩ và lập kế hoạch cho bạn..."):
                    response = run_ai_prompt_safe(user_prompt)
                    st.markdown("### 🌟 Kết quả từ Trợ lý AI:")
                    st.write(response)
            else:
                st.info("Hệ thống kết nối AI đang được đồng bộ...")
