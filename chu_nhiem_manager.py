import streamlit as st

def render_chu_nhiem_section(run_ai_prompt_safe=None):
    st.subheader("📋 7. KẾ HOẠCH CÔNG TÁC CHỦ NHIỆM LỚP")
    st.caption("Hệ sinh thái số hỗ trợ giáo viên chủ nhiệm quản lý lớp học và lập kế hoạch thông minh.")
    
    st.write("---")
    st.write("#### 🛠 CẤU HÌNH THÔNG TIN CHỦ NHIỆM")
    
    # Tạo 3 cột nằm ngang để chứa 3 menu xổ xuống
    col_khoi, col_lop, col_thang = st.columns(3)
    
    with col_khoi:
        # 1. Menu chọn Khối lớp
        khoi_options = ["Khối lớp 6", "Khối lớp 7", "Khối lớp 8", "Khối lớp 9"]
        selected_khoi = st.selectbox("Chọn Khối lớp:", khoi_options, key="sb_khoi_lop")
        
    with col_lop:
        # 2. Logic tự động thay đổi danh sách lớp tương ứng theo khối đã chọn
        if selected_khoi == "Khối lớp 6":
            lop_options = ["6A", "6B", "6C", "6D", "6E", "6F"]
        elif selected_khoi == "Khối lớp 7":
            lop_options = ["7A", "7B", "7C", "7D", "7E", "7F"]
        elif selected_khoi == "Khối lớp 8":
            lop_options = ["8A", "8B", "8C", "8D", "8E", "8F"]
        else: # Khối lớp 9
            lop_options = ["9A", "9B", "9C", "9D", "9E", "9F", "9G"]
            
        selected_lop = st.selectbox("Chọn Lớp chủ nhiệm:", lop_options, key="sb_lop_chu_nhiem")
        
    with col_thang:
        # 3. Menu chọn tháng (Từ tháng 8/2026 đến tháng 5/2027)
        thang_options = [
            "Tháng 8/2026", "Tháng 9/2026", "Tháng 10/2026", "Tháng 11/2026", "Tháng 12/2026",
            "Tháng 1/2027", "Tháng 2/2027", "Tháng 3/2027", "Tháng 4/2027", "Tháng 5/2027"
        ]
        selected_thang = st.selectbox("Chọn Tháng công tác:", thang_options, key="sb_thang_cong_tac")
        
    st.write("---")
    st.write(f"#### 🤖 Trợ lý AI lập kế hoạch cho lớp **{selected_lop}** - **{selected_thang}**")
    
    # Tự động chèn thông tin cấu hình vào ô nhập liệu để AI đọc hiểu ngữ cảnh
    context_prefix = f"Tôi là giáo viên chủ nhiệm lớp {selected_lop} ({selected_khoi}). Hãy giúp tôi lập kế hoạch công tác chủ nhiệm cho {selected_thang} về việc: "
    
    user_prompt = st.text_area(
        "Nhập tình huống sư phạm hoặc nội dung trọng tâm cần AI hỗ trợ:", 
        placeholder="Ví dụ: Ổn định nề nếp đầu năm học, giải quyết tình trạng học sinh đi học muộn...",
        key="txt_chu_nhiem_prompt"
    )
    
    if st.button("🚀 Trợ lý AI xử lý", type="primary", key="btn_chu_nhiem_ai"):
        if not user_prompt.strip():
            st.warning("Vui lòng nhập nội dung yêu cầu trước khi nhấn nút!")
        else:
            if run_ai_prompt_safe is not None:
                with st.spinner("AI đang suy nghĩ và thiết kế kế hoạch chi tiết cho lớp của bạn..."):
                    # Gộp ngữ cảnh chọn lớp/tháng vào câu lệnh gửi tới AI
                    full_prompt = context_prefix + user_prompt
                    response = run_ai_prompt_safe(full_prompt)
                    
                    st.markdown(f"### 🌟 Kế hoạch công tác chủ nhiệm lớp {selected_lop} ({selected_thang})")
                    st.write(response)
            else:
                st.info("Hệ thống kết nối AI đang được đồng bộ...")
