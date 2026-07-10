st.markdown("---")
    
    danh_sach_file = ["KH Mẫu", "Thông tin HSKT", "Nội dung học tập HK"]
    
    for label in danh_sach_file:
        # Cấu trúc: Nhãn(2) | Upload(6) | Lưu(1)
        # Việc dùng tỷ lệ [2, 6, 1] sẽ tạo sự cân đối
        c_label, c_up, c_save = st.columns([2, 6, 1])
        
        c_label.markdown(f"**{label}**")
        
        # Tạo file_uploader - Đã ẩn label để tiết kiệm không gian
        uploaded = c_up.file_uploader(f"...", key=f"up_{label}", label_visibility="collapsed")
        
        # Nút Lưu - Đặt trong cột nhỏ nhất
        if c_save.button("💾 Lưu", key=f"save_{label}", use_container_width=True):
            if uploaded:
                st.session_state["my_files"][label] = uploaded
                st.toast(f"Đã lưu: {label}")
            else:
                st.warning(f"Chọn file trước!")
        
        # Trạng thái file
        if label in st.session_state["my_files"]:
            st.caption(f"✅ Đang dùng tệp: {st.session_state['my_files'][label].name}")

    st.markdown("---")
