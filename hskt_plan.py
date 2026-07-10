st.markdown("---")
    
    # Định nghĩa danh sách các loại file
    danh_sach_file = ["KH Mẫu", "Thông tin HSKT", "Nội dung học tập HK"]
    
    for label in danh_sach_file:
        # Chia cột: Nhãn(1.5) | Upload(4) | Lưu(1)
        c_label, c_up, c_save = st.columns([1.5, 4, 1])
        
        c_label.markdown(f"**{label}**")
        
        # Tạo file_uploader
        uploaded = c_up.file_uploader(f"Chọn file:", key=f"up_{label}", label_visibility="collapsed")
        
        # Nút Lưu
        if c_save.button("💾 Lưu", key=f"save_{label}", use_container_width=True):
            if uploaded:
                st.session_state["my_files"][label] = uploaded
                st.toast(f"Đã lưu: {label}")
            else:
                st.warning(f"Vui lòng chọn file cho {label}")
        
        # Hiển thị trạng thái file đã lưu
        if label in st.session_state["my_files"]:
            st.caption(f"✅ Đang dùng tệp: {st.session_state['my_files'][label].name}")

    st.markdown("---")
