import streamlit as st
import io
import re
from docx import Document
from google import genai

# Hàm hỗ trợ dọn dẹp Markdown và tạo file Word
def create_word_file(title, content):
    doc = Document()
    doc.add_heading(title, 0)
    
    clean_content = re.sub(r'^#+\s*', '', content, flags=re.MULTILINE) 
    clean_content = clean_content.replace('**', '') 
    clean_content = re.sub(r'^\*\s+', '- ', clean_content, flags=re.MULTILINE) 
    clean_content = clean_content.replace('*', '') 
    
    doc.add_paragraph(clean_content)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =========================================================
# THẺ 1: CÁC SẢN PHẨM STEM
# =========================================================
def render_tab_1():
    # Dùng màu xanh dương (info) cho Thẻ 1
    st.info("💡 **THẺ 1 - BỘ LỌC TÌM KIẾM:** Chọn các tiêu chí để AI gợi ý dự án STEM.")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        chon_khoi_t1 = st.selectbox("📌 Khối lớp:", ["Khối 9", "Khối 8", "Khối 7", "Khối 6"], key="khoi_t1")
    with col_m2:
        chon_mon_t1 = st.selectbox("📌 Môn chủ đạo:", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"], key="mon_t1")
    with col_m3:
        chon_chu_de_t1 = st.selectbox("📌 Lĩnh vực:", [
            "Tiết kiệm năng lượng", "Bảo vệ môi trường", "Nông nghiệp thông minh", "Nhà thông minh", "Chủ đề tự do"
        ], key="chude_t1")
        
    mon_tich_hop_t1 = st.multiselect("➕ Các môn học cần tích hợp:", ["Toán học", "Khoa học", "Công nghệ", "Kỹ thuật", "Nghệ thuật"], key="montichhop_t1")
    tich_hop_ai_t1 = st.checkbox("🔌 Yêu cầu AI ưu tiên dự án có Vi điều khiển (Arduino/ESP8266)", value=True, key="chk_ai_t1")
    tich_hop_khuyet_tat_t1 = st.checkbox("🤝 Ưu tiên dự án phù hợp Giáo dục hòa nhập", value=True, key="chk_kt_t1")
    
    if st.button("✨ KÍCH HOẠT AI GỢI Ý CHỦ ĐỀ MỚI", use_container_width=True, key="btn_ai_t1"):
        with st.spinner("AI đang tổng hợp dữ liệu..."):
            yeu_cau_iot = "Có tích hợp AI hoặc vi điều khiển." if tich_hop_ai_t1 else "Không bắt buộc."
            mon_tich_hop = ", ".join(mon_tich_hop_t1) if mon_tich_hop_t1 else "Không yêu cầu thêm"
            
            prompt_goi_y = f"""
            Đóng vai chuyên gia STEM, đề xuất 3 dự án STEM:
            - Khối: {chon_khoi_t1} | Môn: {chon_mon_t1} | Lĩnh vực: {chon_chu_de_t1} | Tích hợp: {mon_tich_hop} | Yêu cầu: {yeu_cau_iot}
            Trình bày: Tên chủ đề, Mô tả hoạt động, Ứng dụng thực tiễn.
            """
            try:
                client = genai.Client() 
                response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt_goi_y)
                st.session_state.stem_ai_suggestions = response.text
            except Exception as e:
                st.error(f"Lỗi API: {e}")
                
    if st.session_state.stem_ai_suggestions:
        with st.container(border=True):
            st.markdown(st.session_state.stem_ai_suggestions)

# =========================================================
# THẺ 2: XÂY DỰNG DỰ ÁN GIÁO DỤC STEM
# =========================================================
def render_tab_2():
    # Dùng màu xanh lá (success) cho Thẻ 2 để tách biệt hoàn toàn
    st.success("🛠️ **THẺ 2 - THIẾT KẾ KHBD:** Sau khi chốt được chủ đề ở Thẻ 1, thầy nhập tên vào đây để AI soạn giáo án.")
    
    ten_chu_de_t2 = st.text_input("✍️ Nhập Tên dự án / Chủ đề STEM:", placeholder="Ví dụ: Hệ thống cảnh báo tốc độ xe đạp điện...", key="ten_t2")
    
    col1, col2 = st.columns(2)
    with col1:
        mon_chu_dao_t2 = st.selectbox("📚 Môn học:", ["Khoa học tự nhiên", "Toán học", "Công nghệ", "Tin học"], key="mon_t2")
        lop_hoc_t2 = st.selectbox("🎓 Dành cho:", ["Lớp 9", "Lớp 8", "Lớp 7", "Lớp 6"], key="lop_t2")
    with col2:
        thoi_luong_t2 = st.text_input("⏱️ Thời lượng:", placeholder="Ví dụ: 3 Tiết", key="thoiluong_t2")
    
    tich_hop_ai_iot_t2 = st.checkbox("🔌 Tích hợp Vi điều khiển (VD: ESP8266) vào tiến trình", value=True, key="chk_ai_t2")
    tich_hop_hoa_nhap_t2 = st.checkbox("🤝 Đưa phương pháp Giáo dục hòa nhập vào KHBD", value=True, key="chk_kt_t2")
    
    if st.button("🚀 KÍCH HOẠT AI BIÊN SOẠN KHBD CHI TIẾT", type="primary", use_container_width=True, key="btn_ai_t2"):
        if not ten_chu_de_t2:
            st.warning("Thầy vui lòng nhập Tên chủ đề STEM nhé!")
        else:
            with st.spinner("AI đang soạn Kế hoạch bài dạy chi tiết (mất khoảng 10-20 giây)..."):
                yeu_cau_iot = "Có tích hợp vi điều khiển (ESP8266, Arduino) vào thiết kế." if tich_hop_ai_iot_t2 else "Không bắt buộc."
                yeu_cau_hoa_nhap = "Ghi chú rõ việc phân hóa cho học sinh khuyết tật trong từng hoạt động." if tich_hop_hoa_nhap_t2 else ""

                prompt = f"""
                Soạn Kế hoạch bài dạy (KHBD) STEM thật chi tiết. KHÔNG dùng từ "Giáo án", chỉ dùng "KHBD".
                - Tên: {ten_chu_de_t2} | Môn: {mon_chu_dao_t2} | Lớp: {lop_hoc_t2} | Thời lượng: {thoi_luong_t2}
                - Yêu cầu: {yeu_cau_iot}. {yeu_cau_hoa_nhap}
                
                CẤU TRÚC:
                I. MỤC TIÊU (Năng lực đặc thù, Năng lực chung, Phẩm chất)
                II. THIẾT BỊ VÀ HỌC LIỆU
                III. TIẾN TRÌNH DẠY HỌC (5 bước chuẩn STEM, viết cực kỳ chi tiết Hoạt động của GV và HS):
                1. Xác định vấn đề
                2. Nghiên cứu kiến thức nền
                3. Lựa chọn giải pháp thiết kế
                4. Chế tạo và Thử nghiệm
                5. Chia sẻ, Đánh giá
                """
                try:
                    client = genai.Client()
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                    st.session_state.stem_generated_content = response.text
                    st.toast("Đã soạn KHBD xong!", icon="🎉")
                except Exception as e:
                    st.error(f"Lỗi API: {e}")

    if st.session_state.stem_generated_content != "":
        st.markdown("### 📖 KẾT QUẢ KHBD")
        with st.container(border=True):
            st.markdown(st.session_state.stem_generated_content)
        
        col_down, col_save = st.columns(2)
        with col_down:
            docx_file = create_word_file(ten_chu_de_t2, st.session_state.stem_generated_content)
            st.download_button("📥 Tải File Word (Đã xóa ký tự #)", data=docx_file, file_name=f"KHBD_STEM.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="btn_down_t2")
        with col_save:
            if st.button("💾 Lưu KHBD vào Thẻ 3", use_container_width=True, key="btn_save_t2"):
                st.session_state.stem_saved_projects[ten_chu_de_t2] = st.session_state.stem_generated_content
                st.toast("Đã lưu thành công!", icon="✅")

# =========================================================
# THẺ 3: QUẢN LÝ DỰ ÁN ĐÃ LƯU
# =========================================================
def render_tab_3():
    st.warning("📁 **THẺ 3 - KHO LƯU TRỮ:** Quản lý các KHBD thầy đã lưu.")
    if len(st.session_state.stem_saved_projects) > 0:
        for ten_da in list(st.session_state.stem_saved_projects.keys()):
            with st.expander(f"📌 {ten_da}"):
                st.markdown(st.session_state.stem_saved_projects[ten_da])
                if st.button("🗑️ Xóa", key=f"btn_del_{ten_da}"):
                    del st.session_state.stem_saved_projects[ten_da]
                    st.rerun() 
    else:
        st.info("Chưa có KHBD nào được lưu.")

# =========================================================
# HÀM CHÍNH ĐƯỢC GỌI TỪ APP.PY
# =========================================================
def render_stem_section():
    st.markdown("## 🚀 HỆ SINH THÁI GIÁO DỤC STEM")
    st.markdown("---")
    
    if "stem_generated_content" not in st.session_state:
        st.session_state.stem_generated_content = ""
    if "stem_saved_projects" not in st.session_state:
        st.session_state.stem_saved_projects = {}
    if "stem_ai_suggestions" not in st.session_state:
        st.session_state.stem_ai_suggestions = ""

    tab1, tab2, tab3 = st.tabs(["💡 1. SẢN PHẨM STEM", "🛠️ 2. XÂY DỰNG KHBD", "📁 3. KHBD ĐÃ LƯU"])

    with tab1:
        render_tab_1()
    with tab2:
        render_tab_2()
    with tab3:
        render_tab_3()
