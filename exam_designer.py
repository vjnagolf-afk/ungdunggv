# exam_designer.py - ĐOẠN 1: CẤU HÌNH & GOOGLE SHEETS
import streamlit as st
import gspread
from document_processor import read_uploaded_docx, read_uploaded_pdf, export_to_docx_vietnam_standard

# Khai báo thông tin cấu hình Google Sheets của bạn
SPREADSHEET_ID = "1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY"
TAB_NAME = "DE_KT"

def sync_exam_to_google_sheet(ten_de, mon, khoi, thoi_gian, noi_dung):
    """Tự động tìm kiếm Service Account trong Secrets và đồng bộ kết quả lên Google Sheets"""
    try:
        creds_dict = None
        priority_keys = ["gspread_credentials", "GSPREAD_CREDENTIALS", "google_sheet_creds", "google_creds", "GOOGLE_KEY"]
        for key in priority_keys:
            if key in st.secrets:
                creds_dict = st.secrets[key]
                break
                
        if creds_dict is None:
            for key in st.secrets.keys():
                node = st.secrets[key]
                if hasattr(node, "get") or isinstance(node, dict):
                    if node.get("type") == "service_account" or "private_key" in node:
                        creds_dict = node
                        break
        
        if creds_dict is None:
            return False
            
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_key(SPREADSHEET_ID)
        worksheet = sh.worksheet(TAB_NAME)
        
        if len(worksheet.get_all_values()) == 0:
            worksheet.append_row(["Tên Đề", "Môn Học", "Khối Lớp", "Thời Gian", "Nội Dung Đề Thi"])
            
        worksheet.append_row([ten_de, mon, khoi, thoi_gian, noi_dung])
        return True
    except:
        return False
# exam_designer.py - ĐOẠN 2: KHỞI TẠO STATE & KHUNG HÀNG 1, HÀNG 2 (CẬP NHẬT 2 NÚT TẢI)
def render_exam_designer_section(run_ai_prompt_safe_func):
    """Giao diện phân hệ thiết kế đề kiểm tra - Tách biệt nút tải Đề cương và Ma trận"""
    
    if "ket_qua_de" not in st.session_state: st.session_state["ket_qua_de"] = ""
    if "model_dung" not in st.session_state: st.session_state["model_dung"] = ""
    
    # Lưu thông tin cấu hình vào session_state để tránh lỗi NameError
    if "save_ten_de" not in st.session_state: st.session_state["save_ten_de"] = "Đề kiểm tra giữa học kỳ I"
    if "save_school" not in st.session_state: st.session_state["save_school"] = "TRƯỜNG THCS NGUYỄN CHÍ THANH"
    if "save_mon_hoc" not in st.session_state: st.session_state["save_mon_hoc"] = "Khoa học tự nhiên"
    if "save_khoi_lop" not in st.session_state: st.session_state["save_khoi_lop"] = "Lớp 8"
    if "save_thoi_gian" not in st.session_state: st.session_state["save_thoi_gian"] = "45 phút"

    tab_tao_de, tab_thu_muc = st.tabs(["🔴 CHỨC NĂNG TẠO ĐỀ KIỂM TRA AI", "🗂️ THƯ MỤC ĐỀ ĐÃ XÂY DỰNG"])
    
    with tab_tao_de:
        # --- HÀNG CẤU HÌNH CHUNG & UPLOAD TÀI LIỆU ---
        col_cấu_hình, col_tài_liệu = st.columns([1.1, 1.0])
        
        with col_cấu_hình:
            hinh_thuc = st.selectbox("Hình thức đề:", ["Trắc nghiệm kết hợp tự luận", "100% Trắc nghiệm", "100% Tự luận"])
            st.session_state["save_mon_hoc"] = st.text_input("Môn học:", value=st.session_state["save_mon_hoc"])
            st.session_state["save_khoi_lop"] = st.selectbox("Khối lớp:", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Khối 10", "Khối 11", "Khối 12"], index=2)
            st.session_state["save_thoi_gian"] = st.text_input("Thời gian:", value=st.session_state["save_thoi_gian"])
            st.session_state["save_school"] = st.text_input("Tên trường hành chính:", value=st.session_state["save_school"])
            
        with col_tài_liệu:
            # 🚀 NÚT TẢI 1: ĐỀ CƯƠNG / TÀI LIỆU NỘI DUNG KIẾN THỨC
            st.markdown("**1. TẢI TÀI LIỆU / ĐỀ CƯƠNG KIẾN THỨC YÊU CẦU BÁM SÁT RA ĐỀ:**")
            file_de_cuong = st.file_uploader("Upload Đề Cương (.docx, .pdf):", type=["docx", "pdf"], key="exam_upload_de_cuong_2026")
            
            noi_dung_de_cuong = ""
            if file_de_cuong:
                if file_de_cuong.name.endswith(".docx"):
                    noi_dung_de_cuong = read_uploaded_docx(file_de_cuong)
                else:
                    noi_dung_de_cuong = read_uploaded_pdf(file_de_cuong)
                st.success(f"✅ Đã nạp tài liệu đề cương: {file_de_cuong.name}")
                
            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            
            # 🚀 NÚT TẢI 2: KHUNG CẤU TRÚC MA TRẬN / ĐẶC TẢ ĐỀ THI
            st.markdown("**2. TẢI KHUNG CẤU TRÚC MA TRẬN / ĐẶC TẢ CHI TIẾT:**")
            file_ma_tran = st.file_uploader("Upload Khung Ma Trận (.docx, .pdf):", type=["docx", "pdf"], key="exam_upload_ma_tran_2026")
            
            noi_dung_ma_tran = ""
            if file_ma_tran:
                if file_ma_tran.name.endswith(".docx"):
                    noi_dung_ma_tran = read_uploaded_docx(file_ma_tran)
                else:
                    noi_dung_ma_tran = read_uploaded_pdf(file_ma_tran)
                st.success(f"✅ Đã nạp cấu trúc ma trận: {file_ma_tran.name}")

        st.markdown("---")

        # --- HÀNG HAI KHỐI: PHẦN TRẮC NGHIỆM & PHẦN TỰ LUẬN ---
        col_tn, col_tl = st.columns([1.1, 1.0])
        
        with col_tn:
            st.markdown("<h4 style='text-align: center; background-color: #FFE4E6; color: #991B1B; padding: 5px; border-radius: 4px; margin-bottom: 15px;'>PHẦN TRẮC NGHIỆM</h4>", unsafe_allow_html=True)
            col_tn_trai, col_tn_phai = st.columns([1.2, 1.0])
            with col_tn_trai:
                st.markdown("**Tổng số câu TNKQ:**")
                sc_nhieu_lua_chon = st.number_input("Số câu nhiều lựa chọn:", min_value=0, max_value=50, value=12)
                sc_dung_sai = st.number_input("Số câu đúng sai:", min_value=0, max_value=50, value=2)
                sc_dien_khuyết = st.number_input("Số câu điền khuyết:", min_value=0, max_value=50, value=0)
                sc_tra_loi_ngan = st.number_input("Số câu trả lời ngắn:", min_value=0, max_value=50, value=0)
            with col_tn_phai:
                tong_diem_tn = st.number_input("Tổng điểm TN:", value=4.00, step=0.25)
                d_nhieu_lua_chon = st.number_input("Tổng điểm dòng này:", value=3.00, step=0.25, key="d_nlc")
                d_dung_sai = st.number_input("Tổng điểm dòng này:", value=1.00, step=0.25, key="d_ds")
                d_dien_khuyết = st.number_input("Tổng điểm dòng này:", value=0.00, step=0.25, key="d_dk")
                d_tra_loi_ngan = st.number_input("Tổng điểm dòng này:", value=0.00, step=0.25, key="d_tln")
                
            tong_cau_tn = sc_nhieu_lua_chon + sc_dung_sai + sc_dien_khuyết + sc_tra_loi_ngan

        with col_tl:
            st.markdown("<h4 style='text-align: center; background-color: #DCFCE7; color: #166534; padding: 5px; border-radius: 4px; margin-bottom: 15px;'>PHẦN TỰ LUẬN</h4>", unsafe_allow_html=True)
            col_tl_trai, col_tl_phai = st.columns(2)
            with col_tl_trai:
                sc_tu_luan = st.number_input("TỔNG SỐ CÂU TỰ LUẬN:", min_value=0, max_value=20, value=3)
            with col_tl_phai:
                tong_diem_tl = st.number_input("ĐIỂM TỔNG:", value=6.00, step=0.25)
            
            st.markdown("<p style='font-weight: bold; margin-top: 10px; margin-bottom: 2px;'>Chỉ định điểm chi tiết từng câu:</p>", unsafe_allow_html=True)
            diem_chi_tiet_tl = []
            cols_diem_tl = st.columns(3)
            for i in range(int(sc_tu_luan)):
                col_idx = i % 3
                with cols_diem_tl[col_idx]:
                    diem_cau = st.number_input(f"Câu {i+1}:", min_value=0.0, max_value=10.0, value=2.0 if i < 3 else 1.0, step=0.5, key=f"tl_c_{i}")
                    diem_chi_tiet_tl.append((i+1, diem_cau))
# exam_designer.py - ĐOẠN 3 (Cập nhật Ép AI sinh đầy đủ Ma trận & Đặc tả)
        st.markdown("---")

        st.markdown("**Tỷ lệ mức độ nhận thức (%):**")
        col_nb, col_th, col_vd, col_vdc = st.columns(4)
        with col_nb:
            tl_nhan_biet = st.number_input("Nhận biết:", min_value=0, max_value=100, value=40, step=10)
        with col_th:
            tl_thong_hieu = st.number_input("Thông hiểu:", min_value=0, max_value=100, value=30, step=10)
        with col_vd:
            tl_van_dung = st.number_input("Vận dụng:", min_value=0, max_value=100, value=20, step=10)
        with col_vdc:
            tl_vd_cao = st.number_input("Vận dụng cao:", min_value=0, max_value=100, value=10, step=10)

        col_input_name, col_input_model = st.columns([2.0, 1.0])
        with col_input_name:
            st.session_state["save_ten_de"] = st.text_input("Tên bài kiểm tra / Đề số:", value=st.session_state["save_ten_de"])
        with col_input_model:
            mo_hinh_uu_tien = st.selectbox("Mô hình xử lý đề thi:", ["3.1 Flash-Lite", "3.5 Flash", "3.1 Pro", "Tư duy mở rộng"], index=0)

        yeu_cau_khac = st.text_area("Nhập yêu cầu khác (Tùy chọn):", placeholder="Ví dụ: Đề thi có chứa 1 câu hỏi thực tế về đồ thị hàm số bậc nhất...")

        col_btn_l, col_btn_r = st.columns([1.5, 1.0])
        with col_btn_l:
            nut_sinh_de = st.button("🔴 Tự động khởi tạo ma trận và đề thi", type="primary", use_container_width=True)
        with col_btn_r:
            st.write(""); st.write("")
            st.checkbox("Yêu cầu bám sát kiến thức trong GIÁO TRÌNH, tài liệu", value=True)

        if nut_sinh_de:
            if int(tl_nhan_biet + tl_thong_hieu + tl_van_dung + tl_vd_cao) != 100:
                st.error("⚠️ Tổng tỷ lệ mức độ nhận thức phải bằng 100%!")
            else:
                with st.spinner("🧠 Hệ thống đang lập bảng Ma trận, Bản đặc tả và dựng đề thi..."):
                    chuỗi_điểm_tl = ", ".join([f"Câu {c_id} ({d}đ)" for c_id, d in diem_chi_tiet_tl])
                    
                    # 🚀 CẢI TIẾN PROMPT: Ép buộc tuyệt đối cấu trúc xuất ra phải chứa bảng ma trận và đặc tả
                    prompt_vi_mo = f"""
                    Bạn là Chuyên gia Khảo thí và Kiểm định Chất lượng Giáo dục phổ thông tại Việt Nam. 
                    Hãy thiết kế một tài liệu kiểm tra hoàn chỉnh cho môn {st.session_state['save_mon_hoc']} - {st.session_state['save_khoi_lop']} (Thời gian làm bài: {st.session_state['save_thoi_gian']}).
                    
                    CẤU TRÚC ĐẦU RA BẮT BUỘC THEO THỨ TỰ (KHÔNG ĐƯỢC BỎ SÓT):
                    
                    1. KHUNG MA TRẬN ĐỀ KIỂM TRA: Thiết lập dưới dạng bảng Markdown bằng ký tự '|'. Các cột bao gồm: Nội dung kiến thức, Đơn vị kiến thức, Số câu NB, Số câu TH, Số câu VD, Số câu VDC, Tổng số câu, Tỷ lệ (%).
                    2. BẢNG ĐẶC TẢ KỸ THUẬT ĐỀ KIỂM TRA: Thiết lập dưới dạng bảng Markdown bằng ký tự '|'. Các cột bao gồm: Nội dung kiến thức, Đơn vị kiến thức, Mức độ đánh giá, Số câu hỏi theo ma trận.
                    3. NỘI DUNG ĐỀ KIỂM TRA CHI TIẾT:
                       - Hình thức đề: {hinh_thuc}
                       - Phần Trắc nghiệm ({tong_diem_tn} điểm, {tong_cau_tn} câu): {sc_nhieu_lua_chon} câu nhiều lựa chọn, {sc_dung_sai} câu đúng sai.
                       - Phần Tự luận ({tong_diem_tl} điểm): Phân bổ câu {chuỗi_điểm_tl}
                       - Tỷ lệ nhận thức: Nhận biết {tl_nhan_biet}%, Thông hiểu {tl_thong_hieu}%, Vận dụng {tl_van_dung}%, Vận dụng cao {tl_vd_cao}%.
                       - Yêu cầu bổ sung: {yeu_cau_khac}
                    
                    [FILE ĐỀ CƯƠNG KIẾN THỨC GỐC]:
                    {noi_dung_de_cuong if 'noi_dung_de_cuong' in locals() and noi_dung_de_cuong else "Theo phân phối chương trình học chuẩn."}
                    
                    [FILE MA TRẬN ĐỂ BẮT CHƯỚC (NẾU CÓ)]:
                    {noi_dung_ma_tran if 'noi_dung_ma_tran' in locals() and noi_dung_ma_tran else "Tự sinh dựa theo số lượng câu trên giao diện."}
                    
                    QUY ĐỊNH ĐỊNH DẠNG: Công thức toán đặt trong $...$, các đáp án trắc nghiệm A., B., C., D. bắt buộc viết tách dòng độc lập.
                    """
                    ket_qua, model_thuc_te = run_ai_prompt_safe_func(prompt_vi_mo, mo_hinh_uu_tien)
                    st.session_state["ket_qua_de"] = ket_qua
                    st.session_state["model_dung"] = model_thuc_te
                    
                    if "⚠️" not in ket_qua and "Lỗi" not in ket_qua:
                        success = sync_exam_to_google_sheet(
                            st.session_state["save_ten_de"], 
                            st.session_state["save_mon_hoc"], 
                            st.session_state["save_khoi_lop"], 
                            st.session_state["save_thoi_gian"], 
                            ket_qua
                        )
                        if success:
                            st.toast("🎉 Đã lưu trữ và đồng bộ đề thi thành công lên Google Sheet!", icon="🚀")

        if st.session_state["ket_qua_de"]:
            st.info(f"🤖 Đề thi được xây dựng thành công bằng mô hình thực tế: `{st.session_state['model_dung']}`")
            st.markdown("### 📝 Xem trước nội dung đề thi:")
            st.markdown(st.session_state["ket_qua_de"])
            
            data_word = export_to_docx_vietnam_standard(
                st.session_state["ket_qua_de"], 
                st.session_state["save_ten_de"], 
                school_name=st.session_state["save_school"]
            )
            st.download_button(
                label="📥 Tải tệp Đề kiểm tra Word (.docx) bản chuẩn hành chính",
                data=data_word,
                file_name=f"De_Kiem_Tra_{st.session_state['save_khoi_lop'].replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
    with tab_thu_muc:
        st.write("📂 Danh sách các đề kiểm tra đã đồng bộ và lưu trữ thành công:")
        st.markdown(f"🔗 [Bấm vào đây để mở trực tiếp Google Sheets quản lý đề thi](https://google.com{SPREADSHEET_ID})")

