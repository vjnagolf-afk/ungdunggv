import streamlit as st  
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import json
import pandas as pd

# --- HÀM TẠO LƯỚI VIỀN CHO BẢNG WORD KHÔNG BỊ MẤT KHUNG ---
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# --- HÀM XUẤT PHÔI WORD CÓ BẢNG BIỂU CHUẨN PHỤ LỤC III ---
def export_plan_to_docx_with_table(teacher_name, subject_name, content_data):
    doc = docx.Document()
    
    # Cấu hình căn lề chuẩn Phụ lục III
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)
        
    # Tiêu đề Quốc hiệu
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_qg = p_header.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n")
    run_qg.bold = True
    run_qg.font.name = 'Times New Roman'
    run_qg.font.size = Pt(13)
    
    # Tên văn bản mẫu
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("KẾ HOẠCH GIÁO DỤC CỦA GIÁO VIÊN\n(PHỤ LỤC III CÔNG VĂN 5512/BGDĐT)\n")
    r_title.bold = True
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(14)
    
    # Thông tin cá nhân giáo viên
    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_info = p_info.add_run(f"Giáo viên thực hiện: {teacher_name}\nMôn học/Hoạt động giáo dục: {subject_name}\n\n")
    r_info.font.name = 'Times New Roman'
    r_info.font.size = Pt(14)
    
    # Mục I. Kế hoạch dạy học phân phối chương trình
    p_section1 = doc.add_paragraph()
    r_sec1 = p_section1.add_run("I. KẾ HOẠCH DẠY HỌC PHÂN PHỐI CHƯƠNG TRÌNH")
    r_sec1.bold = True
    r_sec1.font.name = 'Times New Roman'
    r_sec1.font.size = Pt(14)
    
    # Khởi tạo bảng biểu có lưới ô (Table Grid)
    headers = ["STT", "Tuần", "Bài học / Chủ đề", "Số tiết", "Thiết bị dạy học"]
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    
    # Thiết lập tiêu đề cột trong Word
    hdr_cells = table.rows.cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        set_cell_margins(hdr_cells[i])
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)

    # Đổ dữ liệu từ mảng tuần tự do AI phân bổ vào bảng Word
    if isinstance(content_data, list):
        for idx, item in enumerate(content_data, 1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(idx)
            row_cells[1].text = str(item.get("Tuan", ""))
            row_cells[2].text = str(item.get("BaiHoc", ""))
            row_cells[3].text = str(item.get("SoTiet", ""))
            row_cells[4].text = str(item.get("ThietBi", "-"))
            
            # Định dạng lại phông chữ căn chỉnh ô dòng dữ liệu
            for cell in row_cells:
                set_cell_margins(cell)
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
                        
    # Mục II. Các nhiệm vụ khác
    doc.add_paragraph()
    p_section2 = doc.add_paragraph()
    r_sec2 = p_section2.add_run("\nII. CÁC NHIỆM VỤ KHÁC ĐƯỢC GIAO")
    r_sec2.bold = True
    r_sec2.font.name = 'Times New Roman'
    r_sec2.font.size = Pt(14)
    
    p_other = doc.add_paragraph()
    r_other = p_other.add_run("- Bồi dưỡng học sinh giỏi và phụ đạo học sinh yếu kém theo kế hoạch nhà trường.\n- Tham gia sinh hoạt chuyên môn tổ, nhóm đầy đủ và đúng quy định.\n- Thực hiện công tác chủ nhiệm lớp và các hoạt động giáo dục ngoại khóa được phân công.")
    r_other.font.name = 'Times New Roman'
    r_other.font.size = Pt(13)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
# --- GIAO DIỆN PHÂN HỆ KẾ HOẠCH CÁ NHÂN ---
def render_personal_plan(run_ai_handler=None):
    st.markdown("<h3 style='text-align: left; color: #1E3A8A;'>🗓️ TRỢ LÝ XÂY DỰNG KẾ HOẠCH GIÁO DỤC CÁ NHÂN (PHỤ LỤC III)</h3>", unsafe_allow_html=True)
    
    # Khởi tạo kho lưu trữ nội bộ hệ thống nếu chưa có dữ liệu nền
    if "db_ke_hoach_da_luu" not in st.session_state:
        st.session_state["db_ke_hoach_da_luu"] = {}
        
    # --- KHU VỰC PHÂN QUYỀN MẬT KHẨU ADMIN TẠI SIDEBAR ---
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
    vai_tro = st.sidebar.radio("Vai trò", ["Giáo viên bộ môn", "Tổ trưởng chuyên môn (Admin)", "Ban giám hiệu"], label_visibility="collapsed", key="vai_tro_plan_sidebar_fixed_v8")
    
    is_admin = False
    if vai_tro == "Tổ trưởng chuyên môn (Admin)":
        ma_pin = st.sidebar.text_input("Nhập mã pin quản lý Admin:", type="password", value="", key="pin_admin_plan_v8")
        if ma_pin == "123456":
            st.sidebar.success("✅ Quyền Admin đã mở.")
            is_admin = True
        elif ma_pin != "": 
            st.sidebar.error("❌ Mã PIN sai.")
            return

    # Biến nhớ lưu tạm kết quả đầu ra cho phiên hoạt động hiện tại
    if "current_ai_output_raw" not in st.session_state: st.session_state["current_ai_output_raw"] = None
    if "current_teacher_active" not in st.session_state: st.session_state["current_teacher_active"] = ""
    if "current_subject_active" not in st.session_state: st.session_state["current_subject_active"] = ""

    with st.form("form_personal_plan_fixed_final_v8", border=False):
        col_t, col_s = st.columns(2)
        t_name = col_t.text_input("Họ và tên Giáo viên giảng dạy:", placeholder="Ví dụ: Thầy Lê Hồng Dưỡng", key="plan_txt_t_name_v8")
        s_name = col_s.selectbox("Môn học / Phân môn phụ trách:", ["Khoa học tự nhiên (Vật lý)", "Khoa học tự nhiên (Sinh học)", "Khoa học tự nhiên (Hóa học)", "Toán học", "Ngữ văn", "GDTC"], key="plan_sb_s_name_v8")
        col_g, col_w = st.columns(2)
        grade_target = col_g.text_input("Khối lớp phân công dạy:", placeholder="Ví dụ: Khối 7, Khối 8", key="plan_txt_grade_target_v8")
        week_count = col_w.text_input("Tổng số tuần thực hiện kế hoạch:", placeholder="Ví dụ: 35 Tuần", key="plan_txt_week_count_v8")
        st.markdown("**💬 Các tiêu chí đặc thù hoặc lưu ý phân bổ tiết (Nếu có):**")
        note_plan = st.text_area("Yêu cầu bổ sung cho AI:", placeholder="Ví dụ: Phân bổ chi tiết số tiết cho chương Tốc độ ở vật lý 7 học kỳ I...", label_visibility="collapsed", key="plan_ta_note_v8")
        
        if not is_admin:
            st.warning("⚠️ Chức năng Lập kế hoạch tự động bằng AI yêu cầu quyền tài khoản Tổ trưởng chuyên môn (Admin). Vui lòng xác thực mã PIN ở thanh bên (Sidebar).")
            run_ai_plan = st.form_submit_button(" Khởi tạo Kế hoạch bằng AI", type="primary", use_container_width=True, disabled=True)
        else:
            run_ai_plan = st.form_submit_button(" Khởi tạo Kế hoạch bằng AI", type="primary", use_container_width=True, disabled=False)
    if run_ai_plan and is_admin:
        if not t_name or not grade_target:
            st.warning("⚠️ Vui lòng điền Họ tên giáo viên và Khối lớp để AI lập kế hoạch!")
        else:
            with st.spinner("Trợ lý AI đang lập tiến trình và phân tách dữ liệu bảng biểu Phụ lục III..."):
                # Ép cấu trúc đầu ra JSON để dựng bảng Word không bao giờ bị lệch dòng kẻ gãy cột
                prompt_plan = (
                    f"Hãy soạn thảo phân bổ tiến trình dạy học cho giáo viên: {t_name}, môn: {s_name}, khối: {grade_target} trong {week_count}. "
                    f"Yêu cầu bổ sung: {note_plan}. Trả về kết quả dưới dạng cấu trúc mảng JSON thuần túy gồm danh sách các đối tượng, "
                    f"không bao bọc bằng ký tự markdown nào khác ngoài thẻ mở và thẻ đóng json. Mỗi đối tượng có cấu trúc chính xác các khóa sau: "
                    f'[{{"Tuan": "Tuần thực hiện", "BaiHoc": "Tên bài học/Chủ đề chi tiết", "SoTiet": "Số tiết", "ThietBi": "Thiết bị học tập sử dụng"}}]'
                )
                
                if run_ai_handler is not None:
                    res_plan, status = run_ai_handler(prompt_plan)
                    if status == "error":
                        st.error(f"❌ Hệ thống kết nối AI gặp sự cố kỹ thuật: {res_plan}")
                    else:
                        try:
                            # Làm sạch đầu ra chuỗi json thừa nếu có
                            clean_json = res_plan.replace("```json", "").replace("```", "").strip()
                            parsed_data = json.loads(clean_json)
                            
                            # Lưu vào kho lưu trữ tạm thời của phiên chạy hiện tại
                            st.session_state["current_ai_output_raw"] = parsed_data
                            st.session_state["current_teacher_active"] = t_name
                            st.session_state["current_subject_active"] = s_name
                            
                            # Tự động đồng bộ lưu trữ vào Cơ sở dữ liệu trường học (Hệ thống nhớ của app)
                            record_id = f"{t_name.replace(' ', '_')}_{s_name.replace(' ', '_')}"
                            st.session_state["db_ke_hoach_da_luu"][record_id] = {
                                "teacher": t_name,
                                "subject": s_name,
                                "grade": grade_target,
                                "weeks": week_count,
                                "data": parsed_data
                            }
                            st.success("🎉 Khởi tạo và Lưu trữ kế hoạch thành công!")
                        except Exception as parse_err:
                            st.error("⚠️ AI phản hồi cấu hình không chuẩn bảng. Vui lòng bấm thử lại.")
                else:
                    st.error("❌ Lỗi cấu hình hệ thống: Không tìm thấy trình điều khiển AI tập trung.")

    # --- KHU VỰC HIỂN THỊ KẾT QUẢ VÀ TRÍCH XUẤT FILE ---
    st.markdown("---")
    st.markdown("### 📊 Nội dung Kế hoạch Giáo dục sinh bởi AI:")
    
    if st.session_state["current_ai_output_raw"]:
        # Tạo bảng Streamlit hiển thị trực quan dữ liệu ngăn nắp ngay trên web
        df_display = pd.DataFrame(st.session_state["current_ai_output_raw"])
        df_display.index = df_display.index + 1
        st.dataframe(df_display, use_container_width=True)
        
        # Gọi hàm xuất Word lưới kẻ bảng chuẩn chỉ
        word_plan_data = export_plan_to_docx_with_table(
            st.session_state["current_teacher_active"], 
            st.session_state["current_subject_active"], 
            st.session_state["current_ai_output_raw"]
        )
        st.download_button(
            label="📥 Tải file Word (.docx) Phụ lục III có bảng biểu lưới chuẩn về máy", 
            data=word_plan_data, 
            file_name=f"Phu_Luc_III_{st.session_state['current_teacher_active'].replace(' ', '_')}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            use_container_width=True
        )
    else:
        st.caption("Khung kế hoạch chi tiết dạng bảng biểu sẽ xuất hiện tại đây sau khi Tổ trưởng bấm nút khởi tạo...")

    # --- KHU VỰC QUẢN LÝ LỊCH SỬ KẾ HOẠCH ĐÃ LƯU TRONG TOÀN HỆ THỐNG ---
    if st.session_state["db_ke_hoach_da_luu"]:
        st.markdown("---")
        st.markdown("### 🗂️ Danh sách Kế hoạch Giáo dục đã lưu trên Hệ thống:")
        for key_id, info in st.session_state["db_ke_hoach_da_luu"].items():
            with st.expander(f"📋 Bản kế hoạch môn {info['subject']} - GV: {info['teacher']} (Khối {info['grade']})"):
                st.write(f"**Tổng số tuần**: {info['weeks']}")
                df_history = pd.DataFrame(info['data'])
                st.dataframe(df_history, use_container_width=True)
                
                # Nút tải lại Word riêng cho từng bản ghi cũ trong lịch sử lưu trữ
                word_history_data = export_plan_to_docx_with_table(info['teacher'], info['subject'], info['data'])
                st.download_button(
                    label=f"📥 Tải file Word của GV {info['teacher']}", 
                    data=word_history_data, 
                    file_name=f"Phu_Luc_III_{info['teacher'].replace(' ', '_')}.docx", 
                    key=f"dl_hist_{key_id}",
                    use_container_width=True
                )
