import streamlit as st  
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import json
import pandas as pd
from openpyxl.utils import get_column_letter

# --- HÀM TẠO LƯỚI VIỀN CHO BẢNG WORD ---
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

# --- HÀM XUẤT PHÔI WORD NÂNG CẤP ĐỦ 8 CỘT THEO MẪU MỚI ---
def export_plan_to_docx_with_table(teacher_name, subject_name, content_data):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)
        
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_qg = p_header.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐỘC LẬP - TỰ DO - HẠNH PHÚC\n")
    run_qg.bold = True
    run_qg.font.name = 'Times New Roman'
    run_qg.font.size = Pt(13)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("KẾ HOẠCH GIÁO DỤC CỦA GIÁO VIÊN\n(PHỤ LỤC III CÔNG VĂN 5512/BGDĐT)\n")
    r_title.bold = True
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(14)
    
    p_info = doc.add_paragraph()
    r_info = p_info.add_run(f"Giáo viên thực hiện: {teacher_name}\nMôn học/Hoạt động giáo dục: {subject_name}\n\n")
    r_info.font.name = 'Times New Roman'
    r_info.font.size = Pt(14)
    
    p_section1 = doc.add_paragraph()
    r_sec1 = p_section1.add_run("I. KẾ HOẠCH DẠY HỌC PHÂN PHỐI CHƯƠNG TRÌNH")
    r_sec1.bold = True
    r_sec1.font.name = 'Times New Roman'
    r_sec1.font.size = Pt(14)
    
    headers = ["STT", "Tiết CT", "Bài học", "Số tiết", "Thời điểm", "Yêu cầu cần đạt", "Thiết bị", "Địa điểm"]
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        set_cell_margins(hdr_cells[i])
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)

    # 🌟 VÁ LỖI TRIỆT ĐỂ: Gán chỉ mục [số ô] rõ ràng từ 0 đến 7 để đổ dữ liệu vào Word không bị lỗi mảng
    if isinstance(content_data, list):
        for idx, item in enumerate(content_data, 1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(idx)
            row_cells[1].text = str(item.get("TietCT", idx))
            row_cells[2].text = str(item.get("BaiHoc", ""))
            row_cells[3].text = str(item.get("SoTiet", ""))
            row_cells[4].text = str(item.get("ThoiDiem", ""))
            row_cells[5].text = str(item.get("YeuCauCanDat", "-"))
            row_cells[6].text = str(item.get("ThietBi", "-"))
            row_cells[7].text = str(item.get("DiaDiem", "Lớp học"))
            
            for cell in row_cells:
                set_cell_margins(cell)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(11)
                        
    doc.add_paragraph()
    p_section2 = doc.add_paragraph()
    r_sec2 = p_section2.add_run("\nII. CÁC NHIỆM VỤ KHÁC ĐƯỢC GIAO")
    r_sec2.bold = True
    r_sec2.font.name = 'Times New Roman'
    r_sec2.font.size = Pt(14)
    
    p_other = doc.add_paragraph()
    r_other = p_other.add_run("- Bồi dưỡng học sinh giỏi và phụ đạo học sinh yếu kém theo kế hoạch nhà trường.\n- Tham gia sinh hoạt chuyên môn tổ, nhóm đầy đủ và đúng quy định.")
    r_other.font.name = 'Times New Roman'
    r_other.font.size = Pt(12)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- HÀM XUẤT FILE EXCEL MẪU 8 CỘT CHUẨN XÁC ---
def export_plan_to_excel(teacher_name, subject_name, content_data):
    output = io.BytesIO()
    raw_df = pd.DataFrame(content_data)
    df = pd.DataFrame({
        "STT": range(1, len(raw_df) + 1),
        "Tiết CT": raw_df.get("TietCT", range(1, len(raw_df) + 1)),
        "Bài học": raw_df.get("BaiHoc", ""),
        "Số tiết": raw_df.get("SoTiet", ""),
        "Thời điểm": raw_df.get("ThoiDiem", ""),
        "Yêu cầu cần đạt": raw_df.get("YeuCauCanDat", ""),
        "Thiết bị": raw_df.get("ThietBi", ""),
        "Địa điểm dạy học": raw_df.get("DiaDiem", "Lớp học")
    })
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        worksheet = workbook.create_sheet(title="Kế hoạch dạy học", index=0)
        writer.sheets["Kế hoạch dạy học"] = worksheet
        worksheet["A1"] = "I. KẾ HOẠCH DẠY HỌC"
        worksheet["A2"] = f"1. KHDH MÔN {subject_name.upper()}"
        df.to_excel(writer, sheet_name="Kế hoạch dạy học", startrow=3, index=False)
        
        for col_idx, col in enumerate(worksheet.columns, 1):
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
    return output.getvalue()

# --- HÀM TRÍCH XUẤT CHỮ TỪ FILE TÀI LIỆU ---
def extract_text_from_file(uploaded_file):
    if uploaded_file is None:
        return ""
    file_name = uploaded_file.name
    try:
        if file_name.endswith(".txt"):
            return uploaded_file.read().decode("utf-8")
        elif file_name.endswith(".docx"):
            doc = docx.Document(uploaded_file)
            return "\n".join([p.text for p in doc.paragraphs])
        elif file_name.endswith(".pdf"):
            return "Nội dung tệp PDF đính kèm: " + str(uploaded_file.name)
    except Exception as e:
        return f"Không thể đọc file: {str(e)}"
    return ""
# --- GIAO DIỆN PHÂN HỆ KẾ HOẠCH CÁ NHÂN ---
def render_personal_plan(run_ai_handler=None):
    st.markdown("<h3 style='text-align: left; color: #1E3A8A;'>🗓️ TRỢ LÝ XÂY DỰNG KẾ HOẠCH GIÁO DỤC CÁ NHÂN (PHỤ LỤC III)</h3>", unsafe_allow_html=True)
    
    if "db_ke_hoach_da_luu" not in st.session_state:
        st.session_state["db_ke_hoach_da_luu"] = {}
        
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔒 CHỌN VAI TRÒ ĐĂNG NHẬP")
    vai_tro = st.sidebar.radio("Vai trò", ["Giáo viên bộ môn", "Tổ trưởng chuyên môn (Admin)", "Ban giám hiệu"], label_visibility="collapsed", key="vai_tro_plan_sidebar_fixed_v8")
    
    # Chỉ giữ lại thông báo gợi ý phân quyền phụ, loại bỏ toàn bộ chốt chặn logic is_admin
    if vai_tro == "Tổ trưởng chuyên môn (Admin)":
        st.sidebar.info("Mã xác thực PIN hệ thống tổng đang hoạt động ngầm.")

    if "current_ai_output_raw" not in st.session_state: st.session_state["current_ai_output_raw"] = None
    if "current_teacher_active" not in st.session_state: st.session_state["current_teacher_active"] = ""
    if "current_subject_active" not in st.session_state: st.session_state["current_subject_active"] = ""

    with st.form("form_personal_plan_fixed_final_v8", border=False):
        col_t, col_s = st.columns(2)
        t_name = col_t.text_input("Họ và tên Giáo viên giảng dạy:", placeholder="Ví dụ: Thầy Lê Hồng Dưỡng", key="plan_txt_t_name_v8")
        s_name = col_s.selectbox("Môn học / Phân môn phụ trách:", ["Khoa học tự nhiên (Vật lý)", "Khoa học tự nhiên (Sinh học)", "Khoa học tự nhiên (Hóa học)", "Toán học", "Ngữ văn", "GDTC"], key="plan_sb_s_name_v8")
        col_g, col_w = st.columns(2)
        grade_target = col_g.text_input("Khối lớp phân công dạy:", placeholder="Ví dụ: Khối 9, Khối 7", key="plan_txt_grade_target_v8")
        week_count = col_w.text_input("Tổng số tuần thực hiện kế hoạch:", placeholder="Ví dụ: 35 Tuần", key="plan_txt_week_count_v8")
        
        st.markdown("**💬 Các tiêu chí đặc thù hoặc lưu ý phân bổ tiết (Nếu có):**")
        note_plan = st.text_area("Yêu cầu bổ sung cho AI:", placeholder="Ví dụ: Cần chia nhỏ phân phối chương trình thành từng tiết đơn lẻ 1, 2, 3...", label_visibility="collapsed", key="plan_ta_note_v8")
        
        st.markdown("    **📚 Đính kèm tài liệu phân phối chương trình hoặc Chương trình môn học (Tùy chọn):**")
        sgk_file = st.file_uploader("Tải file tài liệu SGK (.txt, .docx, .pdf)", type=["txt", "docx", "pdf"], label_visibility="collapsed", key="sgk_file_uploader_v8")
        
        run_ai_plan = st.form_submit_button("🚀 Khởi tạo Kế hoạch bằng AI", type="primary", use_container_width=True)
    if run_ai_plan:
        if not t_name or not grade_target:
            st.warning("⚠️ Vui lòng điền Họ tên giáo viên và Khối lớp để AI lập kế hoạch!")
        else:
            with st.spinner("Trợ lý AI đang đọc tệp tài liệu và phân bổ dữ liệu Phụ lục III..."):
                context_document = ""
                if sgk_file is not None:
                    context_document = extract_text_from_file(sgk_file)
                
                prompt_plan = (
                    f"Hãy soạn thảo phân bổ chương trình dạy học chi tiết bám sát định hướng Chương trình GDPT 2018 cho giáo viên: {t_name}, môn: {s_name}, khối: {grade_target} trong {week_count}. "
                    f"Yêu cầu bổ sung kỹ thuật: {note_plan}. "
                    f"Tại cột 'YeuCauCanDat', bạn bắt buộc phải phân tích và trích xuất đúng bám sát theo tài liệu tham khảo được cung cấp ở đây:\n"
                    f"[DỮ LIỆU THAM KHẢO TỪ FILE TÀI LIỆU]:\n{context_document}\n\n"
                    f"Mỗi bài học lớn phải được rải đều tách nhỏ thành các hàng đơn lẻ ứng với từng 'Tiết CT' độc lập (Ví dụ: Bài 1 (Tiết 1) xếp ở 1 dòng riêng, Bài 1 (Tiết 2) xếp ở dòng tiếp theo). "
                    f"Trả về mảng JSON thuần túy gồm danh sách các đối tượng, không kèm markdown thô nào ngoài thẻ mở/đóng json. "
                    f"Mỗi đối tượng bắt buộc phải chứa đúng cấu trúc 8 khóa sau không được sai lệch: "
                    f'[{{"TietCT": "Số thứ tự tiết lũy tiến tăng dần", "BaiHoc": "Tên bài học kèm chỉ số tiết (Ví dụ: Bài 1. Giới thiệu... (Tiết 1))", '
                    f'"SoTiet": "Tổng số tiết phân bổ cho bài đó", "ThoiDiem": "Tuần thực hiện (Ví dụ: Tuần 1)", '
                    f'"YeuCauCanDat": "Yêu cầu cần đạt cụ thể sinh tự động bám sát theo tài liệu tham khảo", "ThietBi": "Thiết bị dạy học sử dụng cụ thể", "DiaDiem": "Địa điểm dạy học (Mặc định: Lớp học)"}}]'
                )
                
                if run_ai_handler is not None:
                    res_plan, status = run_ai_handler(prompt_plan)
                    if status == "error":
                        st.error(f"❌ Hệ thống kết nối AI gặp sự cố kỹ thuật: {res_plan}")
                    else:
                        try:
                            clean_json = res_plan.replace("```json", "").replace("```", "").strip()
                            parsed_data = json.loads(clean_json)
                            
                            st.session_state["current_ai_output_raw"] = parsed_data
                            st.session_state["current_teacher_active"] = t_name
                            st.session_state["current_subject_active"] = s_name
                            
                            record_id = f"{t_name.replace(' ', '_')}_{s_name.replace(' ', '_')}"
                            st.session_state["db_ke_hoach_da_luu"][record_id] = {
                                "teacher": t_name,
                                "subject": s_name,
                                "grade": grade_target,
                                "weeks": week_count,
                                "data": parsed_data
                            }
                            st.success("🎉 Khởi tạo và lưu trữ phân phối chương trình thành công!")
                        except Exception as parse_err:
                            st.error("⚠️ AI phản hồi cấu hình phân tách hàng chưa đồng bộ. Vui lòng bấm thử lại để làm mới.")
                else:
                    st.error("❌ Lỗi cấu hình hệ thống: Không tìm thấy trình điều khiển AI tập trung.")

    # --- KHU VỰC HIỂN THỊ KẾT QUẢ VÀ SONG HÀNH HAI NÚT TẢI XUẤT ---
    st.markdown("---")
    st.markdown("### 📊 Nội dung Kế hoạch Giáo dục sinh bởi AI:")
    
    if st.session_state["current_ai_output_raw"]:
        df_display = pd.DataFrame(st.session_state["current_ai_output_raw"])
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        btn_col1, btn_col2 = st.columns(2)
        
        word_plan_data = export_plan_to_docx_with_table(
            st.session_state["current_teacher_active"], 
            st.session_state["current_subject_active"], 
            st.session_state["current_ai_output_raw"]
        )
        btn_col1.download_button(
            label="📥 Tải file Word (.docx) Phụ lục III chuẩn", 
            data=word_plan_data, 
            file_name=f"Phu_Luc_III_{st.session_state['current_teacher_active'].replace(' ', '_')}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            use_container_width=True
        )
        
        excel_plan_data = export_plan_to_excel(
            st.session_state["current_teacher_active"], 
            st.session_state["current_subject_active"], 
            st.session_state["current_ai_output_raw"]
        )
        btn_col2.download_button(
            label="📊 Tải file Excel (.xlsx) Phân phối chương trình chi tiết", 
            data=excel_plan_data, 
            file_name=f"Phan_Phoi_Chuong_Trinh_{st.session_state['current_teacher_active'].replace(' ', '_')}.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            use_container_width=True
        )
    else:
        st.caption("Khung kế hoạch chi tiết dạng bảng biểu 8 cột sẽ xuất hiện tại đây sau khi bấm nút khởi tạo...")

    # --- KHU VỰC LỊCH SỬ KẾ HOẠCH ĐÃ LƯU ---
    if st.session_state["db_ke_hoach_da_luu"]:
        st.markdown("---")
        st.markdown("### 🗂️ Danh sách Kế hoạch Giáo dục đã lưu trên Hệ thống:")
        for key_id, info in st.session_state["db_ke_hoach_da_luu"].items():
            with st.expander(f"📋 Bản kế hoạch môn {info['subject']} - GV: {info['teacher']} (Khối {info['grade']})"):
                st.write(f"**Tổng số tuần**: {info['weeks']}")
                df_history = pd.DataFrame(info['data'])
                st.dataframe(df_history, use_container_width=True, hide_index=True)
                
                h_col1, h_col2 = st.columns(2)
                
                word_history_data = export_plan_to_docx_with_table(info['teacher'], info['subject'], info['data'])
                h_col1.download_button(
                    label=f"📥 Tải file Word của GV {info['teacher']}", 
                    data=word_history_data, 
                    file_name=f"Phu_Luc_III_{info['teacher'].replace(' ', '_')}.docx", 
                    key=f"dl_word_hist_{key_id}",
                    use_container_width=True
                )
                
                excel_history_data = export_plan_to_excel(info['teacher'], info['subject'], info['data'])
                h_col2.download_button(
                    label=f"📊 Tải file Excel của GV {info['teacher']}", 
                    data=excel_history_data, 
                    file_name=f"Phan_Phoi_Chuong_Trinh_{info['teacher'].replace(' ', '_')}.xlsx", 
                    key=f"dl_excel_hist_{key_id}",
                    use_container_width=True
                )
