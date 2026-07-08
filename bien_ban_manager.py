import streamlit as st
import sqlite3
import pandas as pd
import io

DB_PATH = "teacher_assistant.db"

def setup_minutes_database():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS org_minutes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, meeting_date TEXT, session_number TEXT UNIQUE,
        present_members TEXT, absent_members TEXT, content_text TEXT, resolution TEXT
    )
    """)
    conn.commit()
    conn.close()

def extract_text_from_minutes_upload(uploaded_file):
    import docx
    from pypdf import PdfReader
    text_content = ""
    try:
        if uploaded_file.name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            for p in doc.paragraphs:
                if p.text: text_content += p.text + "\n"
        elif uploaded_file.name.endswith('.pdf'):
            reader = PdfReader(uploaded_file)
            for page in reader.pages:
                text_content += (page.extract_text() or "") + "\n"
        elif uploaded_file.name.endswith('.txt'):
            text_content = uploaded_file.read().decode("utf-8")
    except Exception as e:
        st.error(f"Lỗi bóc tách file: {e}")
    return text_content.strip()

def export_minutes_to_docx(meeting_date, session_number, present_members, absent_members, content_text, resolution):
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n").bold = True
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"\nBIÊN BẢN SINH HOẠT TỔ CHUYÊN MÔN\n(Số: {session_number})\n")
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(255, 0, 0)
    items = [f"- Ngày họp: {meeting_date}", f"- Thành phần: {present_members}", f"- Vắng mặt: {absent_members}", "\n--- DIỄN BIẾN CUỘC HỌP ---", content_text, "\n--- NGHỊ QUYẾT CUỘC HỌP ---", resolution]
    for item in items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r = p.add_run(item)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
def render_meeting_minutes():
    setup_minutes_database()
    st.markdown("<h3 style='text-align: left; color: #1E3A8A;'>📝 TRÌNH LẬP BIÊN BẢN HỌP TỔ CHUYÊN MÔN NÂNG CAO</h3>", unsafe_allow_html=True)
    tab_lap, tab_kho = st.tabs(["✍️ Lập biên bản mới", "🗄️ Kho lưu trữ biên bản họp"])
    
    if "minutes_content_draft" not in st.session_state: st.session_state["minutes_content_draft"] = ""
    if "minutes_resolution_draft" not in st.session_state: st.session_state["minutes_resolution_draft"] = ""

    with tab_lap:
        st.markdown("**📁 Nạp file tài liệu thảo luận (Chèn tự động vào biên bản):**")
        file_ke_hoach = st.file_uploader("Kéo thả file kế hoạch (.docx, .pdf, .txt):", type=["docx", "pdf", "txt"], key="file_min_up_isolated_v5")
        
        if file_ke_hoach and st.button("🧠 AI: Nghiên cứu file & Tóm tắt chèn vào diễn biến cuộc họp", type="secondary", use_container_width=True, key="btn_ai_extract_minutes_fixed_v5"):
            with st.spinner("AI đang xử lý..."):
                raw_text = extract_text_from_minutes_upload(file_ke_hoach)
                if raw_text:
                    try:
                        from app import run_ai_prompt_safe
                        # Gọi cấu trúc prompt rút gọn hành vi định lượng sạch sẽ
                        res_text, _ = run_ai_prompt_safe(f"Hãy đọc tệp văn bản này và tóm tắt thành một nội dung biên bản cuộc họp tổ chuyên môn cực kỳ chi tiết, dùng dấu gạch ngang '-', không dùng dấu **: {raw_text}", st.secrets.get("GEMINI_API_KEY", ""))
                        st.session_state["minutes_content_draft"] = res_text
                        st.success("✅ Đã chèn dữ liệu tự động thành công vào khung bên dưới!")
                        st.rerun()
                    except Exception as e: st.error(f"Lỗi: {e}")

        # Khung biểu mẫu Form bảo vệ đứng yên nháy chuột
        with st.form("form_create_minutes_secure_v5_clean", border=False):
            c1, c2 = st.columns(2)
            m_date, m_num = c1.date_input("Ngày họp:"), c2.text_input("Số biên bản (Ký hiệu):", key="txt_m_num_fixed_v5")
            m_present = st.text_area("Thành phần tham dự:", key="txt_m_present_fixed_v5")
            m_absent = st.text_input("Vắng mặt:", key="txt_m_absent_fixed_v5")
            m_content = st.text_area("Nội dung cuộc họp:", value=st.session_state["minutes_content_draft"], height=180, key="txt_m_content_fixed_v5")
            
            run_summary_ai = st.form_submit_button("🧠 AI: Tóm tắt Quyết nghị/Nghị quyết dựa trên diễn biến cuộc họp ở trên", type="secondary")
            if run_summary_ai:
                if m_content:
                    with st.spinner("AI đang đúc kết..."):
                        try:
                            from app import run_ai_prompt_safe
                            res_ai, _ = run_ai_prompt_safe(f"Tóm tắt ngắn gọn nghị quyết cuộc họp từ diễn biến này, dùng dấu '-', không dùng **: {m_content}", st.secrets.get("GEMINI_API_KEY", ""))
                            st.session_state["minutes_resolution_draft"] = res_ai
                            st.rerun()
                        except: pass
                else: st.warning("⚠️ Vui lòng điền nội dung cuộc họp trước!")

            m_resolution = st.text_area("Nghị quyết / Kết luận chung:", value=st.session_state["minutes_resolution_draft"], height=100, key="txt_m_resolution_fixed_v5")
            st.markdown("<br>", unsafe_allow_html=True)
            submit_save_minutes = st.form_submit_button("💾 KHÓA & LƯU BIÊN BẢN VÀO HỆ THỐNG", type="primary", use_container_width=True)
            
        if submit_save_minutes:
            if m_num and m_content:
                conn = sqlite3.connect(DB_PATH)
                conn.execute("INSERT OR REPLACE INTO org_minutes (meeting_date, session_number, present_members, absent_members, content_text, resolution) VALUES (?, ?, ?, ?, ?, ?)", (str(m_date), m_num.strip(), m_present.strip(), m_absent.strip(), m_content.strip(), m_resolution.strip()))
                conn.commit(); conn.close()
                st.success(f"🎉 Biên bản số **{m_num}** đã được lưu trữ vĩnh viễn thành công!")
                st.session_state["minutes_content_draft"] = ""
                st.session_state["minutes_resolution_draft"] = ""
                st.rerun()
            else: st.warning("⚠️ Vui lòng điền đầy đủ Số biên bản và Diễn biến cuộc họp!")

    with tab_kho:
        st.markdown("#### 🗄️ Thư viện lưu trữ biên bản sinh hoạt tổ")
        conn = sqlite3.connect(DB_PATH)
        all_minutes = conn.execute("SELECT id, meeting_date, session_number, present_members, absent_members, content_text, resolution FROM org_minutes ORDER BY meeting_date DESC").fetchall()
        conn.close()
        for idx, row in enumerate(all_minutes):
            m_id, m_dt, m_snum, m_pres, m_abs, m_ct, m_resol = row
            col_exp, col_del = st.columns([0.88, 0.12])
            with col_exp:
                with st.expander(f"📝 {idx+1}. Biên bản số: {m_snum} - Ngày họp: {m_dt}"):
                    st.write(f"**Thành phần:** {m_pres} | **Vắng:** {m_abs}")
                    st.caption(f"**Diễn biến:**\n{m_ct}")
                    st.write(f"**Nghị quyết:**\n{m_resol}")
                    w_data = export_minutes_to_docx(m_dt, m_snum, m_pres, m_abs, m_ct, m_resol)
                    st.download_button(label="📥 Tải file Word (.docx)", data=w_data, file_name=f"Bien_Ban_So_{m_snum}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_w_minutes_btn_{m_id}", use_container_width=True)
            with col_del:
                st.write("")
                if st.button("🗑️ Xóa", key=f"del_m_minutes_btn_{m_id}", use_container_width=True):
                    conn = sqlite3.connect(DB_PATH); conn.execute("DELETE FROM org_minutes WHERE id=?", (m_id,)); conn.commit(); conn.close(); st.rerun()
