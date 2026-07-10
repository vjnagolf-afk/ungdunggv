import streamlit as st
import docx  
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
from pypdf import PdfReader
import gspread
from datetime import datetime

from math_compiler import process_runs_with_math, generate_plot_stream

SHEET_ID = '1C6642jk_oQ0g9UC2By2qsNxxfQVR0MrZYj52tRdWDlY' 

def get_khbd_sheet():
    try:
        creds_dict = None
        for key in ["gspread_credentials", "GSPREAD_CREDENTIALS", "google_sheet_creds"]:
            if key in st.secrets: creds_dict = st.secrets[key]; break
        if not creds_dict: return None
        gc = gspread.service_account_from_dict(creds_dict)
        return gc.open_by_key(SHEET_ID).worksheet("KHBD")
    except: return None

def save_khbd_to_sheet(ten_bai, lop, bo_sach, thoi_luong, content):
    sheet = get_khbd_sheet()
    if sheet:
        sheet.append_row([ten_bai, lop, bo_sach, thoi_luong, content, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        return True
    return False

def get_all_khbd_from_sheet():
    sheet = get_khbd_sheet()
    return sheet.get_all_records() if sheet else []

def delete_khbd_from_sheet(row_index):
    sheet = get_khbd_sheet()
    if sheet:
        sheet.delete_rows(row_index + 2)
        return True
    return False

def set_paragraph_spacing(paragraph, before_pt=3.0, after_pt=4.5):
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(before_pt * 20)))
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    p_pr.append(spacing)

def export_khbd_to_docx(markdown_content, images_list):
    markdown_content = re.sub(r'(?m)^#+\s*', '', markdown_content)
    markdown_content = markdown_content.replace("<br>", "\n").replace("<br/>", "\n")
    
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79); section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18); section.right_margin = Inches(0.59)

    MAU_DO, MAU_XANH_DUONG, MAU_DEN = RGBColor(255, 0, 0), RGBColor(0, 51, 153), RGBColor(0, 0, 0)
    lines = markdown_content.split('\n')
    table_rows = []

    for line in lines:
        cl = line.strip().replace('**', '').replace('*', '') 
        if not cl: continue

        if '|' in cl and not cl.startswith(('-', 'I.', 'II.', '1.', '2.')):
            row_data = [cell.strip() for cell in cl.split('|') if cell.strip()]
            if row_data: table_rows.append(row_data)
            continue
        elif table_rows:
            max_cols = max(len(row) for row in table_rows)
            tbl = doc.add_table(rows=len(table_rows), cols=max_cols)
            tbl.style = 'Table Grid'
            for i, row in enumerate(table_rows):
                for j in range(max_cols):
                    val = row[j] if j < len(row) else ""
                    cell = tbl.cell(i, j)
                    for p_cell in cell.paragraphs: p_cell._element.getparent().remove(p_cell._element)
                    p_cell = cell.add_paragraph()
                    process_runs_with_math(p_cell, val)
                    for r in p_cell.runs: r.font.name = 'Times New Roman'; r.font.size = Pt(12)
            table_rows = []
            doc.add_paragraph()

        p = doc.add_paragraph()
        set_paragraph_spacing(p, 3.0, 4.5)
        p.paragraph_format.left_indent = Inches(0); p.paragraph_format.right_indent = Inches(0)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        upper_cl = cl.upper()
        if upper_cl.startswith(("KẾ HOẠCH BÀI DẠY", "BÀI")):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(upper_cl)
            run.bold = True; run.font.name = 'Times New Roman'; run.font.size = Pt(14); run.font.color.rgb = MAU_DO
        elif any(upper_cl.startswith(x) for x in ["MÔN HỌC", "LỚP", "THỜI LƯỢNG", "TIẾT"]):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(upper_cl)
            run.bold = True; run.font.name = 'Times New Roman'; run.font.size = Pt(14); run.font.color.rgb = MAU_XANH_DUONG
        elif re.match(r'^(I|II|III|IV|V|VI|VII)\.', cl):
            process_runs_with_math(p, upper_cl)
            for r in p.runs: r.bold = True; r.font.name = 'Times New Roman'; r.font.size = Pt(14); r.font.color.rgb = MAU_XANH_DUONG
        elif re.match(r'^\d+\.\s+(Kiến thức|Năng lực|Phẩm chất)', cl, re.IGNORECASE):
            process_runs_with_math(p, cl)
            for r in p.runs: r.bold = True; r.font.name = 'Times New Roman'; r.font.size = Pt(14); r.font.color.rgb = MAU_DO
        else:
            process_runs_with_math(p, cl)
            for r in p.runs: r.font.name = 'Times New Roman'; r.font.size = Pt(14); r.font.color.rgb = MAU_DEN

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def render_khbd_section(run_ai_prompt_safe_func):
    st.markdown("<h3 style='text-align: center; color: blue;'>🧠 TRỢ LÝ THIẾT KẾ KHBD AI</h3>", unsafe_allow_html=True)
    tab_thiet_ke, tab_thu_vien = st.tabs(["📝 THIẾT KẾ", "🗄️ THƯ VIỆN"])
    
    with tab_thiet_ke:
        ten_bai = st.text_input("Tên bài học:")
        if st.button("⚡ Thiết kế"):
            st.session_state["ket_qua_khbd"] = run_ai_prompt_safe_func(f"Soạn bài: {ten_bai}")[0]
        
        if st.session_state.get("ket_qua_khbd"):
            st.markdown(st.session_state["ket_qua_khbd"])
            if st.button("💾 Lưu tạm"):
                save_khbd_to_sheet(ten_bai, "Lớp 7", "Kết nối", "2 tiết", st.session_state["ket_qua_khbd"])
                st.success("Đã lưu!")
                
    with tab_thu_vien:
        for idx, bai in enumerate(get_all_khbd_from_sheet()):
            with st.expander(f"{bai['Tên bài']} ({bai['Thời gian']})"):
                if st.button("📥 Gọi", key=f"load_{idx}"):
                    st.session_state["ket_qua_khbd"] = bai['Nội dung chi tiết']; st.rerun()
                if st.button("🗑️ Xóa", key=f"del_{idx}"):
                    delete_khbd_from_sheet(idx); st.rerun()
