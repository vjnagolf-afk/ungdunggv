import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re

def set_paragraph_spacing(paragraph, before_pt=3.0, after_pt=4.5):
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(before_pt * 20)))
    spacing.set(qn('w:after'), str(int(after_pt * 20)))
    spacing.set(qn('w:line'), '240')
    spacing.set(qn('w:lineRule'), 'auto')
    p_pr.append(spacing)

def export_rubric_to_docx(title_text, markdown_content):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(1.18)
        section.right_margin = Inches(0.59)

    MAU_DO = RGBColor(255, 0, 0)
    MAU_XANH_DUONG = RGBColor(0, 51, 153)
    MAU_DEN = RGBColor(0, 0, 0)

    # Tiêu đề Rubric căn giữa màu đỏ
    p_title = doc.add_paragraph()
    set_paragraph_spacing(p_title, 6.0, 6.0)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(title_text.upper())
    r_title.bold = True
    r_title.font.name = 'Times New Roman'
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = MAU_DO

    lines = markdown_content.split('\n')
    in_table = False
    table_data = []

    for line in lines:
        clean_line = line.strip().replace('**', '').replace('###', '').replace('##', '').replace('#', '')
        
        if line.strip().startswith('|') and line.strip().endswith('|'):
            if '---|' in line or ':---|' in line: continue
            in_table = True
            cells = [c.strip().replace('**', '') for c in line.split('|')[1:-1]]
            table_data.append(cells)
            continue
        else:
            if in_table and table_data:
                num_rows = len(table_data)
                num_cols = len(table_data) if num_rows > 0 else 0
                if num_cols > 0:
                    word_table = doc.add_table(rows=num_rows, cols=num_cols)
                    word_table.style = 'Table Grid'
                    word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    for r_idx, row in enumerate(table_data):
                        for c_idx, val in enumerate(row):
                            if c_idx < num_cols:
                                cell = word_table.cell(r_idx, c_idx)
                                cell.text = val
                                for para in cell.paragraphs:
                                    set_paragraph_spacing(para, 2.0, 3.0)
                                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                                    for r in para.runs:
                                        r.font.name = 'Times New Roman'
                                        r.font.size = Pt(14)
                                        r.font.color.rgb = MAU_DEN
                in_table = False
                table_data = []

        if not clean_line: continue

        if re.match(r'^(I|II|III|IV|V)\.', clean_line) or re.match(r'^\d+\.', clean_line):
            p = doc.add_paragraph()
            set_paragraph_spacing(p, 4.0, 4.5)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(clean_line)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.color.rgb = MAU_XANH_DUONG
        else:
            p = doc.add_paragraph()
            set_paragraph_spacing(p)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.add_run(clean_line)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()
