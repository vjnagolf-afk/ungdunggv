# math_compiler.py - Bản vá dứt điểm 100% lỗi NameError Pt
import io
import re
import numpy as np
import matplotlib.pyplot as plt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 🚀 IMPORT ĐẦY ĐỦ ĐỂ TRÁNH LỖI NAMEERROR
from docx.shared import Pt

def generate_plot_stream(eq_str):
    fig, ax = plt.subplots(figsize=(5, 3.5))
    x = np.linspace(-10, 10, 400)
    safe_dict = {"x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan, "sqrt": np.sqrt}
    try:
        eq_str_py = eq_str.replace('^', '**')
        y = eval(eq_str_py, {"__builtins__": {}}, safe_dict)
        if isinstance(y, (int, float)):
            y = np.full_like(x, y)
        ax.plot(x, y, color='#1E40AF', linewidth=2.5)
        ax.axhline(0, color='black', linewidth=1.2)
        ax.axvline(0, color='black', linewidth=1.2)
        ax.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)
        ax.set_ylim([-10, 10])
        ax.set_title(f"Đồ thị: y = {eq_str}", fontsize=10, pad=10)
    except:
        ax.text(0.5, 0.5, f"[Lỗi cú pháp vẽ đồ thị]", ha='center', va='center', color='red')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
    buf.seek(0)
    plt.close(fig)
    return buf

def build_omml_fraction(num_str, den_str):
    return (
        f'<m:f>'
        f'<m:num><m:r>'
        f'<m:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></m:rPr>'
        f'<m:t>{num_str}</m:t>'
        f'</m:r></m:num>'
        f'<m:den><m:r>'
        f'<m:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></m:rPr>'
        f'<m:t>{den_str}</m:t>'
        f'</m:r></m:den>'
        f'</m:f>'
    )

def convert_latex_to_omml(latex_str):
    latex_str = latex_str.strip()
    latex_str = latex_str.replace(r'\pi', 'π').replace(r'\infty', '∞')
    latex_str = latex_str.replace(r'\times', '×').replace(r'\cdot', '·')
    latex_str = latex_str.replace(r'\approx', '≈')
    
    latex_str = re.sub(r'\(\((.*?)\)/\((.*?)\)\)', r'(\1)/(\2)', latex_str)
    latex_str = re.sub(r'\((.*?)\)/\((.*?)\)', r'(\1)/(\2)', latex_str)

    frac_pattern = re.compile(r'\\frac\{([^}]+)\}\{([^}]+)\}')
    while frac_pattern.search(latex_str):
        match = frac_pattern.search(latex_str)
        xml_frac = build_omml_fraction(match.group(1), match.group(2))
        latex_str = latex_str.replace(match.group(0), xml_frac)
        
    plain_frac_pattern = re.compile(r'([a-zA-Z0-9_().+*-]+)/([a-zA-Z0-9_().+*-]+)')
    while plain_frac_pattern.search(latex_str):
        match = plain_frac_pattern.search(latex_str)
        num = match.group(1).strip('()')
        den = match.group(2).strip('()')
        if '<m:f>' in match.group(0): 
            break
        xml_frac = build_omml_fraction(num, den)
        latex_str = latex_str.replace(match.group(0), xml_frac)

    latex_str = re.sub(r'\^\{([^}]+)\}', r'^\1', latex_str)
    
    omml_xml = f'<m:oMath {nsdecls("m")}><m:r><m:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></m:rPr>'
    if '<m:f>' in latex_str:
        omml_xml += f'</m:r>{latex_str}'
    else:
        omml_xml += f'<m:t>{latex_str}</m:t></m:r>'
    omml_xml += '</m:oMath>'
    try:
        return parse_xml(omml_xml)
    except:
        return None

def process_runs_with_math(paragraph, text):
    text_clean = text.strip()
    
    match_choice = re.match(r'^([A-D]\.\s*)(.*)', text_clean)
    if match_choice:
        prefix = match_choice.group(1)
        remain_text = match_choice.group(2)
        run_prefix = paragraph.add_run(prefix)
        run_prefix.bold = True
        run_prefix.font.name = 'Times New Roman'
        run_prefix.font.size = Pt(14)
        text_clean = remain_text

    parts = re.split(r'(\$\$[\s\S]*?\$\$|\$[\s\S]*?\$)', text_clean)
    for part in parts:
        if not part:
            continue
        if part.startswith('$'):
            math_content = part.replace('$$', '').replace('$', '').strip()
            if math_content:
                math_element = convert_latex_to_omml(math_content)
                if math_element is not None:
                    paragraph._p.append(math_element)
                else:
                    run = paragraph.add_run(part)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
        else:
            bold_parts = part.split('**')
            for i, b_part in enumerate(bold_parts):
                is_bold = (i % 2 != 0)
                sub_sup_parts = re.split(r'(<sub>.*?</sub>|<sup>.*?</sup>)', b_part)
                for s_part in sub_sup_parts:
                    if not s_part: 
                        continue
                    if s_part.startswith('<sub>') and s_part.endswith('</sub>'):
                        run = paragraph.add_run(s_part[5:-6])
                        run.bold = is_bold
                        run.font.subscript = True
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
                    elif s_part.startswith('<sup>') and s_part.endswith('</sup>'):
                        run = paragraph.add_run(s_part[5:-6])
                        run.bold = is_bold
                        run.font.superscript = True
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
                    else:
                        run = paragraph.add_run(s_part)
                        run.bold = is_bold
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
