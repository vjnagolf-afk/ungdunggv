# math_compiler.py - ĐOẠN 1: KHAI BÁO THƯ VIỆN & ENGINE ĐỒ THỊ
import io
import re
import numpy as np
import matplotlib.pyplot as plt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt

def generate_plot_stream(eq_str):
    """Vẽ đồ thị tự động từ biểu thức AI cung cấp, hỗ trợ cả dạng scatter (sơ đồ điểm)"""
    fig, ax = plt.subplots(figsize=(5.5, 3.8))
    x = np.linspace(-10, 10, 500)
    
    # Thiết lập bộ từ điển toán học an toàn
    safe_dict = {
        "x": x, "np": np, "pi": np.pi, "e": np.e,
        "sin": np.sin, "cos": np.cos, "tan": np.tan, 
        "sqrt": np.sqrt, "log": np.log10, "ln": np.log, "exp": np.exp
    }
    
    try:
        clean_eq = eq_str.strip()
        # Loại bỏ tiền tố "y=" nếu AI sinh thừa
        if clean_eq.lower().startswith("y="):
            clean_eq = clean_eq[2:].strip()
            
        if "scatter" in clean_eq.lower():
            # Nếu là biểu đồ điểm scatter (Phục vụ thống kê số liệu)
            np.random.seed(42)
            xs = np.random.uniform(-5, 5, 30)
            ys = 2 * xs + 1 + np.random.normal(0, 1.5, 30)
            ax.scatter(xs, ys, color='#1E40AF', edgecolors='black', alpha=0.8, s=45, label="Số liệu thực tế")
            ax.plot(x, 2 * x + 1, color='#EF4444', linestyle='--', linewidth=1.5, label="Đường xu hướng")
            ax.legend(fontsize=8, loc="upper left")
            ax.set_title("Biểu đồ phân tán số liệu", fontsize=10, fontname="Times New Roman", pad=10)
        else:
            # Biên dịch các hàm số thông thường sang cú pháp Python vắt qua lũy thừa ^
            eq_py = clean_eq.replace('^', '**')
            y = eval(eq_py, {"__builtins__": {}}, safe_dict)
            if isinstance(y, (int, float)):
                y = np.full_like(x, y)
                
            ax.plot(x, y, color='#1E40AF', linewidth=2.5, label=f"y = {clean_eq}")
            ax.set_title(f"Đồ thị hàm số: y = {clean_eq}", fontsize=10, fontname="Times New Roman", pad=10)
            
        # Cấu hình hệ trục tọa độ đề-các trực quan cho giáo án, đề thi
        ax.axhline(0, color='black', linewidth=1.2)
        ax.axvline(0, color='black', linewidth=1.2)
        ax.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.6)
        ax.set_xlim([-7, 7])
        ax.set_ylim([-7, 7])
        
    except Exception:
        ax.text(0.5, 0.5, f"[Lỗi cú pháp đồ thị: {eq_str}]", ha='center', va='center', color='red', fontname="Times New Roman")
        
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
    buf.seek(0)
    plt.close(fig)
    return buf
# math_compiler.py - ĐOẠN 2: LÕI TRANSLATOR SANG WORD EQUATION (OMML ENGINE)
def clean_latex_spacing(latex_str):
    """Thuật toán tăng khả năng chịu lỗi: Dọn sạch khoảng trắng dị biệt quanh ngoặc nhọn {}"""
    latex_str = re.sub(r'\\dfrac', r'\\frac', latex_str)
    latex_str = re.sub(r'\\\s*([a-zA-Z]+)', r'\\\1', latex_str)
    
    pattern = r'\\(frac|sqrt|sum|int|vec)\s*\{\s*([^}]*?)\s*\}\s*\{\s*([^}]*?)\s*\}'
    while re.search(pattern, latex_str):
        latex_str = re.sub(pattern, r'\\\1{\2}{\3}', latex_str)
        
    pattern_single = r'\\(sqrt|vec|sum|int)\s*\{\s*([^}]*?)\s*\}'
    while re.search(pattern_single, latex_str):
        latex_str = re.sub(pattern_single, r'\\\1{\2}', latex_str)
        
    return latex_str.strip()

def xml_tag(tag_name, content, attributes=""):
    attr_str = f" {attributes}" if attributes else ""
    return f"<{tag_name}{attr_str}>{content}</{tag_name}>"

def omml_run(text, is_bold=False, is_italic=True):
    """Tạo khối phần tử chữ OMML Equation Editor ép cứng phôi Times New Roman 14pt"""
    style_tags = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
    if is_bold: style_tags += '<w:b/><w:bCs/>'
    if not is_italic: style_tags += '<m:nor/>'
    return f'<m:r><m:rPr>{style_tags}</m:rPr><m:t>{text}</m:t></m:r>'

def parse_latex_to_omml_xml(latex_str):
    """Phân rã đệ quy bóc tách mã lệnh toán Vật lý và Toán học sang Office Math XML"""
    latex_str = clean_latex_spacing(latex_str)
    
    greek_replacements = {
        r'\alpha': 'α', r'\beta': 'β', r'\theta': 'θ', r'\pi': 'π', r'\infty': '∞',
        r'\times': '×', r'\cdot': '·', r'\approx': '≈', r'\omega': 'ω', r'\Delta': 'Δ',
        r'\lambda': 'λ', r'\rho': 'ρ', r'\sigma': 'σ', r'\phi': 'φ', r'\gamma': 'γ'
    }
    for k, v in greek_replacements.items():
        latex_str = latex_str.replace(k, v)

    # 1. XỬ LÝ PHÂN SỐ CHỒNG TẦNG: \frac{tử}{mẫu}
    frac_match = re.search(r'\\frac\{([^}]+)\}\{([^}]+)\}', latex_str)
    if frac_match:
        tu = parse_latex_to_omml_xml(frac_match.group(1))
        mau = parse_latex_to_omml_xml(frac_match.group(2))
        f_xml = xml_tag("m:f", xml_tag("m:num", tu) + xml_tag("m:den", mau))
        return latex_str.replace(frac_match.group(0), f_xml)

    # 2. XỬ LÝ CĂN BẬC HAI: \sqrt{biểu_thức}
    sqrt_match = re.search(r'\\sqrt\{([^}]+)\}', latex_str)
    if sqrt_match:
        loi_can = parse_latex_to_omml_xml(sqrt_match.group(1))
        rad_xml = xml_tag("m:rad", xml_tag("m:radPr", "") + xml_tag("m:deg", "") + xml_tag("m:e", loi_can))
        return latex_str.replace(sqrt_match.group(0), rad_xml)

    # 3. XỬ LÝ VECTOR: \vec{A}
    vec_match = re.search(r'\\vec\{([^}]+)\}', latex_str)
    if vec_match:
        loi_vec = parse_latex_to_omml_xml(vec_match.group(1))
        acc_xml = xml_tag("m:acc", xml_tag("m:accPr", '<m:chr m:val="➔"/>') + xml_tag("m:e", loi_vec))
        return latex_str.replace(vec_match.group(0), acc_xml)

    # 4. XỬ LÝ CHỈ SỐ TRÊN (LŨY THỪA): chữ^{số}
    sup_match = re.search(r'([a-zA-Z0-9α-γΔπ-σ]+)\^\{([^}]+)\}', latex_str) or re.search(r'([a-zA-Z0-9α-γΔπ-σ]+)\^([a-zA-Z0-9])', latex_str)
    if sup_match:
        co_so = omml_run(sup_match.group(1))
        so_mu = parse_latex_to_omml_xml(sup_match.group(2))
        sup_xml = xml_tag("m:sSup", xml_tag("m:e", co_so) + xml_tag("m:sup", so_mu))
        return latex_str.replace(sup_match.group(0), sup_xml)

    # 5. XỬ LÝ CHỈ SỐ DƯỚI: chữ_{chi_so}
    sub_match = re.search(r'([a-zA-Z0-9α-γΔπ-σ]+)_\{([^}]+)\}', latex_str) or re.search(r'([a-zA-Z0-9α-γΔπ-σ]+)_([a-zA-Z0-9])', latex_str)
    if sub_match:
        co_so = omml_run(sub_match.group(1))
        chi_so = parse_latex_to_omml_xml(sub_match.group(2))
        sub_xml = xml_tag("m:sSub", xml_tag("m:e", co_so) + xml_tag("m:sub", chi_so))
        return latex_str.replace(sub_match.group(0), sub_xml)

    # 6. XỬ LÝ TÍCH PHÂN VÀ TỔNG (int, sum): \int_{a}^{b}
    nary_match = re.search(r'\\(int|sum)_\{([^}]+)\}\^\{([^}]+)\}', latex_str)
    if nary_match:
        type_char = "∫" if nary_match.group(1) == "int" else "∑"
        lim_down = parse_latex_to_omml_xml(nary_match.group(2))
        lim_up = parse_latex_to_omml_xml(nary_match.group(3))
        nary_xml = xml_tag("m:nary", xml_tag("m:naryPr", f'<m:chr m:val="{type_char}"/><m:limLoc m:val="undOvr"/>') + xml_tag("m:sub", lim_down) + xml_tag("m:sup", lim_up) + xml_tag("m:e", ""))
        return latex_str.replace(nary_match.group(0), nary_xml)

    if not latex_str.strip().startswith('<m:'):
        return omml_run(latex_str.strip(), is_italic=True)
    return latex_str

def convert_latex_to_omml(latex_str):
    if not latex_str.strip(): 
        return None
    body_xml = parse_latex_to_omml_xml(latex_str)
    full_xml = f'<m:oMath {nsdecls("m")}>{body_xml}</m:oMath>'
    try:
        return parse_xml(full_xml)
    except Exception:
        return None
# math_compiler.py - ĐOẠN 3: PARSER HÓA HỌC TỰ ĐỘNG & ĐIỀU PHỐI LIÊN MODULE
def convert_chemical_to_omml(chem_str):
    """Tự động nhận diện công thức Hóa học phẳng và chuyển đổi sang Equation chỉ số đứng"""
    chem_str = chem_str.strip()
    tokens = re.findall(r'([A-Z][a-z]?)|(\d+)\^?([-+2]*)|\^\{?([-+2a-zA-Z0-9]+)\}?|([+\-=➔])|(\s+)', chem_str)
    
    omml_body = ""
    for tok in tokens:
        if tok[0]:  # Chữ cái hóa học
            omml_body += omml_run(tok[0], is_italic=False)
        elif tok[1]:  # Chỉ số nguyên tử dưới
            val_down = tok[1]
            val_up = tok[2] if tok[2] else ""
            if val_up:
                e_block = xml_tag("m:e", omml_run("", is_italic=False))
                sub_block = xml_tag("m:sub", omml_run(val_down, is_italic=False))
                sup_block = xml_tag("m:sup", omml_run(val_up, is_italic=False))
                omml_body += xml_tag("m:sPre", xml_tag("m:sPrePr", "") + sub_block + sup_block + e_block)
            else:
                e_block = xml_tag("m:e", omml_run("", is_italic=False))
                sub_block = xml_tag("m:sub", omml_run(val_down, is_italic=False))
                omml_body += xml_tag("m:sSub", e_block + sub_block)
        elif tok[3]:  # Điện tích mũ trên đầu
            e_block = xml_tag("m:e", omml_run("", is_italic=False))
            sup_block = xml_tag("m:sup", omml_run(tok[3], is_italic=False))
            omml_body += xml_tag("m:sSup", e_block + sup_block)
        elif tok[4]:  # Dấu phản ứng
            omml_body += omml_run(f" {tok[4]} ", is_italic=False)
        elif tok[5]:  # Khoảng trắng
            omml_body += omml_run(" ", is_italic=False)

    if not omml_body: 
        return None
    full_xml = f'<m:oMath {nsdecls("m")}>{omml_body}</m:oMath>'
    try:
        return parse_xml(full_xml)
    except Exception:
        return None

def process_runs_with_math(paragraph, text):
    """Trạm điều phối trung tâm tối cao vạn năng hỗ trợ đa phân hệ"""
    text_clean = text.strip()
    if not text_clean: 
        return

    # Đồng bộ 4 kiểu bao bọc về một phom đô-la duy nhất
    text_clean = text_clean.replace(r'$$', '$').replace(r'\[', '$').replace(r'\]', '$')
    text_clean = text_clean.replace(r'\(', '$').replace(r'\)', '$')
    
    parts = re.split(r'(\$.*?\$)', text_clean)
    
    for part in parts:
        if not part: 
            continue
            
        if part.startswith('$') and part.endswith('$'):
            math_content = part[1:-1].strip()
            if math_content:
                math_element = convert_latex_to_omml(math_content)
                if math_element is not None:
                    paragraph._p.append(math_element)
                else:
                    run = paragraph.add_run(part)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
        else:
            # Bộ lọc Hóa học tự động nhận biết chất đứng độc lập trong câu text thường
            chem_tokens = re.split(r'(\b[A-Z][a-z]?\d*(?:\^[-+2a-zA-Z0-9]+)?(?:[A-Z][a-z]?\d*(?:\^[-+2a-zA-Z0-9]+)?)*\b)', part)
            for c_part in chem_tokens:
                if not c_part: 
                    continue
                if re.match(r'^[A-Z]', c_part) and any(char.isdigit() or char in '^+- ' for char in c_part):
                    chem_element = convert_chemical_to_omml(c_part)
                    if chem_element is not None:
                        paragraph._p.append(chem_element)
                    else:
                        run = paragraph.add_run(c_part)
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
                else:
                    run = paragraph.add_run(c_part)
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
