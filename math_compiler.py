# math_compiler.py - Bản Nâng Cấp Tối Thượng (Toán - Lý - Hóa - Graph)
import io
import re
import numpy as np
import matplotlib.pyplot as plt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt

# ================= TỪ ĐIỂN KÝ HIỆU HY LẠP & TOÁN HỌC CƠ BẢN =================
GREEK = {
    r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ', r'\epsilon': 'ε', r'\varepsilon': 'ε',
    r'\zeta': 'ζ', r'\eta': 'η', r'\theta': 'θ', r'\vartheta': 'ϑ', r'\iota': 'ι', r'\kappa': 'κ',
    r'\lambda': 'λ', r'\mu': 'μ', r'\nu': 'ν', r'\xi': 'ξ', r'\pi': 'π',
    r'\rho': 'ρ', r'\sigma': 'σ', r'\tau': 'τ', r'\upsilon': 'υ', r'\phi': 'φ', r'\varphi': 'ϕ',
    r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
    r'\Delta': 'Δ', r'\Gamma': 'Γ', r'\Theta': 'Θ', r'\Lambda': 'Λ', r'\Pi': 'Π',
    r'\Sigma': 'Σ', r'\Phi': 'Φ', r'\Psi': 'Ψ', r'\Omega': 'Ω'
}

SYMBOLS = {
    r'\infty': '∞', r'\rightarrow': '→', r'\leftarrow': '←', r'\Rightarrow': '⇒',
    r'\Leftarrow': '⇐', r'\leftrightarrow': '↔', r'\Leftrightarrow': '⇔',
    r'\approx': '≈', r'\neq': '≠', r'\leq': '≤', r'\geq': '≥', r'\times': '×',
    r'\cdot': '·', r'\pm': '±', r'\mp': '∓', r'\circ': '°', r'\partial': '∂',
    r'\nabla': '∇', r'\forall': '∀', r'\exists': '∃', r'\in': '∈', r'\notin': '∉',
    r'\subset': '⊂', r'\supset': '⊃', r'\cup': '∪', r'\cap': '∩', r'\equiv': '≡',
    r'\sim': '∼', r'\propto': '∝', r'\angle': '∠', r'\parallel': '∥', r'\perp': '⊥',
    r'\int': '∫', r'\sum': '∑'
}

# Bảng mã Marker Unicode tạo cấu trúc XML phân lớp cho Word Equation
TAG_MAP = {
    'Ⓕ': '<m:f>', 'ⓕ': '</m:f>',
    'Ⓝ': '<m:num>', 'ⓝ': '</m:num>',
    'Ⓓ': '<m:den>', 'ⓓ': '</m:den>',
    'Ⓠ': '<m:rad><m:radPr><m:deg m:val=""/></m:radPr>', 'ⓠ': '</m:rad>',
    'Ⓔ': '<m:e>', 'ⓔ': '</m:e>',
    'Ⓑ': '<m:sSub>', 'ⓑ': '</m:sSub>',
    'Ⓟ': '<m:sSup>', 'ⓟ': '</m:sSup>',
    'Ⓩ': '<m:sSubSup>', 'ⓩ': '</m:sSubSup>',
    'Ⓢ': '<m:sub>', 'ⓢ': '</m:sub>',
    'Ⓤ': '<m:sup>', 'ⓤ': '</m:sup>',
    'Ⓧ': '<m:nary><m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="subSup"/></m:naryPr>', 'ⓧ': '</m:nary>',
    'Ⓨ': '<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>', 'ⓨ': '</m:nary>',
    'Ⓥ': '<m:acc><m:accPr><m:chr m:val="&#x20D7;"/></m:accPr>', 'ⓥ': '</m:acc>',
    'Ⓜ': '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="10"/><m:mcJc m:val="c"/></m:mcPr></m:mc></m:mcs></m:mPr>', 'ⓜ': '</m:m>',
    'Ⓡ': '<m:mr>', 'ⓡ': '</m:mr>'
}

# ================= ĐỘNG CƠ BIÊN DỊCH LATEX -> WORD EQUATION =================
def convert_latex_to_omml(latex_str):
    # 1. Bảo vệ cấu trúc XML
    latex_str = latex_str.strip()
    latex_str = latex_str.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    
    # 2. Xử lý khả năng chịu lỗi (Fault Tolerance) cao cấp
    latex_str = latex_str.replace(r'\dfrac', r'\frac')
    latex_str = latex_str.replace(r'\limits', '')
    latex_str = latex_str.replace(r'\displaystyle', '')
    # Tự động dọn dẹp khoảng trắng thừa bọc quanh ngoặc nhọn, chỉ số của AI
    latex_str = re.sub(r'\s*\{\s*', '{', latex_str)
    latex_str = re.sub(r'\s*\}\s*', '}', latex_str)
    latex_str = re.sub(r'\s*_\s*', '_', latex_str)
    latex_str = re.sub(r'\s*\^\s*', '^', latex_str)

    # 3. Chuyển đổi ký hiệu, Text và các hàm thông dụng
    for k, v in GREEK.items(): latex_str = latex_str.replace(k, v)
    for k, v in SYMBOLS.items(): latex_str = latex_str.replace(k, v)
    latex_str = latex_str.replace(r'\,', ' ').replace(r'\;', ' ')
    latex_str = re.sub(r'\\(text|mathrm)\{([^{}]*)\}', r'\2', latex_str)
    latex_str = re.sub(r'\\(sin|cos|tan|cot|log|ln|lim|max|min|det)', r'\1', latex_str) # Chuyển hàm thành text thường trong OMML

    # 4. Xử lý Ma trận (matrix, pmatrix, vmatrix...)
    def repl_matrix(m):
        rows = m.group(1).split('\\\\')
        res = 'Ⓜ'
        for row in rows:
            res += 'Ⓡ'
            cols = row.split('&')
            for col in cols: res += f'Ⓔ{col.strip()}ⓔ'
            res += 'ⓡ'
        res += 'ⓜ'
        return res
    latex_str = re.sub(r'\\begin\{(?:p|b|v|V)?matrix\}(.*?)\\end\{(?:p|b|v|V)?matrix\}', repl_matrix, latex_str, flags=re.DOTALL)

    # 5. Phân tách vòng lặp cấu trúc toán học từ trong ra ngoài
    chars = r'a-zA-Z0-9_>\]/()ⒻⓕⓃⓝⒹⓓⓆⓠⒺⓔⒷⓑⓅⓟⓈⓢⓊⓤⓍⓧⓎⓨⓋⓥⓂⓜⓇⓡα-ωΑ-Ω∞=+\-.,'
    while True:
        prev = latex_str
        
        # Phân số, căn, vector
        latex_str = re.sub(r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'ⒻⓃ\1ⓝⒹ\2ⓓⓕ', latex_str)
        latex_str = re.sub(r'\\sqrt\{([^{}]+)\}', r'ⓆⒺ\1ⓔⓠ', latex_str)
        latex_str = re.sub(r'\\vec\{([^{}]+)\}', r'ⓋⒺ\1ⓔⓥ', latex_str)
        
        # Tích phân / Tổng Sigma
        latex_str = re.sub(r'(∫|∑)_\{([^{}]+)\}\^\{([^{}]+)\}', r'ⓍⒺ\1ⓔⓈ\2ⓢⓊ\3ⓤⓧ', latex_str)
        latex_str = re.sub(r'(∫|∑)_([a-zA-Z0-9])\^([a-zA-Z0-9])', r'ⓍⒺ\1ⓔⓈ\2ⓢⓊ\3ⓤⓧ', latex_str)
        
        # Chỉ số trên và dưới lồng ghép (Vật lý/Hóa học: SO_4^{2-})
        latex_str = re.sub(rf'([{chars}]+)_\{{([^{{}}]+)\}}\^\{{([^{{}}]+)\}}', r'ⓏⒺ\1ⓔⓈ\2ⓢⓊ\3ⓤⓩ', latex_str)
        latex_str = re.sub(rf'([{chars}]+)\^\{{([^{{}}]+)\}}_\{{([^{{}}]+)\}}', r'ⓏⒺ\1ⓔⓈ\3ⓢⓊ\2ⓤⓩ', latex_str)

        # Chỉ số dưới độc lập
        latex_str = re.sub(rf'([{chars}]+)_\{{([^{{}}]+)\}}', r'ⒷⒺ\1ⓔⓈ\2ⓢⓑ', latex_str)
        latex_str = re.sub(rf'([{chars}]+)_([a-zA-Z0-9])', r'ⒷⒺ\1ⓔⓈ\2ⓢⓑ', latex_str)
        
        # Chỉ số trên độc lập
        latex_str = re.sub(rf'([{chars}]+)\^\{{([^{{}}]+)\}}', r'ⓅⒺ\1ⓔⓊ\2ⓤⓟ', latex_str)
        latex_str = re.sub(rf'([{chars}]+)\^([a-zA-Z0-9])', r'ⓅⒺ\1ⓔⓊ\2ⓤⓟ', latex_str)

        if latex_str == prev: break

    # 6. Lắp ráp Word Equation XML
    MARKERS = "".join(TAG_MAP.keys())
    parts = re.split(f'([{MARKERS}])', latex_str)
    
    RPR = '<m:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="28"/><w:szCs w:val="28"/></m:rPr>'
    omml_xml = f'<m:oMath {nsdecls("m")}>'
    
    for p in parts:
        if p in TAG_MAP:
            omml_xml += TAG_MAP[p]
        elif p.strip() or p == ' ':
            omml_xml += f'<m:r>{RPR}<m:t>{p}</m:t></m:r>'
            
    omml_xml += '</m:oMath>'
    
    try:
        return parse_xml(omml_xml)
    except:
        return None

# ================= ENGINE BÓC TÁCH TEXT & AUTO-CHEMISTRY =================
def process_runs_with_math(paragraph, text):
    """
    Quét mọi thẻ toán ($...$, $$...$$, \(...\), \[...\]).
    Tự động format Hóa Học/Vật lý trong Text thường.
    """
    text_clean = text.strip()
    
    # 1. AUTO-CHEMISTRY BẰNG REGEX (chỉ tác động văn bản thường, an toàn tuyệt đối)
    # Tự động chuyển đổi: H2SO4 -> H₂SO₄, SO4^2- -> SO₄²⁻, Na^+ -> Na⁺
    ELEMENTS = "H|He|Li|Be|B|C|N|O|F|Ne|Na|Mg|Al|Si|P|S|Cl|Ar|K|Ca|Sc|Ti|V|Cr|Mn|Fe|Co|Ni|Cu|Zn|Ga|Ge|As|Se|Br|Kr|Rb|Sr|Y|Zr|Nb|Mo|Tc|Ru|Rh|Pd|Ag|Cd|In|Sn|Sb|Te|I|Xe|Cs|Ba|La|Ce|Pr|Nd|Pm|Sm|Eu|Gd|Tb|Dy|Ho|Er|Tm|Yb|Lu|Hf|Ta|W|Re|Os|Ir|Pt|Au|Hg|Tl|Pb|Bi|Po|At|Rn|Fr|Ra|Ac|Th|Pa|U|Np|Pu|Am|Cm|Bk|Cf|Es|Fm|Md|No|Lr"
    text_clean = re.sub(rf'({ELEMENTS})(\d+)\^([0-9]*[+-])', r'\1<sub>\2</sub><sup>\3</sup>', text_clean)
    text_clean = re.sub(rf'({ELEMENTS})\^([0-9]*[+-])', r'\1<sup>\2</sup>', text_clean)
    text_clean = re.sub(rf'({ELEMENTS})(\d+)', r'\1<sub>\2</sub>', text_clean)

    # 2. QUÉT ĐA ĐỊNH DẠNG LATEX
    delimiter_pattern = r'(\$\$[\s\S]*?\$\$|\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\)|\$[\s\S]*?\$)'
    parts = re.split(delimiter_pattern, text_clean)
    
    for part in parts:
        if not part: continue
        
        is_math = False
        math_content = ""
        # Nhận diện thẻ
        if part.startswith('$$') and part.endswith('$$'):
            math_content = part[2:-2].strip(); is_math = True
        elif part.startswith('\\[') and part.endswith('\\]'):
            math_content = part[2:-2].strip(); is_math = True
        elif part.startswith('\\(') and part.endswith('\\)'):
            math_content = part[2:-2].strip(); is_math = True
        elif part.startswith('$') and part.endswith('$'):
            math_content = part[1:-1].strip(); is_math = True

        # Ghi khối toán học
        if is_math and math_content:
            math_element = convert_latex_to_omml(math_content)
            if math_element is not None:
                paragraph._p.append(math_element)
            else:
                run = paragraph.add_run(part) # Fallback nếu lỗi
                run.font.name = 'Times New Roman'
                run.font.size = Pt(14)
        else:
            # Ghi khối Text (Kèm xử lý in đậm ** và chỉ số trên/dưới HTML)
            bold_parts = part.split('**')
            for i, b_part in enumerate(bold_parts):
                is_bold = (i % 2 != 0)
                sub_sup_parts = re.split(r'(<sub>.*?</sub>|<sup>.*?</sup>)', b_part)
                for s_part in sub_sup_parts:
                    if not s_part: continue
                    run = paragraph.add_run()
                    run.bold = is_bold
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(14)
                    
                    if s_part.startswith('<sub>') and s_part.endswith('</sub>'):
                        run.text = s_part[5:-6]
                        run.font.subscript = True
                    elif s_part.startswith('<sup>') and s_part.endswith('</sup>'):
                        run.text = s_part[5:-6]
                        run.font.superscript = True
                    else:
                        run.text = s_part

# ================= GRAPH ENGINE ĐA CHỨC NĂNG =================
def generate_plot_stream(eq_str):
    fig, ax = plt.subplots(figsize=(5, 3.5))
    
    if eq_str.lower().strip() == 'scatter':
        x = np.random.rand(50) * 10
        y = np.random.rand(50) * 10
        ax.scatter(x, y, color='#1E40AF', alpha=0.7)
        ax.set_title("Đồ thị phân tán (Scatter)", fontsize=10, pad=10)
    else:
        x = np.linspace(-10, 10, 400)
        safe_dict = {
            "x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan, 
            "sqrt": np.sqrt, "log": np.log, "log10": np.log10, "exp": np.exp, "e": np.e, "pi": np.pi
        }
        try:
            # Làm sạch chuỗi, AI thường hay viết 'y = ...'
            eq_clean = eq_str.lower().replace('y=', '').replace('y =', '').strip()
            eq_py = eq_clean.replace('^', '**').replace('e**x', 'np.exp(x)')
            # Chữa lỗi tự động 3x -> 3*x
            eq_py = re.sub(r'(\d)(x)', r'\1*\2', eq_py)
            
            y = eval(eq_py, {"__builtins__": {}}, safe_dict)
            if isinstance(y, (int, float)): y = np.full_like(x, y)
            
            ax.plot(x, y, color='#1E40AF', linewidth=2.5)
            ax.axhline(0, color='black', linewidth=1.2)
            ax.axvline(0, color='black', linewidth=1.2)
        except Exception as e:
            ax.text(0.5, 0.5, f"[Lỗi biểu thức: {eq_str}]", ha='center', va='center', color='red')
        
        ax.set_title(f"Đồ thị: y = {eq_str.replace('y=','').replace('y =','')}", fontsize=10, pad=10)

    ax.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
    buf.seek(0)
    plt.close(fig)
    return buf
