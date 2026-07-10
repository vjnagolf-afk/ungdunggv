import io
import re
import numpy as np
import matplotlib.pyplot as plt
from docx.oxml import parse_xml
from docx.shared import Pt

# [PHẦN MÃ ĐÃ CUNG CẤP TRƯỚC ĐÓ - ĐẢM BẢO KHÔNG CÓ KÝ TỰ LẠ Ở ĐẦU FILE]

# ================= TỪ ĐIỂN KÝ HIỆU HY LẠP & TOÁN HỌC =================
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

# ================= ĐỘNG CƠ BIÊN DỊCH LATEX SANG OMML WORD =================
def convert_latex_to_omml(latex_str):
    latex_str = latex_str.strip()
    latex_str = latex_str.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    
    # Chuẩn hóa các lỗi gõ của AI (Khoảng trắng dư thừa)
    latex_str = latex_str.replace(r'\dfrac', r'\frac').replace(r'\limits', '').replace(r'\displaystyle', '')
    latex_str = re.sub(r'\s*\{\s*', '{', latex_str)
    latex_str = re.sub(r'\s*\}\s*', '}', latex_str)
    latex_str = re.sub(r'\s*_\s*', '_', latex_str)
    latex_str = re.sub(r'\s*\^\s*', '^', latex_str)

    for k, v in GREEK.items(): latex_str = latex_str.replace(k, v)
    for k, v in SYMBOLS.items(): latex_str = latex_str.replace(k, v)
    latex_str = latex_str.replace(r'\,', ' ').replace(r'\;', ' ')
    latex_str = re.sub(r'\\(text|mathrm)\{([^{}]*)\}', r'\2', latex_str)
    latex_str = re.sub(r'\\(sin|cos|tan|cot|log|ln|lim|max|min|det)', r'\1', latex_str)

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

    chars = r'a-zA-Z0-9_>\]/()ⒻⓕⓃⓝⒹⓓⓆⓠⒺⓔⒷⓑⓅⓟⓈⓢⓊⓤⓍⓧⓎⓨⓋⓥⓂⓜⓇⓡα-ωΑ-Ω∞=+\-.,'
    while True:
        prev = latex_str
        latex_str = re.sub(r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'ⒻⓃ\1ⓝⒹ\2ⓓⓕ', latex_str)
        latex_str = re.sub(r'\\sqrt\{([^{}]+)\}', r'ⓆⒺ\1ⓔⓠ', latex_str)
        latex_str = re.sub(r'\\vec\{([^{}]+)\}', r'ⓋⒺ\1ⓔⓥ', latex_str)
        
        latex_str = re.sub(r'(∫|∑)_\{([^{}]+)\}\^\{([^{}]+)\}', r'ⓍⒺ\1ⓔⓈ\2ⓢⓊ\3ⓤⓧ', latex_str)
        latex_str = re.sub(r'(∫|∑)_([a-zA-Z0-9])\^([a-zA-Z0-9])', r'ⓍⒺ\1ⓔⓈ\2ⓢⓊ\3ⓤⓧ', latex_str)
        
        latex_str = re.sub(rf'([{chars}]+)_\{{([^{{}}]+)\}}\^\{{([^{{}}]+)\}}', r'ⓏⒺ\1ⓔⓈ\2ⓢⓊ\3ⓤⓩ', latex_str)
        latex_str = re.sub(rf'([{chars}]+)\^\{{([^{{}}]+)\}}_\{{([^{{}}]+)\}}', r'ⓏⒺ\1ⓔⓈ\3ⓢⓊ\2ⓤⓩ', latex_str)

        latex_str = re.sub(rf'([{chars}]+)_\{{([^{{}}]+)\}}', r'ⒷⒺ\1ⓔⓈ\2ⓢⓑ', latex_str)
        latex_str = re.sub(rf'([{chars}]+)_([a-zA-Z0-9])', r'ⒷⒺ\1ⓔⓈ\2ⓢⓑ', latex_str)
        
        latex_str = re.sub(rf'([{chars}]+)\^\{{([^{{}}]+)\}}', r'ⓅⒺ\1ⓔⓊ\2ⓤⓟ', latex_str)
        latex_str = re.sub(rf'([{chars}]+)\^([a-zA-Z0-9])', r'ⓅⒺ\1ⓔⓊ\2ⓤⓟ', latex_str)

        if latex_str == prev: break

    MARKERS = "".join(TAG_MAP.keys())
    parts = re.split(f'([{MARKERS}])', latex_str)
    
    RPR = '<m:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="28"/><w:szCs w:val="28"/></m:rPr>'
    
    # [QUAN TRỌNG NHẤT]: Bổ sung namespace w: giúp Word không bị crash khi đọc rPr!
    omml_xml = f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    
    for p in parts:
        if p in TAG_MAP:
            omml_xml += TAG_MAP[p]
        elif p.strip() or p == ' ':
            omml_xml += f'<m:r>{RPR}<m:t>{p}</m:t></m:r>'
            
    omml_xml += '</m:oMath>'
    
    try:
        return parse_xml(omml_xml)
    except Exception as e:
        return None

def process_runs_with_math(paragraph, text):
    text_clean = text.strip()
    
    # Auto-Chemistry (Chỉ áp dụng cho text thường)
    ELEMENTS = "H|He|Li|Be|B|C|N|O|F|Ne|Na|Mg|Al|Si|P|S|Cl|Ar|K|Ca|Sc|Ti|V|Cr|Mn|Fe|Co|Ni|Cu|Zn|Ag|Cd|I|Ba|Pt|Au|Hg|Pb"
    text_clean = re.sub(rf'\b({ELEMENTS})(\d+)\^([0-9]*[+-])', r'\1<sub>\2</sub><sup>\3</sup>', text_clean)
    text_clean = re.sub(rf'\b({ELEMENTS})\^([0-9]*[+-])', r'\1<sup>\2</sup>', text_clean)
    text_clean = re.sub(rf'\b({ELEMENTS})(\d+)\b', r'\1<sub>\2</sub>', text_clean)

    # Bóc tách Đa định dạng LaTeX
    delimiter_pattern = r'(\$\$[\s\S]*?\$\$|\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\)|\$[\s\S]*?\$)'
    parts = re.split(delimiter_pattern, text_clean)
    
    for part in parts:
        if not part: continue
        
        is_math = False
        math_content = ""
        if part.startswith('$$') and part.endswith('$$'):
            math_content = part[2:-2].strip(); is_math = True
        elif part.startswith('\\[') and part.endswith('\\]'):
            math_content = part[2:-2].strip(); is_math = True
        elif part.startswith('\\(') and part.endswith('\\)'):
            math_content = part[2:-2].strip(); is_math = True
        elif part.startswith('$') and part.endswith('$'):
            math_content = part[1:-1].strip(); is_math = True

        if is_math and math_content:
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

def generate_plot_stream(eq_str):
    fig, ax = plt.subplots(figsize=(5, 3.5))
    if eq_str.lower().strip() == 'scatter':
        x, y = np.random.rand(50) * 10, np.random.rand(50) * 10
        ax.scatter(x, y, color='#1E40AF', alpha=0.7)
        ax.set_title("Đồ thị phân tán (Scatter)", fontsize=10, pad=10)
    else:
        x = np.linspace(-10, 10, 400)
        safe_dict = {"x": x, "np": np, "sin": np.sin, "cos": np.cos, "tan": np.tan, "sqrt": np.sqrt, "log": np.log, "log10": np.log10, "exp": np.exp, "e": np.e, "pi": np.pi}
        try:
            eq_clean = eq_str.lower().replace('y=', '').replace('y =', '').strip()
            eq_py = eq_clean.replace('^', '**').replace('e**x', 'np.exp(x)')
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
