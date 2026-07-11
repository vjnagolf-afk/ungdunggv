import re

def process_science_formulas(raw_text: str) -> str:
    """
    Module trung gian chuẩn hóa toàn bộ chuỗi văn bản và công thức Toán - Lý - Hóa
    trước khi chuyển tiếp đến giao diện Streamlit.
    """
    if not raw_text:
        return ""

    text = str(raw_text)

    # 1. Khôi phục các ký tự xuống dòng và tab thô nếu có
    text = text.replace('\\n', '\n').replace('\\t', '\t')

    # 2. Tự động phát hiện và sửa lỗi OCR dính chữ (ví dụ: fracst -> \frac{s}{t})
    text = re.sub(r'\bfracst\b', r'\\frac{s}{t}', text)
    text = re.sub(r'\bfrac\s*s\s*t\b', r'\\frac{s}{t}', text)

    # 3. Đảm bảo các hàm văn bản tiếng Việt trong LaTeX (nếu có) được bọc qua \text{...} đúng chuẩn
    text = re.sub(r'\\text\s*([^\{][^\$\n]*)', r'\\text{\1}', text)

    # 4. GIẢI PHÁP ĐẶC TRỊ LỖI HIỂN THỊ TRÊN STREAMLIT:
    # Bản chất st.markdown coi dấu \ là escape character nên sẽ nuốt mất dấu của các lệnh LaTeX.
    # Ta cần nhân đôi dấu gạch chéo ngược cho các lệnh toán học/khoa học phổ biến.
    latex_keywords = [
        'frac', 'sqrt', 'alpha', 'beta', 'gamma', 'delta', 'pi', 'mu', 'rho', 'sigma', 'tau', 'omega',
        'times', 'div', 'pm', 'mp', 'le', 'ge', 'leq', 'geq', 'neq', 'approx', 'equiv', 'cdot',
        'text', 'right', 'left', 'sum', 'int', 'partial', 'nabla', 'degree', 'rightarrow', 'longrightarrow'
    ]
    
    for kw in latex_keywords:
        # Biến \keyword thành \\keyword một cách an toàn
        text = re.sub(r'\\(' + kw + r')\b', r'\\\\\1', text)

    return text
