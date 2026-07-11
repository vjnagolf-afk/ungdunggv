import re

def format_latex_response(raw_text: str) -> str:
    """
    Module trung gian chuẩn hóa toàn bộ chuỗi văn bản và công thức Toán - Lý - Hóa
    trước khi chuyển tiếp đến giao diện Streamlit.
    """
    if not raw_text:
        return ""

    text = str(raw_text)

    # 1. Khôi phục các ký tự xuống dòng và tab thô từ chuỗi JSON/LangChain
    text = text.replace('\\n', '\n').replace('\\t', '\t')

    # 2. Sửa lỗi OCR dính chữ phổ biến (ví dụ: fracst -> \frac{s}{t})
    # Tự động phát hiện và chuyển đổi cụm dính chữ 'fracst' thành dạng phân số chuẩn
    text = re.sub(r'fracst\b', r'\\frac{s}{t}', text)
    text = re.sub(r'frac\s*s\s*t\b', r'\\frac{s}{t}', text)

    # 3. Đảm bảo các hàm toán học tiếng Việt trong công thức LaTeX được bọc qua \text{...}
    # Ví dụ: biến $\textTốc độ = ...$ thành $\text{Tốc độ} = ...$
    text = re.sub(r'\\text\s*([^\{][^\$\n]*)', r'\\text{\1}', text)

    # 4. Bảo vệ dấu gạch chéo ngược khỏi bộ lọc escape của Markdown Streamlit
    # Trong Streamlit Markdown, dấu \ đứng trước một số ký tự đặc biệt có thể bị nuốt mất.
    # Chúng ta chỉ nhân đôi các dấu \ đi liền với các từ khóa LaTeX phổ biến để tránh làm vỡ định dạng tổng thể.
    latex_keywords = [
        'frac', 'sqrt', 'alpha', 'beta', 'gamma', 'delta', 'pi', 'mu', 'rho', 'sigma', 'tau', 'omega',
        'times', 'div', 'pm', 'mp', 'le', 'ge', 'leq', 'geq', 'neq', 'approx', 'equiv', 'cdot',
        'text', 'right', 'left', 'sum', 'int', 'partial', 'nabla', 'degree', 'rightarrow', 'longrightarrow'
    ]
    
    for kw in latex_keywords:
        # Thay thế \keyword thành \\keyword một cách an toàn bằng regex
        text = re.sub(r'\\(' + kw + r')\b', r'\\\\\1', text)

    return text
