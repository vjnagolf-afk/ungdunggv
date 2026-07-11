import re

def process_science_formulas(text: str) -> str:
    """
    Bộ lọc trung gian chuẩn hóa công thức Toán - Lý - Hóa 
    trước khi đưa ra giao diện Streamlit.
    """
    if not isinstance(text, str):
        return str(text)

    # 1. Bảo vệ dấu gạch chéo ngược của LaTeX khỏi Markdown của Streamlit
    clean_text = text.replace('\\', '\\\\')
    
    # 2. Khôi phục các ký tự điều hướng bị lỗi do nhân đôi
    clean_text = clean_text.replace('\\\\n', '\n').replace('\\\\t', '\t')
    
    # 3. Bộ từ điển sửa lỗi OCR tự động cho các đại lượng Khoa học Tự nhiên
    # (Thầy có thể liên tục bổ sung thêm các quy tắc vào đây trong tương lai)
    ocr_corrections = {
        "fracst": r"\\frac{s}{t}",
        "frac ": r"\\frac",
        "extTốc": r"\\text{Tốc",
        "extQuãng": r"\\text{Quãng",
        "extThời": r"\\text{Thời",
        "=\n\\\\frac": "= \\\\frac",
        "= \n\\\\frac": "= \\\\frac"
    }
    
    for error, correction in ocr_corrections.items():
        clean_text = clean_text.replace(error, correction)
        
    return clean_text
