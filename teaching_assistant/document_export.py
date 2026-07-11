from docx import Document
from latex_formatter import process_science_formulas

def export_to_docx(content, filename="tai_lieu.docx"):
    doc = Document()
    
    # 1. Làm sạch nội dung qua bộ lọc LaTeX
    clean_content = process_science_formulas(content)
    
    # 2. Xử lý chia nhỏ đoạn văn (giả sử AI trả về các dòng cách nhau bằng \n)
    lines = clean_content.split('\n')
    
    for line in lines:
        if line.strip():
            # Thêm vào docx và loại bỏ các ký hiệu LaTeX thừa cho Word
            # Ví dụ: thay thế $v = \frac{s}{t}$ thành "v = s/t" để Word dễ đọc
            plain_text = line.replace('$', '').replace('\\frac', '')
            doc.add_paragraph(plain_text)
            
    doc.save(filename)
    return filename
