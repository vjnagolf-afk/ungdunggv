import sys
import os
import ast
import streamlit as st
from datetime import datetime

# 1. Ép đường dẫn
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))

from ai_service import run_ai_prompt_safe 
from .processor import process_and_vectorize, query_rag, backup_to_googlesheet

def extract_text_safely(raw_output):
    """
    Hàm lột vỏ dữ liệu chuyên dụng: Cắt bỏ chữ ký bảo mật Base64 của LangChain/Gemini 
    và trích xuất văn bản sạch một cách an toàn.
    """
    text_str = str(raw_output).strip()
    
    # Kỹ thuật cắt chuỗi thủ công để vượt qua các lỗi mã hóa phức tạp
    marker_start = "'text': '"
    marker_end = "', 'extras': {"
    
    if marker_start in text_str and marker_end in text_str:
        start_idx = text_str.find(marker_start) + len(marker_start)
        end_idx = text_str.rfind(marker_end)
        extracted_text = text_str[start_idx:end_idx]
        # Khôi phục lại các dấu xuống dòng bị chuỗi hóa
        return extracted_text.replace('\\n', '\n').replace('\\t', '\t')
        
    # Dự phòng an toàn cho các định dạng khác
    if text_str.startswith("[{"):
        try:
            parsed = ast.literal_eval(text_str)
            if isinstance(parsed, list) and isinstance(parsed[0], dict):
                return parsed[0].get('text', text_str)
        except Exception:
            pass
            
    return text_str

def render_rag():
    st.subheader("🤖 AI Hỏi - Đáp Theo Tài Liệu (RAG)")
    st.write("Hệ thống tự động phân tách tài liệu, nhúng vector và truy xuất dữ liệu có kèm trích dẫn nguồn.")

    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []

    # ==========================================
    # BƯỚC 1: CHỌN NGUỒN TÀI LIỆU
    # ==========================================
    st.markdown("### 📥 Bước 1: Chọn nguồn tài liệu giảng dạy")
    
    source_type = st.radio(
        "Hình thức cung cấp học liệu:",
        ["Tài liệu tải lên (PDF, DOCX, Ảnh)", "Đường dẫn Website"],
        horizontal=True
    )

    temp_path = None

    if "Tài liệu tải lên" in source_type:
        uploaded_file = st.file_uploader(
            "Tải lên tài liệu của thầy/cô (PDF, DOCX, PNG, JPG):", 
            type=["pdf", "docx", "png", "jpg", "jpeg"]
        )
        if uploaded_file:
            os.makedirs("temp_dir", exist_ok=True)
            temp_path = os.path.join("temp_dir", "temp_" + uploaded_file.name)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
    else:
        web_url = st.text_input("Nhập địa chỉ URL của bài báo hoặc website học liệu:", placeholder="https://...")
        if web_url:
            temp_path = web_url  

    if temp_path:
        if st.button("🚀 Bắt đầu phân tích tài liệu", type="primary"):
            with st.spinner("🔄 Hệ thống đang thực hiện: OCR/Parser ➔ Chia nhỏ đoạn (Chunk) ➔ Nhúng Vector DB..."):
                try:
                    vectorstore = process_and_vectorize(temp_path)
                    st.session_state["vectorstore"] = vectorstore
                    st.success("🎉 Cấu hình Vector Database hoàn tất! Thầy/cô đã có thể bắt đầu hỏi đáp.")
                except Exception as e:
                    st.error(f"❌ Lỗi xử lý tài liệu: {str(e)}. Vui lòng kiểm tra lại định dạng file.")

    st.markdown("---")

    # ==========================================
    # BƯỚC 2: KHÔNG GIAN HỎI ĐÁP
    # ==========================================
    st.markdown("### 💬 Bước 2: Tương tác và truy vấn với AI")

    if "vectorstore" not in st.session_state:
        st.info("💡 Vui lòng hoàn thành **Bước 1** (Tải tài liệu và bấm phân tích) để kích hoạt Trợ lý AI.")
    else:
        for message in st.session_state["chat_history"]:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                if "model_info" in message:
                    st.caption(f"⚡ Sinh bởi: {message['model_info']}")

        if question := st.chat_input("Nhập câu hỏi của thầy/cô về tài liệu..."):
            
            with st.chat_message("user"):
                st.markdown(question)
            st.session_state["chat_history"].append({"role": "user", "content": question})

            with st.chat_message("assistant"):
                with st.spinner("🔍 AI đang truy xuất các đoạn ngữ cảnh liên quan..."):
                    try:
                        context = query_rag(st.session_state["vectorstore"], question)
                        
                        # --- CẬP NHẬT LỆNH BẮT BUỘC SỬA LỖI TOÁN HỌC ---
                        prompt = f"""Dựa vào ngữ cảnh tài liệu sau đây, hãy trả lời câu hỏi của giáo viên một cách chi tiết.

YÊU CẦU ĐỊNH DẠNG TỐI QUAN TRỌNG:
1. LUÔN đính kèm chính xác nguồn trích dẫn hoặc số trang (nếu có).
2. TỰ ĐỘNG SỬA LỖI CÔNG THỨC: Quá trình quét tài liệu có thể làm dính các ký tự (ví dụ: biến fracst thành \\frac{{s}}{{t}}, hay các ký hiệu Khoa học Tự nhiên bị lỗi). BẠN BẮT BUỘC phải nhận diện và viết lại thành công thức LaTeX chuẩn.
3. Không viết dính liền các đại lượng. Các biến số Toán/Lý/Hóa phải nằm trong dấu $ (ví dụ: $v$, $s$, $t$).
4. Các công thức độc lập phải nằm trên một dòng riêng với dấu $$ (ví dụ: $$v = \\frac{{s}}{{t}}$$).
5. CHỈ trả lời bằng ngôn ngữ tự nhiên được định dạng, tuyệt đối không trả về mảng dữ liệu.

Ngữ cảnh: {context}

Câu hỏi: {question}"""
                        
                        ai_output = run_ai_prompt_safe(prompt)
                        
                        # Xử lý lột vỏ dữ liệu
                        if isinstance(ai_output, tuple):
                            raw_answer = ai_output[0]
                            model_info = ai_output[1] if len(ai_output) > 1 else "AI"
                        else:
                            raw_answer = ai_output
                            model_info = "AI"

                        # Gọi hàm cắt chuỗi phẫu thuật
                        clean_answer = extract_text_safely(raw_answer)
                        
                        st.markdown(clean_answer)
                        st.caption(f"⚡ Sinh bởi: {model_info}")
                        
                        st.session_state["chat_history"].append({
                            "role": "assistant", 
                            "content": clean_answer,
                            "model_info": model_info
                        })
                        
                        backup_to_googlesheet({
                            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'query': question,
                            'response': clean_answer
                        })
                        st.caption("✅ Đã tự động đồng bộ cuộc hội thoại này vào Nhật ký giảng dạy trên Google Sheet!")
                        
                    except Exception as e:
                        st.error(f"💥 Đã xảy ra lỗi trong quá trình xử lý câu hỏi: {str(e)}")

        if st.session_state["chat_history"]:
            if st.button("🗑️ Xóa lịch sử trò chuyện"):
                st.session_state["chat_history"] = []
                st.rerun()
