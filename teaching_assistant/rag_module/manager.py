import sys
import os
import ast
import streamlit as st
from datetime import datetime

# 1. Ép đường dẫn để import các file ở thư mục gốc
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))

# 2. Import các module cần thiết
from ai_service import run_ai_prompt_safe 
from .processor import process_and_vectorize, query_rag, backup_to_googlesheet
# TÍCH HỢP MODULE XỬ LÝ CÔNG THỨC TOÁN/LÝ/HÓA
from .latex_formatter import process_science_formulas

def extract_text_safely(raw_output):
    """
    Hàm lột vỏ dữ liệu: Cắt bỏ chữ ký bảo mật Base64 của LangChain
    và trích xuất văn bản sạch.
    """
    text_str = str(raw_output).strip()
    marker_start = "'text': '"
    marker_end = "', 'extras': {"
    
    if marker_start in text_str and marker_end in text_str:
        start_idx = text_str.find(marker_start) + len(marker_start)
        end_idx = text_str.rfind(marker_end)
        return text_str[start_idx:end_idx]
    
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
    
    # Khởi tạo lịch sử trò chuyện
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
            with st.spinner("⚙️ Hệ thống đang thực hiện: OCR/Parser ➔ Chia nhỏ đoạn (Chunk) ➔ Nhúng Vector DB..."):
                try:
                    vectorstore = process_and_vectorize(temp_path)
                    st.session_state["vectorstore"] = vectorstore
                    st.success("🎉 Cấu hình Vector Database hoàn tất! Thầy/cô đã có thể bắt đầu hỏi đáp.")
                except Exception as e:
                    st.error(f"❌ Lỗi xử lý tài liệu: {str(e)}. Vui lòng kiểm tra lại định dạng file.")
                    
    st.markdown("---")
    
    # ==========================================
    # BƯỚC 2: KHÔNG GIAN HỎI ĐÁP (Giao diện Chatbot)
    # ==========================================
    st.markdown("### 💬 Bước 2: Tương tác và truy vấn với AI")
    
    if "vectorstore" not in st.session_state:
        st.info("💡 Vui lòng hoàn thành **Bước 1** (Tải tài liệu và bấm phân tích) để kích hoạt Trợ lý AI.")
    else:
        # Hiển thị lịch sử hội thoại
        for message in st.session_state["chat_history"]:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                if "model_info" in message and message["model_info"]:
                    st.caption(f"🤖 Sinh bởi: {message['model_info']}")
                    
        if question := st.chat_input("Nhập câu hỏi của thầy/cô về tài liệu..."):
            
            # --- HIỂN THỊ CÂU HỎI CỦA NGƯỜI DÙNG LÊN GIAO DIỆN ---
            with st.chat_message("user"):
                st.markdown(question)
            st.session_state["chat_history"].append({"role": "user", "content": question})
            
            with st.chat_message("assistant"):
                with st.spinner("🔍 AI đang truy xuất các đoạn ngữ cảnh liên quan..."):
                    model_info = "AI"
                    final_response = ""
                    
                    try:
                        context = query_rag(st.session_state["vectorstore"], question)
                        
                        prompt = f"""Dựa vào ngữ cảnh tài liệu sau đây, hãy trả lời câu hỏi của giáo viên một cách chi tiết.
YÊU CẦU ĐỊNH DẠNG SƯ PHẠM CHUẨN:
- LUÔN đính kèm chính xác nguồn trích dẫn hoặc số trang (nếu có ngữ cảnh).
- BẮT BUỘC sử dụng cú pháp LaTeX chuẩn cho các công thức Khoa học Tự nhiên và đặt trong dấu $ (ví dụ: $v = \\frac{{s}}{{t}}$) hoặc $$ (nằm độc lập trên một dòng).
- Tự động sửa lỗi OCR dính chữ (ví dụ: fracst phải viết lại thành công thức phân số chuẩn).

Ngữ cảnh: {context}
Câu hỏi: {question}"""
                        
                        # 1. Gọi dịch vụ AI
                        ai_output = run_ai_prompt_safe(prompt)
                        
                        # 2. Tách Tuple (nội dung, tên model)
                        if isinstance(ai_output, tuple):
                            raw_answer = ai_output[0]
                            model_info = ai_output[1] if len(ai_output) > 1 else "AI"
                        else:
                            raw_answer = ai_output
                            model_info = "AI"
                            
                        # 3. Làm sạch mảng JSON của LangChain
                        clean_response = extract_text_safely(raw_answer)
                        
                        # 4. --- GỌI MODULE CHUẨN HÓA CÔNG THỨC TRUNG GIAN ---
                        final_response = process_science_formulas(clean_response)
                        
                        # 5. Hiển thị kết quả trực tiếp
                        st.markdown(final_response)
                        st.caption(f"🤖 Sinh bởi: {model_info}")
                        
                        # Lưu lịch sử
                        st.session_state["chat_history"].append({
                            "role": "assistant", 
                            "content": final_response,
                            "model_info": model_info
                        })
                        
                        # 6. Sao lưu tự động bảo mật qua cấu hình st.secrets
                        if "gcp_service_account" in st.secrets:
                            creds_data = dict(st.secrets["gcp_service_account"])
                            backup_to_googlesheet({
                                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'query': question,
                                'response': final_response
                            }, google_creds=creds_data)
                            st.caption("✅ Đã tự động đồng bộ cuộc hội thoại này vào Nhật ký giảng dạy trên Google Sheet!")
                            
                    except Exception as e:
                        st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý câu hỏi: {str(e)}")
                        
        # Nút xóa lịch sử trò chuyện
        if st.session_state["chat_history"]:
            if st.button("🗑️ Xóa lịch sử trò chuyện"):
                st.session_state["chat_history"] = []
                st.rerun()
