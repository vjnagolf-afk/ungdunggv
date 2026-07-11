import sys
import os
import ast
import streamlit as st
from datetime import datetime

# 1. Ép đường dẫn để import các file ở thư mục gốc (app.py, ai_service.py,...)
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../')))

# 2. Import các module cần thiết
from ai_service import run_ai_prompt_safe 
from .processor import process_and_vectorize, query_rag, backup_to_googlesheet

def render_rag():
    st.subheader("🤖 AI Hỏi - Đáp Theo Tài Liệu (RAG)")
    st.write("Hệ thống tự động phân tách tài liệu, nhúng vector và truy xuất dữ liệu có kèm trích dẫn nguồn.")

    # Khởi tạo lịch sử trò chuyện nếu chưa có
    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []

    # ==========================================
    # BƯỚC 1: CHỌN NGUỒN TÀI LIỆU (Theo sơ đồ)
    # ==========================================
    st.markdown("### 📥 Bước 1: Chọn nguồn tài liệu giảng dạy")
    
    # Cho phép giáo viên chọn nguồn linh hoạt theo sơ đồ
    source_type = st.radio(
        "Hình thức cung cấp học liệu:",
        ["Tài liệu tải lên (PDF, DOCX, Ảnh)", "Đường dẫn Website"],
        horizontal=True
    )

    is_processed = False
    temp_path = None

    if "Tài liệu tải lên" in source_type:
        uploaded_file = st.file_uploader(
            "Tải lên tài liệu của thầy/cô (PDF, DOCX, PNG, JPG):", 
            type=["pdf", "docx", "png", "jpg", "jpeg"]
        )
        if uploaded_file:
            # Tạo thư mục tạm an toàn nếu chưa có
            os.makedirs("temp_dir", exist_ok=True)
            temp_path = os.path.join("temp_dir", "temp_" + uploaded_file.name)
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
    else:
        web_url = st.text_input("Nhập địa chỉ URL của bài báo hoặc website học liệu:", placeholder="https://...")
        if web_url:
            temp_path = web_url  # Gửi trực tiếp URL sang hàm xử lý ở processor.py

    # Nút bấm kích hoạt chuỗi xử lý: OCR + Parser -> Chunk -> Embedding -> Vector DB
    if temp_path:
        if st.button("🚀 Bắt đầu phân tích tài liệu", type="primary"):
            with st.spinner("🔄 Hệ thống đang thực hiện: OCR/Parser ➔ Chia nhỏ đoạn (Chunk) ➔ Nhúng Vector DB..."):
                try:
                    # Gọi hàm xử lý cốt lõi từ processor.py của bạn
                    vectorstore = process_and_vectorize(temp_path)
                    st.session_state["vectorstore"] = vectorstore
                    st.success("🎉 Cấu hình Vector Database hoàn tất! Thầy/cô đã có thể bắt đầu hỏi đáp.")
                    is_processed = True
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
        # Hiển thị lại toàn bộ lịch sử cuộc hội thoại trực quan
        for message in st.session_state["chat_history"]:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                if "model_info" in message:
                    st.caption(f"⚡ Sinh bởi: {message['model_info']}")

        # Ô nhập câu hỏi dạng Chat Input hiện đại ở cuối trang
        if question := st.chat_input("Nhập câu hỏi của thầy/cô về tài liệu..."):
            
            # 1. Hiển thị ngay câu hỏi của người dùng
            with st.chat_message("user"):
                st.markdown(question)
            st.session_state["chat_history"].append({"role": "user", "content": question})

            # 2. Xử lý logic RAG tìm kiếm và sinh câu trả lời
            with st.chat_message("assistant"):
                with st.spinner("🔍 AI đang truy xuất các đoạn ngữ cảnh liên quan..."):
                    try:
                        # Lấy dữ liệu trích dẫn/ngữ cảnh từ processor
                        context = query_rag(st.session_state["vectorstore"], question)
                        
                        # Thiết lập prompt chuẩn bám sát yêu cầu trích dẫn số trang/nguồn
                        prompt = f"""Dựa vào ngữ cảnh tài liệu sau đây, hãy trả lời câu hỏi của giáo viên một cách chi tiết.

YÊU CẦU ĐỊNH DẠNG:
- Hãy luôn đính kèm chính xác nguồn trích dẫn hoặc số trang (nếu có trong ngữ cảnh) ở cuối câu trả lời.
- BẮT BUỘC sử dụng cú pháp LaTeX chuẩn cho TẤT CẢ các công thức Toán, Khoa học Tự nhiên (Vật lý, Hóa học) và đặt trong dấu $ (ví dụ: $v = \\frac{{s}}{{t}}$) hoặc $$ (nằm độc lập trên một dòng).
- Trình bày mạch lạc, xuống dòng rõ ràng, không viết dính liền công thức.

Ngữ cảnh: {context}

Câu hỏi: {question}"""
                        
                        # Gọi dịch vụ AI an toàn của bạn
                        ai_output = run_ai_prompt_safe(prompt)
                        
                        # --- BỘ LỌC LÀM SẠCH DỮ LIỆU ---
                        answer = ""
                        model_info = "AI"
                        
                        if isinstance(ai_output, tuple):
                            answer = ai_output[0]
                            model_info = ai_output[1] if len(ai_output) > 1 else "AI"
                        else:
                            answer = ai_output

                        # Bóc tách cấu trúc mảng JSON nếu có
                        if isinstance(answer, str) and answer.strip().startswith("[{"):
                            try:
                                parsed_answer = ast.literal_eval(answer.strip())
                                if isinstance(parsed_answer, list) and isinstance(parsed_answer[0], dict):
                                    answer = parsed_answer[0].get('text', answer)
                            except Exception:
                                pass
                        
                        # Đảm bảo answer là chuỗi
                        answer = str(answer)
                        # -------------------------------
                        
                        # Hiển thị câu trả lời lên màn hình
                        st.markdown(answer)
                        st.caption(f"⚡ Sinh bởi: {model_info}")
                        
                        st.session_state["chat_history"].append({
                            "role": "assistant", 
                            "content": answer,
                            "model_info": model_info
                        })
                        
                        # 3. Sao lưu tự động vào Google Sheet nhật ký giảng dạy
                        backup_to_googlesheet({
                            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            'query': question,
                            'response': answer
                        })
                        st.caption("✅ Đã tự động đồng bộ cuộc hội thoại này vào Nhật ký giảng dạy trên Google Sheet!")
                        
                    except Exception as e:
                        st.error(f"💥 Đã xảy ra lỗi trong quá trình xử lý câu hỏi: {str(e)}")

        # Nút hỗ trợ xóa lịch sử hội thoại nhanh nếu cần
        if st.session_state["chat_history"]:
            if st.button("🗑️ Xóa lịch sử trò chuyện"):
                st.session_state["chat_history"] = []
                st.rerun()
