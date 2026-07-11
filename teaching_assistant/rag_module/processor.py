import streamlit as st
from langchain_community.document_loaders import PyPDFLoader, Docx2txtLoader
from langchain_google_genai import GoogleGenerativeAIEmbeddings
from langchain_community.vectorstores import Chroma
import os
import gspread # Cần cài đặt thư viện này
from oauth2client.service_account import ServiceAccountCredentials

def backup_to_googlesheet(data_dict):
    """
    Gửi thông tin giao tiếp về Google Sheet để quản lý và kiểm tra dữ liệu 
    theo thời gian thực
    """
    try:
        # Cấu hình kết nối Google Sheet qua file JSON credential
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('google_key.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open("Data_Nhat_Ky_Giang_Day").sheet1
        
        # Ghi dữ liệu: Ngày giờ, Nội dung hỏi, Nội dung trả lời
        sheet.append_row([data_dict['timestamp'], data_dict['query'], data_dict['response']])
    except Exception as e:
        st.error(f"Lỗi khi sao lưu dữ liệu lên hệ thống Cloud: {e}")
# Cấu hình Embedding model (Sử dụng Google Gemini Embedding)
def get_embedding_model():
    # Lưu ý: Thay API_KEY bằng biến môi trường hoặc input của giáo viên
    return GoogleGenerativeAIEmbeddings(model="models/embedding-001")

def process_and_vectorize(file_path):
    # 1. Đọc tài liệu
    if file_path.endswith('.pdf'):
        loader = PyPDFLoader(file_path)
    elif file_path.endswith('.docx'):
        loader = Docx2txtLoader(file_path)
    else:
        return None
    
    documents = loader.load()
    
    # 2. Băm văn bản (Chunking)
    # Chia nhỏ văn bản thành các đoạn 1000 ký tự, overlap 200 ký tự để giữ ngữ cảnh
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)
    texts = text_splitter.split_documents(documents)
    
    # 3. Nhúng Vector (Embedding) và lưu vào ChromaDB
    # Lưu tại thư mục tạm để truy vấn nhanh
    vectorstore = Chroma.from_documents(
        documents=texts, 
        embedding=get_embedding_model(),
        persist_directory="./rag_db" 
    )
    return vectorstore

def query_rag(vectorstore, question):
    # Tìm kiếm tài liệu liên quan
    retriever = vectorstore.as_retriever(search_kwargs={"k": 3})
    docs = retriever.get_relevant_documents(question)
    
    # Ghép nội dung để gửi cho AI
    context = "\n".join([d.page_content for d in docs])
    return context
