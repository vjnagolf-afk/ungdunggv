import streamlit as st
from langchain_community.document_loaders import PyPDFLoader, Docx2txtLoader
from langchain_core.documents import Document
from langchain_google_genai import GoogleGenerativeAIEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Chroma
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

    """
    Gửi thông tin giao tiếp về Google Sheet bằng tài khoản dịch vụ
    """
    try:
        client = gspread.service_account_from_dict(google_creds)
        sheet = client.open("Data_Nhat_Ky_Giang_Day").sheet1
        sheet.append_row([data_dict['timestamp'], data_dict['query'], data_dict['response']])
        return True
    except Exception as e:
        st.error(f"Lỗi khi sao lưu dữ liệu lên hệ thống Cloud: {e}")
        return False

def get_embedding_model():
    """
    Cấu hình Embedding model sử dụng mô hình Gemini chính xác
    """
    return GoogleGenerativeAIEmbeddings(model="models/gemini-embedding-001")

def process_and_vectorize(file_path):
    """
    Sử dụng cơ chế Lazy Import để nạp bộ loader tương ứng khi cần,
    giúp tránh lỗi sập ứng dụng khi hệ thống thiếu thư viện hệ thống của OCR.
    """
    documents = []
    ext = os.path.splitext(file_path)[-1].lower()
    
    if ext == ".pdf":
        # Chỉ nạp PyPDFLoader khi người dùng tải lên file PDF
        from langchain_community.document_loaders import PyPDFLoader
        loader = PyPDFLoader(file_path)
        documents = loader.load()
        
    elif ext == ".docx":
        # Chỉ nạp Docx2txtLoader khi người dùng tải lên file Word
        from langchain_community.document_loaders import Docx2txtLoader
        loader = Docx2txtLoader(file_path)
        documents = loader.load()
        
    elif ext in [".png", ".jpg", ".jpeg", ".tiff", ".bmp"]:
        # Chỉ nạp bộ OCR khi người dùng tải lên ảnh
        try:
            from langchain_community.document_loaders import RapidOCRLoader
            loader = RapidOCRLoader(file_path)
            documents = loader.load()
        except ImportError as e:
            st.error("Không thể sử dụng tính năng quét chữ từ ảnh (OCR) do máy chủ thiếu thư viện hệ thống.")
            st.info("Vui lòng bổ sung tệp 'packages.txt' chứa dòng chữ 'libgl1' lên GitHub để khắc phục.")
            raise e
    else:
        raise ValueError(f"Định dạng tệp '{ext}' không được hỗ trợ.")
        
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000, 
        chunk_overlap=200
    )
    splits = text_splitter.split_documents(documents)
    
    embedding_model = get_embedding_model()
    vectorstore = Chroma.from_documents(documents=splits, embedding=embedding_model)
    return vectorstore

def query_rag(vectorstore, question):
    """
    Tìm kiếm tài liệu liên quan
    """
    retriever = vectorstore.as_retriever(search_kwargs={"k": 3})
    docs = retriever.invoke(question)
    return docs
