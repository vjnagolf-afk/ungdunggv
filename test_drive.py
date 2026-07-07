import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st

def check_connection():
    try:
        # Đường dẫn tới tệp key
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('key_google.json', scope)
        client = gspread.authorize(creds)
        
        # Thử mở một file bất kỳ (thầy thay 'Tên_File_Sheets_Của_Thầy' bằng tên file thật)
        sh = client.open('Tên_File_Sheets_Của_Thầy')
        st.success("✅ Kết nối Google Sheets thành công!")
        return True
    except Exception as e:
        st.error(f"❌ Kết nối thất bại: {e}")
        return False

if st.button("Kiểm tra kết nối"):
    check_connection()