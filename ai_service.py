# ai_service.py - Bổ sung xử lý lỗi bận hệ thống 503 và 429
import streamlit as st
from google import genai
from google.genai import errors

def get_system_api_key():
    return st.secrets.get("GEMINI_API_KEY", "")

def run_ai_prompt_safe(prompt_text, preferred_model="3.5 Flash", is_admin_owner=True):
    if is_admin_owner:
        api_key_to_use = get_system_api_key()
        nguon_key = "Tài khoản hệ thống (Chính chủ)"
    else:
        api_key_ca_nhan = st.session_state.get("gv_api_key_input", "").strip()
        if api_key_ca_nhan:
            api_key_to_use = api_key_ca_nhan
            nguon_key = "Tài khoản cá nhân Giáo viên"
        else:
            return "⚠️ Bạn đang truy cập từ thiết bị thành viên. Vui lòng nhập API Key Gemini cá nhân của bạn ở mục '🔑 TRẠNG THÁI TÀI KHOẢN' tại thanh bên trái để kích hoạt quyền ra đề/soạn bài!", "error"
            
    if not api_key_to_use:
        return "⚠️ Hệ thống chưa được cấu hình API Key. Vui lòng liên hệ Admin hoặc tự cung cấp mã Key cá nhân!", "error"
    
    # 🌟 THIẾT LẬP LUỒNG DỰ PHÒNG TOÀN DIỆN CHO TẤT CẢ CÁC PHÂN HỆ
    # Gộp gemini-1.5-pro làm máy gác cổng cuối cùng cho mọi lựa chọn
    model_pool = {
        "3.1 Flash-Lite": ["gemini-2.5-flash", "gemini-1.5-pro"],
        "3.5 Flash": ["gemini-2.5-flash", "gemini-1.5-pro"],
        "3.1 Pro": ["gemini-1.5-pro", "gemini-2.5-flash"],
        "Tư duy mở rộng": ["gemini-1.5-pro", "gemini-2.5-flash"]
    }
    models_to_try = model_pool.get(preferred_model, ["gemini-2.5-flash", "gemini-1.5-pro"])
    
    last_error_message = "Không có thông tin lỗi cụ thể."
    client = genai.Client(api_key=api_key_to_use)
    
    for model_name in models_to_try:
        try:
            config_params = {}
            if preferred_model == "Tư duy mở rộng" and "pro" in model_name:
                config_params["thinking_config"] = {"thinking_budget": 2048}
            
            response = client.models.generate_content(
                model=model_name,
                contents=prompt_text,
                config=config_params if config_params else None
            )
            
            if response and response.text:
                return response.text, f"{model_name} ({nguon_key})"
            else:
                continue
                
        except errors.APIError as error:  
            last_error_message = f"Mô hình {model_name} báo lỗi API: {str(error)}"
            
            # 🌟 THUẬT TOÁN BẮT BÀI LỖI 503 (QUÁ TẢI) HOẶC 429 (HẾT HẠN MỨC) ĐỂ TỰ ĐỘNG CHUYỂN DÒNG MÁY
            if "503" in str(error) or "UNAVAILABLE" in str(error):
                st.toast(f"⏳ {model_name} đang bận cục bộ. Hệ thống tự động lùi sang dòng máy dự phòng...", icon="🔄")
            elif "429" in str(error):
                st.toast(f"⏳ {model_name} đạt giới hạn hạn mức câu hỏi của ngày. Hệ thống đang lùi dòng máy...", icon="⚠️")
            continue  
            
        except Exception as e:
            last_error_message = f"Sự cố đường truyền: {str(e)}"
            continue
            
    return f"Lỗi hệ thống: Tất cả các dòng máy dự phòng cấu hình đều không phản hồi thành công. Ghi nhận lỗi cuối cùng: {last_error_message}", "error"
