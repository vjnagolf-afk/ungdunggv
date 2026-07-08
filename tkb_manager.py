import streamlit as st
import pandas as pd

def render_tkb_manager():
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU")
    uploaded_tkb = st.file_uploader("Tải lên file TKB (.xlsx)", type=["xlsx"])
    
    if uploaded_tkb:
        try:
            # 1. Đọc thô file để tự động tìm dòng tiêu đề chuẩn (Tránh cố định header=3 bị lệch)
            df_raw = pd.read_excel(uploaded_tkb, header=None)
            header_idx = 3 # Giá trị mặc định phòng hờ
            for idx, row in df_raw.iterrows():
                row_str = [str(x).upper() for x in row.values if pd.notna(x)]
                if "THỨ" in row_str and "TIẾT" in row_str:
                    header_idx = idx
                    break
            
            # 2. Đọc lại file Excel từ dòng tiêu đề tìm được
            df = pd.read_excel(uploaded_tkb, header=header_idx)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Sửa lỗi gộp ô: Điền đầy đủ dữ liệu cho cột "THỨ" bị trống (Thứ 2, Thứ 3...)
            if "THỨ" in df.columns:
                df["THỨ"] = df["THỨ"].ffill()
                
            # Sửa lỗi hiển thị: Thay thế tất cả giá trị trống, None, nan thành chuỗi rỗng để sạch bảng
            df = df.fillna("")
            df = df.astype(str).replace(["None", "nan", "NaN"], "")
            
            # Lấy danh sách tên Giáo viên (Loại bỏ các cột định vị hệ thống)
            teachers = [c for c in df.columns if "Unnamed" not in c and c.upper() not in ["THỨ", "TIẾT", "STT", "CỘT 1", "CỘT 2"]]
            
            # --- ĐIỀU HƯỚNG GIAO DIỆN TABS ---
            tab1, tab2 = st.tabs(["📊 Thời khóa biểu chung", "👤 TKB theo giáo viên"])
            
            with tab1:
                # Hiển thị bảng sạch sẽ hoàn toàn không còn chữ Unnamed hay None
                st.dataframe(df, use_container_width=True, hide_index=True)
                
            with tab2:
                if not teachers:
                    st.error("Không tìm thấy tên giáo viên trong file. Vui lòng kiểm tra lại cấu trúc file.")
                else:
                    selected_teacher = st.selectbox("👤 Chọn giáo viên cần tra cứu:", sorted(teachers))
                    
                    # Trích xuất lịch dạy của riêng giáo viên được chọn
                    cols_filter = []
                    if "THỨ" in df.columns: cols_filter.append("THỨ")
                    if "TIẾT" in df.columns: cols_filter.append("TIẾT")
                    cols_filter.append(selected_teacher)
                    
                    tkb_gv = df[cols_filter].copy()
                    tkb_gv.columns = ["Thứ", "Tiết", "Lịch dạy"]
                    
                    # Thuật toán lọc: Chỉ giữ lại các hàng thực sự có tiết dạy (bỏ các hàng trống)
                    tkb_gv_clean = tkb_gv[tkb_gv["Lịch dạy"].str.strip() != ""]
                    
                    if not tkb_gv_clean.empty:
                        st.success(f"📋 Lịch giảng dạy trong tuần của thầy/cô: **{selected_teacher}**")
                        st.dataframe(tkb_gv_clean, use_container_width=True, hide_index=True)
                    else:
                        st.info(f"ℹ️ Thầy/cô **{selected_teacher}** không có tiết giảng dạy nào trong tuần này.")
                        
        except Exception as e:
            st.error(f"Lỗi khi đọc file TKB: {e}")
    else:
        st.info("💡 Vui lòng tải file Excel Thời khóa biểu (.xlsx) lên.")
