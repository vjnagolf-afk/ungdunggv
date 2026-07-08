import streamlit as st
import pandas as pd
import re

def render_tkb_manager():
    st.header("📅 QUẢN LÝ THỜI KHÓA BIỂU")
    uploaded_tkb = st.file_uploader("Tải lên file TKB (.xlsx)", type=["xlsx"])
    
    if uploaded_tkb:
        try:
            # 1. Đọc thô file để tự động tìm dòng tiêu đề chứa chữ THỨ và TIẾT
            df_raw = pd.read_excel(uploaded_tkb, header=None)
            header_idx = 4  
            for idx, row in df_raw.iterrows():
                row_str = [str(x).upper() for x in row.values if pd.notna(x)]
                if "THỨ" in row_str and "TIẾT" in row_str:
                    header_idx = idx
                    break
            
            # 2. Đọc lại file Excel từ dòng tiêu đề tìm được
            df = pd.read_excel(uploaded_tkb, header=header_idx)
            df.columns = [str(c).strip() for c in df.columns]
            
            # --- THUẬT TOÁN ĐIỀN KHUYẾT VÀ CHUẨN HÓA CỘT THỨ ---
            if "THỨ" in df.columns:
                df["THỨ"] = df["THỨ"].ffill()
                df["THỨ"] = df["THỨ"].apply(lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x).strip())
                
            # Loại bỏ các giá trị rác hệ thống phát sinh
            df = df.fillna("")
            df = df.astype(str).replace(["None", "nan", "NaN"], "")
            
            # Lấy danh sách các cột Lớp học (ví dụ: '6A (Hiếu)', '6B (Duy)',...)
            class_columns = [c for c in df.columns if "Unnamed" not in c and c.upper() not in ["THỨ", "TIẾT", "STT", "CỘT 1", "CỘT 2"]]
            
            # --- THUẬT TOÁN RÚT TRÍCH DANH SÁCH TÊN GIÁO VIÊN BỘ MÔN ---
            all_teachers = set()
            for col in class_columns:
                for cell_value in df[col].values:
                    cell_str = str(cell_value).strip()
                    if cell_str and "-" in cell_str:
                        parts = cell_str.split("-")
                        if len(parts) >= 2:
                            teacher_name = parts[1].strip()
                            if teacher_name: all_teachers.add(teacher_name)
                            
            # --- ĐIỀU HƯỚNG GIAO DIỆN TABS ---
            tab1, tab2 = st.tabs(["📊 Thời khóa biểu chung", "👤 TKB theo giáo viên"])
            
            with tab1:
                st.markdown("##### 📋 Bảng xem trước Thời khóa biểu toàn trường")
                st.dataframe(df, use_container_width=True, hide_index=True)
                
            with tab2:
                if not all_teachers:
                    st.error("Không tìm thấy thông tin giáo viên bộ môn trong phân phối lịch dạy. Vui lòng kiểm tra lại file.")
                else:
                    selected_teacher = st.selectbox("👤 Chọn tên Giáo viên cần xem thời khóa biểu cá nhân:", sorted(list(all_teachers)))
                    
                    # --- THUẬT TOÁN ĐỈNH CAO: DỰNG LƯỚI TKB MA TRẬN 5 TIẾT x 6 THỨ (GIỐNG ẢNH MẪU) ---
                    # Khởi tạo bảng khung trống hoàn toàn từ Thứ 2 đến Thứ 7
                    days_list = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"]
                    slots_list = ["1", "2", "3", "4", "5"]
                    
                    # Dựng ma trận rỗng bằng từ điển (Dictionary)
                    matrix_data = {day: ["" for _ in slots_list] for day in days_list}
                    
                    # Quét qua bảng TKB tổng để bốc dữ liệu điền vào ma trận lưới
                    for _, row in df.iterrows():
                        thu_raw = row.get("THỨ", "").strip()
                        tiet_raw = str(row.get("TIẾT", "")).strip()
                        
                        # Chuẩn hóa tên thứ để khớp với khung ma trận
                        thu_clean = f"Thứ {thu_raw}" if thu_raw.isdigit() else thu_raw
                        if "2" in thu_clean: thu_clean = "Thứ 2"
                        elif "3" in thu_clean: thu_clean = "Thứ 3"
                        elif "4" in thu_clean: thu_clean = "Thứ 4"
                        elif "5" in thu_clean: thu_clean = "Thứ 5"
                        elif "6" in thu_clean: thu_clean = "Thứ 6"
                        elif "7" in thu_clean: thu_clean = "Thứ 7"
                        
                        # Chỉ xử lý nếu khớp đúng Tiết từ 1 đến 5 và Thứ từ 2 đến 7
                        if thu_clean in matrix_data and tiet_raw in slots_list:
                            tiet_idx = slots_list.index(tiet_raw)
                            
                            # Quét qua từng lớp của tiết học đó xem giáo viên được chọn dạy lớp nào
                            cell_lessons = []
                            for col_class in class_columns:
                                cell_content = str(row[col_class]).strip()
                                if "-" in cell_content:
                                    parts = cell_content.split("-")
                                    if len(parts) >= 2 and parts[1].strip() == selected_teacher:
                                        # Trích xuất môn học và tên lớp thu gọn (ví dụ: '6A (Hiếu)' -> lấy '6A')
                                        sub_name = parts[0].strip()
                                        class_short = col_class.split("(")[0].strip()
                                        
                                        # Gom lại thành chuỗi 'Môn - Lớp' (Ví dụ: T.Anh - 9E) đúng như ảnh mẫu của thầy Bình
                                        cell_lessons.append(f"{sub_name} - {class_short}")
                            
                            # Nếu tiết đó giáo viên có dạy (có thể dạy 1 lớp hoặc song hành), điền vào ô ma trận
                            if cell_lessons:
                                matrix_data[thu_clean][tiet_idx] = " / ".join(cell_lessons)
                                
                    # Chuyển đổi Ma trận từ điển thành DataFrame để Streamlit hiển thị
                    df_matrix = pd.DataFrame(matrix_data)
                    # Chèn thêm cột "TIẾT" vào đầu bảng để thầy cô dễ định vị dòng
                    df_matrix.insert(0, "TIẾT", slots_list)
                    
                    # Hiển thị tiêu đề tên giáo viên in đậm to rõ
                    st.markdown(f"### 🗂️ Thời khóa biểu giảng dạy: Thầy/Cô **{selected_teacher}**")
                    
                    # Đổ dữ liệu ra lưới hiển thị trực quan giống hệt hình vẽ mẫu Excel
                    st.dataframe(df_matrix, use_container_width=True, hide_index=True)
                    
                    # Tính năng cộng điểm: Cho phép tải TKB cá nhân này về dạng file Excel thu nhỏ
                    output_personal = io.BytesIO()
                    with pd.ExcelWriter(output_personal, engine='openpyxl') as writer:
                        df_matrix.to_excel(writer, index=False, sheet_name=f"TKB_{selected_teacher}")
                    st.download_button(
                        label=f"📥 Tải lịch dạy Thầy/Cô {selected_teacher} (.xlsx)",
                        data=output_personal.getvalue(),
                        file_name=f"TKB_{selected_teacher}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                        
        except Exception as e:
            st.error(f"Lỗi khi xử lý ma trận lưới TKB: {e}")
    else:
        st.info("💡 Vui lòng tải file Excel Thời khóa biểu (.xlsx) lên.")
