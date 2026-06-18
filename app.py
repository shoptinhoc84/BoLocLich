def parse_pdf_content(raw_text):
    # Sử dụng Regex để bóc tách từng ô dữ liệu được bọc trong dấu ngoặc kép "" 
    # Cấu trúc bảng của bạn dạng: "STT","Cơ sở đào tạo","Số lượng","Thời gian","Địa điểm"
    
    # Tìm tất cả các cụm hàng dựa trên cấu trúc các ô text liền kề
    # Tìm chuỗi có dạng "giá trị 1","giá trị 2","giá trị 3","giá trị 4","giá trị 5"
    pattern = r'"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"'
    matches = re.findall(pattern, raw_text)
    
    final_list = []
    stt = 1
    
    for match in matches:
        col_stt = match[0].strip()
        co_so = match[1].strip()
        hang_va_so_luong = match[2].strip()
        thoi_gian = match[3].strip()
        dia_diem = match[4].strip()
        
        # Bỏ qua dòng tiêu đề bảng nếu có
        if "Cơ sở đào tạo" in co_so or "Số lượng" in col_stt:
            continue
            
        # Chuẩn hóa chuỗi để tìm kiếm không phân biệt hoa thường, dấu xuống dòng
        combined_text_lower = (co_so + " " + dia_diem).lower()
        
        # ĐIỀU KIỆN LỌC: Nếu tên Cơ sở HOẶC Địa điểm thi có chứa "Nguyễn Trình"
        if "nguyễn trình" in combined_text_lower:
            
            # --- XỬ LÝ NGÀY THI ---
            # Trích xuất ngày dạng dd/mm/yyyy từ cột Thời gian
            date_match = re.search(r"(\d{2}/\d{1,2}/\d{4})", thoi_gian)
            ngay_thi = date_match.group(1) if date_match else thoi_gian.replace("\n", " ").strip()
            
            # --- XỬ LÝ HẠNG THI & SỐ LƯỢNG (Tách dòng trong ô) ---
            # Làm sạch dữ liệu xuống dòng trong ô số lượng
            parts = [p.strip() for p in hang_va_so_luong.split("\n") if p.strip()]
            
            hang_list = []
            qty_list = []
            
            for part in parts:
                if "hạng" in part.lower():
                    hang_list.append(part)
                elif part.isdigit():
                    qty_list.append(part)
                    
            # Gộp lại thành chuỗi thống nhất nếu ô đó có nhiều hạng thi/số lượng
            hang_thi = " | ".join(hang_list) if hang_list else hang_va_so_luong.replace("\n", " ").strip()
            so_luong = " | ".join(qty_list) if qty_list else "Chưa rõ"
            
            # --- CHUẨN HÓA TÊN VÀ ĐỊA ĐIỂM (Xóa bớt dấu xuống dòng lỗi định dạng) ---
            co_so_clean = co_so.replace("\n", " ").replace("- ", "").strip()
            dia_diem_clean = dia_diem.replace("\n", " ").replace(";", "").strip()
            
            # Khử ký tự lạ hệ thống tự sinh như $\hat{A}p$ thành "Ấp"
            dia_diem_clean = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', dia_diem_clean)
            dia_diem_clean = re.sub(r'\$\dot\{A\}p\$', 'Ấp', dia_diem_clean)
            
            final_list.append({
                "STT": stt,
                "Ngày thi": ngay_thi,
                "Cơ sở đào tạo": co_so_clean,
                "Hạng thi": hang_thi,
                "Số lượng (Học viên)": so_luong,
                "Địa điểm tổ chức": dia_diem_clean
            })
            stt += 1
            
    return final_list
