def parse_pdf_content(raw_text):
    """
    Bóc tách chính xác từng dòng theo cấu trúc bảng công văn
    Khắc phục triệt để lỗi thiếu/mất chữ ở cột Cơ sở đào tạo.
    """
    # 1. Làm sạch ký tự latex/đặc biệt
    text_clean = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', raw_text)
    
    # 2. Split văn bản dựa trên STT dạng đầu dòng (01, 02, 03,...)
    pattern = r'(?=\b\d{2}\s*\||\n\d{2}\s+)'
    items = re.split(pattern, text_clean)
    
    final_list = []
    stt_counter = 1
    
    for item in items:
        item_clean = re.sub(r'\s+', ' ', item).strip()
        
        # Chỉ lấy các mục có chứa Nguyễn Trình / Nguyễn Trinh
        if "nguyễn trình" in item_clean.lower() or "nguyễn trinh" in item_clean.lower():
            
            # Lấy Ngày thi
            date_match = re.search(r'(\d{2}/\d{1,2}/\d{4})', item_clean)
            if not date_match:
                continue
            ngay_thi = date_match.group(1)
            
            # ----------------------------------------------------
            # XỬ LÝ CHÍNH XÁC CƠ SỞ ĐÀO TẠO (TRÁNH BỊ MẤT/THIẾU CHỮ)
            # ----------------------------------------------------
            co_so = ""
            
            # Trường hợp 1: Có dấu phân cách |
            if "|" in item_clean:
                parts = [p.strip() for p in item_clean.split('|')]
                # Loại bỏ số STT ở đầu phần thứ nhất (nếu có)
                part_cs = re.sub(r'^\d{2}\s*', '', parts[0]).strip()
                if len(part_cs) > 3:
                    co_so = part_cs
            
            # Trường hợp 2: Nếu chia theo | bị trích xuất thiếu hoặc không có |
            # Bắt trọn vẹn từ các từ khóa bắt đầu cơ sở đào tạo cho tới chữ "Hạng"
            if not co_so or len(co_so) < 10:
                match_cs = re.search(
                    r'((?:Trung tâm|Công ty|Trường|Khoa)\s+[^Hạng]+?)(?=\s+Hạng|\s+\|\s*Hạng|\s+\d{2,4}\b)', 
                    item_clean, 
                    re.IGNORECASE
                )
                if match_cs:
                    co_so = match_cs.group(1).strip()
            
            # Làm sạch các ký tự rác nếu còn sót ở đầu/cuối tên cơ sở
            co_so = re.sub(r'^\d{2}\s*[-–\.\,\|]*\s*', '', co_so).strip()
            co_so = re.sub(r'\s+', ' ', co_so)
            
            if not co_so:
                co_so = "Trung tâm GDNN và SHLX Nguyễn Trình"

            # ----------------------------------------------------
            # XỬ LÝ HẠNG THI
            # ----------------------------------------------------
            hang_match = re.search(r'(Hạng\s+[A-Z0-9,\s\-\/]+?)(?=\s+\d{2,4}\b|\s*\||\s*\d{2}/\d{1,2}/\d{4})', item_clean, re.IGNORECASE)
            hang_xe = hang_match.group(1).strip() if hang_match else "Hạng A1, A"
            hang_xe = re.sub(r'\s+', ' ', hang_xe)

            # ----------------------------------------------------
            # XỬ LÝ SỐ LƯỢNG HỌC VIÊN
            # ----------------------------------------------------
            # Tìm số có 2-4 chữ số nằm kề trước ngày thi hoặc ở cột số lượng
            qty_match = re.search(r'\b(\d{2,4})\b(?=\s*\|?\s*\d{2}/\d{1,2}/\d{4})', item_clean)
            if not qty_match:
                qty_match = re.search(r'\b(\d{2,4})\b', item_clean)
            
            so_luong = qty_match.group(1) if qty_match else "Đang cập nhật"

            # ----------------------------------------------------
            # XỬ LÝ ĐỊA ĐIỂM TỔ CHỨC
            # ----------------------------------------------------
            dia_diem = ""
            if "Địa chỉ:" in item_clean:
                dia_diem = item_clean[item_clean.find("Địa chỉ:"):].strip()
            elif "|" in item_clean:
                parts = [p.strip() for p in item_clean.split('|')]
                dia_diem = parts[-1]
            
            if len(dia_diem) < 15:
                dia_diem = "Trung tâm Giáo dục nghề nghiệp và Sát hạch lái xe Nguyễn Trình (Ấp Giồng Trôm, xã Châu Thành, Vĩnh Long)"
            else:
                dia_diem = re.sub(r'\s+', ' ', dia_diem).strip()

            # Đưa vào danh sách kết quả
            final_list.append({
                "STT": stt_counter,
                "Ngày thi": ngay_thi,
                "Cơ sở đào tạo": co_so,
                "Hạng thi": hang_xe,
                "Số lượng (Học viên)": so_luong,
                "Địa điểm tổ chức": dia_diem
            })
            stt_counter += 1
            
    return final_list
