import streamlit as st
import pandas as pd
import pypdf
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt

# Cấu hình giao diện trang web
st.set_page_config(page_title="Bộ Lọc Lịch Sát Hạch Nguyễn Trình", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size:28px; font-weight:bold; color:#1E3A8A; text-align:center; margin-bottom:5px; }
    .sub-title { font-size:18px; text-align:center; color:#4B5563; margin-bottom:20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 HỆ THỐNG LỌC LỊCH SÁT HẠCH TỰ ĐỘNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Hỗ trợ xuất file Excel & Word chuẩn định dạng, kẻ ô bắt mắt</div>', unsafe_allow_html=True)
st.write("---")

uploaded_file = st.file_uploader("Kéo và thả file PDF Thông báo lịch sát hạch vào đây:", type=["pdf", "txt"])

def extract_text_from_pdf(file):
    pdf_reader = pypdf.PdfReader(file)
    full_text = ""
    for page in pdf_reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    return full_text

def parse_pdf_content(raw_text):
    # Bước 1: Tách văn bản thành các dòng độc lập và dọn dẹp khoảng trắng thừa
    raw_lines = raw_text.split('\n')
    lines = []
    for line in raw_lines:
        cleaned = line.strip()
        # Loại bỏ các ký tự bọc ngoặc kép rác do công cụ OCR tạo ra nếu có
        cleaned = re.sub(r'^["\s,]+|["\s,]+$', '', cleaned)
        if cleaned:
            lines.append(cleaned)
            
    # Bước 2: Gom cụm các dòng dựa trên sự xuất hiện của Ngày thi (dd/mm/yyyy)
    blocks = []
    current_block = []
    
    for line in lines:
        # Kiểm tra xem dòng này có chứa ngày thi hay không
        if re.search(r'\b\d{2}/\d{1,2}/\d{4}\b', line):
            if current_block:
                blocks.append(current_block)
            current_block = [line]
        else:
            if current_block:
                current_block.append(line)
                
    if current_block:
        blocks.append(current_block)
        
    final_list = []
    stt = 1
    
    # Bước 3: Phân tích cú pháp từng khối dữ liệu đã gom cụm
    for block in blocks:
        # Hợp nhất khối văn bản thành một chuỗi thống nhất phục vụ tìm kiếm
        full_block_text = " \n ".join(block)
        full_block_lower = full_block_text.lower()
        
        # ĐIỀU KIỆN LỌC: Kiểm tra sự tồn tại của từ khóa Nguyễn Trình / Nguyễn Trinh
        if "nguyễn trình" in full_block_lower or "nguyễn trinh" in full_block_lower:
            
            # 1. Trích xuất Ngày thi chính xác
            date_match = re.search(r'(\d{2}/\d{1,2}/\d{4})', full_block_text)
            ngay_thi = date_match.group(1) if date_match else "Chưa rõ"
            
            # 2. Phân tách và làm sạch Cơ sở đào tạo & Địa điểm tổ chức
            co_so = "Trung tâm Nguyễn Trình"
            dia_diem = "Sân sát hạch Nguyễn Trình (Vĩnh Long)"
            
            # Trích xuất địa điểm từ chuỗi "Địa chỉ: ..." hoặc cụm từ khóa sân bãi
            addr_match = re.search(r'(Địa chỉ:.*)', full_block_text, re.IGNORECASE)
            if addr_match:
                raw_addr = addr_match.group(1).replace("\n", " ").strip()
                # Khử lỗi mã hóa ký tự gốc tự sinh trong văn bản PDF hành chính
                raw_addr = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', raw_addr)
                raw_addr = re.sub(r'\$\dot\{A\}p\$', 'Ấp', raw_addr)
                raw_addr = re.sub(r'\s+', ' ', raw_addr)
                
                if "trung tâm" in full_block_lower:
                    # Lấy phần mô tả trước chữ Địa chỉ để làm tên Địa điểm đầy đủ
                    prefix_text = full_block_text[:addr_match.start()].split("\n")[-1].strip()
                    if len(prefix_text) > 10:
                        dia_diem = f"{prefix_text} {raw_addr}"
                    else:
                        dia_diem = f"Trung tâm Sát hạch lái xe Nguyễn Trình - {raw_addr}"
                else:
                    dia_diem = raw_addr

            # Xác định Cơ sở đào tạo thực tế ghi trên dòng đầu
            if len(block) > 0:
                first_line = block[0]
                # Loại bỏ phần ngày tháng thi lọt vào tên trường đào tạo
                first_line_clean = re.sub(r'\b\d{2}/\d{1,2}/\d{4}\b', '', first_line).strip()
                if len(first_line_clean) > 5:
                    co_so = first_line_clean
                else:
                    co_so = block[1] if len(block) > 1 else "Trung tâm Nguyễn Trình"
            
            # Dọn dẹp dấu gạch đầu dòng dư thừa ở Tên cơ sở đào tạo
            co_so = re.sub(r'^[-\s]+', '', co_so).replace("\n", " ").strip()
            co_so = re.sub(r'\s+', ' ', co_so)
            
            # 3. Trích xuất Hạng thi & Số lượng học viên dự kiến
            hang_list = []
            qty_list = []
            
            # Tìm kiếm các cụm từ chứa chữ "Hạng"
            hang_matches = re.findall(r'(Hạng\s+[A-Z0-9,\s]+)', full_block_text, re.IGNORECASE)
            for h in hang_matches:
                h_clean = h.replace("\n", " ").strip()
                h_clean = re.sub(r'\s+', ' ', h_clean)
                if h_clean not in hang_list:
                    hang_list.append(h_clean)
                    
            # Tìm kiếm các số lượng học viên dự kiến (số có từ 2 đến 3 chữ số tách biệt)
            numbers = re.findall(r'\b\d{2,3}\b', full_block_text)
            for num in numbers:
                # Loại trừ trường hợp số lượng trùng với số ngày hoặc số tháng thi
                if ngay_thi != "Chưa rõ":
                    day_part = ngay_thi.split('/')[0]
                    month_part = ngay_thi.split('/')[1]
                    if num == day_part or num == month_part or num == "14": # Loại trừ số đường 14/9 nếu có
                        continue
                if num not in qty_list:
                    qty_list.append(num)
            
            hang_thi = " | ".join(hang_list) if hang_list else "Hạng A1, A"
            so_luong = " | ".join(qty_list) if qty_list else "Đang cập nhật"
            
            final_list.append({
                "STT": stt,
                "Ngày thi": ngay_thi,
                "Cơ sở đào tạo": co_so,
                "Hạng thi": hang_thi,
                "Số lượng (Học viên)": so_luong,
                "Địa điểm tổ chức": dia_diem
            })
            stt += 1
            
    return final_list

def export_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lịch Sát Hạch Nguyễn Trình"
    ws.views.sheetView[0].showGridLines = True
    
    headers = ["STT", "Ngày thi", "Cơ sở đào tạo", "Hạng thi", "Số lượng (Học viên)", "Địa điểm tổ chức"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        
    data_font = Font(name="Times New Roman", size=11)
    for row_idx, row_data in enumerate(data, 2):
        row_values = [row_data["STT"], row_data["Ngày thi"], row_data["Cơ sở đào tạo"], row_data["Hạng thi"], row_data["Số lượng (Học viên)"], row_data["Địa điểm tổ chức"]]
        ws.append(row_values)
        
        row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") if row_idx % 2 == 0 else PatternFill(fill_type=None)
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if row_fill.fill_type:
                cell.fill = row_fill
            
            if col_idx in [1, 2, 5]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    ws.row_dimensions[1].height = 28
    
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def export_to_word(data):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    title = doc.add_paragraph()
    run = title.add_run("THÔNG BÁO LỊCH SÁT HẠCH LÁI XE - TRUNG TÂM NGUYỄN TRÌNH")
    run.bold = True
    run.font.size = Pt(14)
    title.alignment = 1 
    
    headers = ["STT", "Ngày thi", "Cơ sở đào tạo", "Hạng thi", "Số lượng", "Địa điểm tổ chức"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].alignment = 1 
        
    for item in data:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item["STT"])
        row_cells[1].text = item["Ngày thi"]
        row_cells[2].text = item["Cơ sở đào tạo"]
        row_cells[3].text = item["Hạng thi"]
        row_cells[4].text = str(item["Số lượng (Học viên)"])
        row_cells[5].text = item["Địa điểm tổ chức"]
        
        for i in range(len(headers)):
            if i in [0, 1, 4]:
                row_cells[i].paragraphs[0].alignment = 1 
            else:
                row_cells[i].paragraphs[0].alignment = 0 
                
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    
    if file_type == "pdf":
        with st.spinner("🔄 Đang xử lý bóc tách cấu trúc dữ liệu bảng thông minh..."):
            raw_text = extract_text_from_pdf(uploaded_file)
    else:
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
        
    if raw_text.strip():
        parsed_data = parse_pdf_content(raw_text)
        
        if parsed_data:
            df_display = pd.DataFrame(parsed_data).drop(columns=["STT"])
            st.success(f"🎉 Lọc thành công {len(parsed_data)} lịch trình sát hạch liên quan tới Nguyễn Trình!")
            
            st.dataframe(df_display, use_container_width=True)
            st.write("---")
            st.subheader("📥 TẢI FILE ĐÃ ĐỊNH DẠNG ĐẸP VỀ MÁY:")
            
            excel_bytes = export_to_excel(parsed_data)
            word_bytes = export_to_word(parsed_data)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="🟢 Tải file EXCEL (.xlsx) - Đã kẻ ô & giãn cột",
                    data=excel_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh_Moi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="🔵 Tải file WORD (.docx) - Chuẩn văn bản hành chính",
                    data=word_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh_Moi.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.warning("⚠️ Không tìm thấy thông tin lịch thi nào liên quan đến Nguyễn Trình trong file này.")
