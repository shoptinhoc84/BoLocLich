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
st.markdown('<div class="sub-title">Thuật toán tích lũy khối dữ liệu - Khắc phục triệt để treo Web</div>', unsafe_allow_html=True)
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
    # Tách văn bản thành danh sách dòng và dọn dẹp ký tự thừa
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
    
    parsed_blocks = []
    current_block = []
    
    # --- BƯỚC 1: GOM CỤM DỮ LIỆU THEO PHƯƠNG PHÁP TÍCH LŨY ---
    # Cứ mỗi khi xuất hiện một dòng có chứa Ngày Thi (dd/mm/yyyy), ta coi đó là điểm kết thúc hoặc khởi đầu của 1 khối.
    for line in lines:
        cleaned_line = line.replace('"', '').replace(' , ', ' ').strip()
        if not cleaned_line:
            continue
            
        # Kiểm tra xem dòng có chứa ngày thi không
        if re.search(r'\b\d{2}/\d{1,2}/\d{4}\b', cleaned_line):
            if current_block:
                parsed_blocks.append(current_block)
            current_block = [cleaned_line]
        else:
            current_block.append(cleaned_line)
            
    if current_block:
        parsed_blocks.append(current_block)
        
    final_list = []
    stt_counter = 1
    
    # --- BƯỚC 2: TRÍCH XUẤT THÔNG TIN TRÊN TỪNG KHỐI DÒNG DỮ LIỆU ---
    for block in parsed_blocks:
        # Gộp toàn bộ văn bản của khối hiện tại để phân tích ngữ cảnh
        block_text = " ".join(block)
        block_lower = block_text.lower()
        
        # BỘ LỌC: Chỉ xử lý nếu khối có liên quan đến từ khóa "Nguyễn Trình" hoặc "Nguyễn Trinh"
        if "nguyễn trình" in block_lower or "nguyễn trinh" in block_lower:
            
            # 1. Trích xuất Ngày thi
            date_match = re.search(r'(\d{2}/\d{1,2}/\d{4})', block_text)
            ngay_thi = date_match.group(1) if date_match else "Chưa rõ"
            
            # 2. Phân tích bóc tách Cơ sở đào tạo
            co_so = "Trung tâm Nguyễn Trình"
            # Tìm dòng chứa từ khóa nhận diện cơ sở đào tạo trong khối
            for b_line in block:
                b_line_lower = b_line.lower()
                if any(k in b_line_lower for k in ["trung tâm gdnn", "trung tâm đào tạo", "công ty", "trường cao đẳng"]):
                    # Loại bỏ các số STT hoặc ngày tháng dính vào dòng tên cơ sở
                    clean_cs = re.sub(r'\b\d{2}/\d{1,2}/\d{4}\b', '', b_line).strip()
                    clean_cs = re.sub(r'^\d+\s*', '', clean_cs).strip()
                    clean_cs = re.sub(r'^[-\s\.,\|]+', '', clean_cs).strip()
                    if len(clean_cs) > 5 and "hạng" not in clean_cs.lower() and "địa chỉ" not in clean_cs.lower():
                        co_so = clean_cs
                        break
            
            # 3. Trích xuất thông tin Hạng Xe
            hang_candidates = []
            # Tìm cấu trúc "Hạng X, Y, Z"
            hang_matches = re.findall(r'(Hạng\s+[A-Z0-9,\s\-\/]+)', block_text, re.IGNORECASE)
            for h in hang_matches:
                h_clean = re.sub(r'\s+', ' ', h).strip()
                if h_clean and h_clean not in hang_candidates:
                    hang_candidates.append(h_clean)
            hang_xe = " | ".join(hang_candidates) if hang_candidates else "Hạng A1, A"
            
            # 4. Trích xuất Số lượng học viên
            qty_candidates = []
            # Tìm tất cả các số có từ 2 đến 3 chữ số đứng tách biệt trong khối
            numbers = re.findall(r'\b\d{2,3}\b', block_text)
            for num in numbers:
                # Loại trừ số trùng khớp với Ngày hoặc Tháng thi, hoặc số hiệu đường xe chạy (ví dụ đường 14/9)
                if ngay_thi != "Chưa rõ":
                    date_parts = ngay_thi.split('/')
                    if num == date_parts[0] or num == date_parts[1] or num == "14":
                        continue
                if num not in qty_candidates:
                    qty_candidates.append(num)
            so_luong = " | ".join(qty_candidates) if qty_candidates else "Đang cập nhật"
            
            # 5. Trích xuất Địa điểm tổ chức sát hạch
            dia_diem = "Trung tâm Sát hạch lái xe Nguyễn Trình (Vĩnh Long)"
            addr_match = re.search(r'(Địa chỉ:.*)', block_text, re.IGNORECASE)
            if addr_match:
                dia_diem = addr_match.group(1).strip()
                # Khử lỗi ký tự hiển thị Latinh cổ tự sinh khi kết xuất tệp PDF hành chính
                dia_diem = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', dia_diem)
                dia_diem = re.sub(r'\$\dot\{A\}p\$', 'Ấp', dia_diem)
                dia_diem = re.sub(r'\s+', ' ', dia_diem)
            
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
        with st.spinner("🔄 Đang bóc tách dữ liệu khối siêu tốc..."):
            raw_text = extract_text_from_pdf(uploaded_file)
    else:
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
        
    if raw_text.strip():
        parsed_data = parse_pdf_content(raw_text)
        
        if parsed_data:
            df_display = pd.DataFrame(parsed_data).drop(columns=["STT"])
            st.success(f"🎉 Đã lọc thành công {len(parsed_data)} lịch thi liên quan tới Nguyễn Trình!")
            
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
                    label="🔵 Tải file WORD (.docx) - Chuẩn văn bản in ấn",
                    data=word_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh_Moi.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.warning("⚠️ Không tìm thấy thông tin lịch thi nào liên quan đến Nguyễn Trình trong file này.")
