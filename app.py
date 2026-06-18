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

# Cấu hình giao diện chuẩn hóa - Bỏ layout="wide" để tránh kích hoạt ngầm tham số cũ
st.set_page_config(page_title="Bộ Lọc Lịch Sát Hạch Nguyễn Trình")

st.markdown("""
    <style>
    .main-title { font-size:26px; font-weight:bold; color:#1E3A8A; text-align:center; margin-bottom:5px; }
    .sub-title { font-size:16px; text-align:center; color:#4B5563; margin-bottom:20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 HỆ THỐNG LỌC LỊCH SÁT HẠCH TỰ ĐỘNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Cập nhật cấu trúc Streamlit 2026 - Tối ưu hóa xử lý bảng hành chính</div>', unsafe_allow_html=True)
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
    # Phẳng hóa văn bản để loại bỏ hoàn toàn các ký tự ngắt dòng gây treo dữ liệu bảng
    text_flat = raw_text.replace('\r\n', ' ').replace('\n', ' ')
    text_flat = text_flat.replace('"', '').replace(' , ', ' ')
    text_flat = re.sub(r'\s+', ' ', text_flat)
    
    # Định vị các mỏ neo dựa trên ngày tháng thi
    date_positions = [m.start() for m in re.finditer(r'\b\d{2}/\d{1,2}/\d{4}\b', text_flat)]
    
    if not date_positions:
        return []
        
    blocks = []
    start_pos = 0
    for pos in date_positions:
        chunk_start = max(0, start_pos - 300) if start_pos > 0 else 0
        blocks.append(text_flat[chunk_start:pos + 500])
        start_pos = pos
    blocks.append(text_flat[max(0, start_pos - 300):])
    
    final_list = []
    stt_counter = 1
    seen_dates = set() 
    
    for block in blocks:
        block_lower = block.lower()
        
        # Chỉ lọc các khối có chứa từ khóa Nguyễn Trình hoặc Nguyễn Trinh
        if "nguyễn trình" in block_lower or "nguyễn trinh" in block_lower:
            
            date_match = re.search(r'(\d{2}/\d{1,2}/\d{4})', block)
            if not date_match:
                continue
            ngay_thi = date_match.group(1)
            
            unique_key = f"{ngay_thi}_{block[:50]}"
            if unique_key in seen_dates:
                continue
            seen_dates.add(unique_key)
            
            # Trích xuất Cơ sở đào tạo
            co_so = "Trung tâm Nguyễn Trình"
            text_before_date = block[:block.find(ngay_thi)]
            
            cs_matches = re.findall(r'(Trung tâm GDNN[^\,]*|Trung tâm đào tạo[^\,]*|Công ty Cổ phần[^\,]*|Trung tâm KTNV[^\,]*)', text_before_date, re.IGNORECASE)
            if cs_matches:
                candidate = cs_matches[-1].strip()
                candidate = re.sub(r'^[-–\s\.,\|]+', '', candidate).strip()
                if len(candidate) > 10 and "hạng" not in candidate.lower():
                    co_so = candidate
            
            # Trích xuất Hạng xe
            hang_candidates = []
            hang_matches = re.findall(r'(Hạng\s+[A-Z0-9,\s\-\/]+)', block, re.IGNORECASE)
            for h in hang_matches:
                h_clean = re.sub(r'\s+', ' ', h).strip()
                if len(h_clean) < 40 and h_clean not in hang_candidates:
                    hang_candidates.append(h_clean)
            hang_xe = " | ".join(hang_candidates) if hang_candidates else "Hạng A1, A"
            
            # Trích xuất Số lượng học viên
            qty_candidates = []
            numbers = re.findall(r'\b\d{2,3}\b', block)
            for num in numbers:
                day_part = ngay_thi.split('/')[0]
                month_part = ngay_thi.split('/')[1]
                if num != day_part and num != month_part and num != "14" and num != "24":
                    if num not in qty_candidates:
                        qty_candidates.append(num)
            so_luong = " | ".join(qty_candidates) if qty_candidates else "Đang cập nhật"
            
            # Trích xuất Địa điểm tổ chức
            dia_diem = "Trung tâm Sát hạch lái xe Nguyễn Trình (Vĩnh Long)"
            addr_match = re.search(r'(Địa chỉ:[^.]*)', block, re.IGNORECASE)
            if addr_match:
                dia_diem_raw = addr_match.group(1).strip()
                dia_diem_raw = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', dia_diem_raw)
                dia_diem = re.sub(r'\s+', ' ', dia_diem_raw)
                
            co_so_clean = re.sub(r'\s+', ' ', co_so).strip()
            
            final_list.append({
                "STT": stt_counter,
                "Ngày thi": ngay_thi,
                "Cơ sở đào tạo": co_so_clean,
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
        with st.spinner("🔄 Đang bóc tách dữ liệu lịch trình..."):
            raw_text = extract_text_from_pdf(uploaded_file)
    else:
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
        
    if raw_text.strip():
        parsed_data = parse_pdf_content(raw_text)
        
        if parsed_data:
            df_display = pd.DataFrame(parsed_data).drop(columns=["STT"])
            st.success(f"🎉 Xuất sắc! Hệ thống đã lọc thành công {len(parsed_data)} lịch thi.")
            
            # Giải pháp an toàn tuyệt đối: Loại bỏ hoàn toàn tham số chiều rộng tùy biến để tránh xung đột thư viện
            st.dataframe(df_display)
            
            st.write("---")
            st.subheader("📥 TẢI FILE ĐÃ ĐỊNH DẠNG ĐẸP VỀ MÁY:")
            
            excel_bytes = export_to_excel(parsed_data)
            word_bytes = export_to_word(parsed_data)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="🟢 Tải file EXCEL (.xlsx)",
                    data=excel_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="🔵 Tải file WORD (.docx)",
                    data=word_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.warning("⚠️ Không tìm thấy thông tin lịch thi nào liên quan đến Nguyễn Trình trong file này.")
