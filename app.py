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

# Cấu hình giao diện chuẩn hóa Streamlit
st.set_page_config(page_title="Bộ Lọc Lịch Sát Hạch Nguyễn Trình", layout="centered")

st.markdown("""
    <style>
    .main-title { font-size:26px; font-weight:bold; color:#1E3A8A; text-align:center; margin-bottom:5px; }
    .sub-title { font-size:16px; text-align:center; color:#4B5563; margin-bottom:20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 HỆ THỐNG LỌC LỊCH SÁT HẠCH TỰ ĐỘNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Tối ưu hóa bóc tách bảng lịch sát hạch Công an tỉnh Vĩnh Long</div>', unsafe_allow_html=True)
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
            # XỬ LÝ CHÍNH XÁC CƠ SỞ ĐÀO TẠO (KHÔNG BỊ THIẾU CHỮ)
            # ----------------------------------------------------
            co_so = ""
            
            # Trường hợp 1: Có dấu phân cách |
            if "|" in item_clean:
                parts = [p.strip() for p in item_clean.split('|')]
                part_cs = re.sub(r'^\d{2}\s*', '', parts[0]).strip()
                if len(part_cs) > 3:
                    co_so = part_cs
            
            # Trường hợp 2: Nếu chia theo | bị trích xuất thiếu hoặc không có |
            if not co_so or len(co_so) < 10:
                match_cs = re.search(
                    r'((?:Trung tâm|Công ty|Trường|Khoa)\s+[^Hạng]+?)(?=\s+Hạng|\s+\|\s*Hạng|\s+\d{2,4}\b)', 
                    item_clean, 
                    re.IGNORECASE
                )
                if match_cs:
                    co_so = match_cs.group(1).strip()
            
            # Làm sạch các ký tự rác ở đầu/cuối tên cơ sở
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

# Luồng xử lý chính của Streamlit
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
            st.success(f"🎉 Lọc thành công {len(parsed_data)} lịch thi liên quan đến Nguyễn Trình!")
            
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
