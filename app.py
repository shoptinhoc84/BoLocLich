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

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Bộ Lọc Lịch Sát Hạch Nguyễn Trình", layout="centered")

st.markdown("""
    <style>
    .main-title { font-size:26px; font-weight:bold; color:#1E3A8A; text-align:center; margin-bottom:5px; }
    .sub-title { font-size:16px; text-align:center; color:#4B5563; margin-bottom:20px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 HỆ THỐNG LỌC LỊCH SÁT HẠCH TỰ ĐỘNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Bóc tách dữ liệu chuẩn xác - Xuất file Excel & Word chuyên nghiệp</div>', unsafe_allow_html=True)
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
    # 1. Làm sạch ký tự mã hóa lỗi từ PDF
    text_clean = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', raw_text)
    
    # 2. Định vị vị trí bắt đầu của các Số Thứ Tự trong bảng (01, 02, ..., 29)
    stt_matches = list(re.finditer(r'(?:^|\n)\s*(\d{2})\s*(?=\||\s+[T|C|K|P|H|Q|X])', text_clean))
    
    blocks = []
    for i in range(len(stt_matches)):
        start_pos = stt_matches[i].start()
        end_pos = stt_matches[i+1].start() if i + 1 < len(stt_matches) else len(text_clean)
        block_text = text_clean[start_pos:end_pos].strip()
        blocks.append(block_text)

    final_list = []
    dia_diem_chuan = "Trung tâm Giáo dục nghề nghiệp và Sát hạch lái xe Nguyễn Trình (Ấp Giồng Trôm, xã Châu Thành, tỉnh Vĩnh Long)"

    for block in blocks:
        # Tìm Ngày thi trong khối
        date_match = re.search(r'(\d{2}/\d{1,2}/\d{4})', block)
        if not date_match:
            continue
        ngay_thi = date_match.group(1)
        
        # ---------------------------------------------------------------------
        # 1. KIỂM TRA ĐỊA ĐIỂM TỔ CHỨC (BẮT BUỘC PHẢI LÀ NGUYỄN TRÌNH)
        # ---------------------------------------------------------------------
        idx_addr = block.find("Địa chỉ:")
        if idx_addr != -1:
            site_text = block[idx_addr:]
        elif "|" in block:
            site_text = block.split('|')[-1]
        else:
            site_text = block

        is_nguyen_trinh_site = (
            "nguyễn trình" in site_text.lower() or 
            "nguyễn trinh" in site_text.lower() or 
            "giồng trôm" in site_text.lower()
        )
        
        if not is_nguyen_trinh_site:
            continue

        # ---------------------------------------------------------------------
        # 2. BÓC TÁCH CƠ SỞ ĐÀO TẠO
        # ---------------------------------------------------------------------
        idx_date = block.find(ngay_thi)
        text_before_date = block[:idx_date] if idx_date != -1 else block
        
        entity_patterns = [
            r'Trung tâm GDNN[^\n|]+',
            r'Trung tâm đào tạo[^\n|]+',
            r'Trung tâm quản lý[^\n|]+',
            r'Công ty Cổ phần[^\n|]+',
            r'Công ty TNHH[^\n|]+',
            r'Trường Cao đẳng[^\n|]+',
            r'Khoa giao thông[^\n|]+',
            r'Quá hạn GPLX[^\n|]*',
            r'Bến xe[^\n|]*'
        ]
        
        cs_found = []
        for pat in entity_patterns:
            matches = re.findall(pat, text_before_date, re.IGNORECASE)
            for m in matches:
                m_clean = re.sub(r'^\s*[-–\.\,\|]*\s*', '', m).strip()
                m_clean = re.sub(r'\s+Hạng.*$', '', m_clean, flags=re.IGNORECASE)
                m_clean = re.sub(r'\s+\d{2,4}\b.*$', '', m_clean)
                if len(m_clean) > 3 and "địa chỉ" not in m_clean.lower() and m_clean not in cs_found:
                    cs_found.append(m_clean)
        
        if not cs_found:
            if "|" in text_before_date:
                parts = [p.strip() for p in text_before_date.split('|')]
                for p in parts[:2]:
                    p_clean = re.sub(r'^\s*\d{2}\s*', '', p).strip()
                    if len(p_clean) > 3 and "hạng" not in p_clean.lower():
                        cs_found.append(p_clean)

        co_so_final = " - ".join(cs_found)
        co_so_final = re.sub(r'\s+', ' ', co_so_final).strip(" -")
        
        if not co_so_final:
            co_so_final = "Trung tâm GDNN và SHLX Nguyễn Trình"

        # ---------------------------------------------------------------------
        # 3. TRÍCH XUẤT SỐ LƯỢNG HỌC VIÊN
        # ---------------------------------------------------------------------
        day_str = ngay_thi.split('/')[0]
        month_str = ngay_thi.split('/')[1]
        
        numbers_found = re.findall(r'\b\d{2,4}\b', block)
        so_luong = "Chưa rõ"
        
        for num in numbers_found:
            if num not in ["2026", day_str, month_str, "14", "24"]:
                if int(num) >= 50:
                    so_luong = num
                    break

        # ---------------------------------------------------------------------
        # 4. TRÍCH XUẤT HẠNG THI
        # ---------------------------------------------------------------------
        hang_match = re.search(r'(Hạng\s+[A-Z0-9,\s\-\/]+?)(?=\s+\d{2,4}\b|\s*\||\s*\d{2}/\d{1,2}/\d{4})', block, re.IGNORECASE)
        hang_xe = hang_match.group(1).strip() if hang_match else "Hạng A1, A"
        hang_xe = re.sub(r'\s+', ' ', hang_xe)

        # Trả về kết quả khớp 100% với các tiêu đề cột xuất file
        final_list.append({
            "Cơ sở đào tạo": co_so_final,
            "Ngày thi": ngay_thi,
            "Hạng thi": hang_xe,
            "Số lượng học viên": so_luong,
            "Địa điểm tổ chức": dia_diem_chuan
        })

    return final_list

def export_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Lịch Sát Hạch Nguyễn Trình"
    ws.views.sheetView[0].showGridLines = True
    
    headers = ["Cơ sở đào tạo", "Ngày thi", "Hạng thi", "Số lượng học viên", "Địa điểm tổ chức"]
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
        row_values = [
            row_data["Cơ sở đào tạo"], 
            row_data["Ngày thi"], 
            row_data["Hạng thi"], 
            row_data["Số lượng học viên"], 
            row_data["Địa điểm tổ chức"]
        ]
        ws.append(row_values)
        
        row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") if row_idx % 2 == 0 else PatternFill(fill_type=None)
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if row_fill.fill_type:
                cell.fill = row_fill
            
            if col_idx in [2, 4]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    ws.row_dimensions[1].height = 28
    
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                val_str = str(cell.value)
                max_line = max([len(line) for line in val_str.split('\n')]) if '\n' in val_str else len(val_str)
                max_len = max(max_len, max_line)
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
    run = title.add_run("THÔNG BÁO LỊCH SÁT HẠCH LÁI XE TẠI TRUNG TÂM NGUYỄN TRÌNH")
    run.bold = True
    run.font.size = Pt(14)
    title.alignment = 1 
    
    headers = ["Cơ sở đào tạo", "Ngày thi", "Hạng thi", "Số lượng học viên", "Địa điểm tổ chức"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].alignment = 1 
        
    for item in data:
        row_cells = table.add_row().cells
        row_cells[0].text = item["Cơ sở đào tạo"]
        row_cells[1].text = item["Ngày thi"]
        row_cells[2].text = item["Hạng thi"]
        row_cells[3].text = str(item["Số lượng học viên"])
        row_cells[4].text = item["Địa điểm tổ chức"]
        
        for i in range(len(headers)):
            if i in [1, 3]:
                row_cells[i].paragraphs[0].alignment = 1 
            else:
                row_cells[i].paragraphs[0].alignment = 0 
                
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# Luồng xử lý chính Streamlit
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
            df_display = pd.DataFrame(parsed_data)
            st.success(f"🎉 Lọc thành công {len(parsed_data)} lịch thi tổ chức tại địa điểm Nguyễn Trình!")
            
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
                    file_name="Lich_Thi_Tai_Nguyen_Trinh.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="🔵 Tải file WORD (.docx)",
                    data=word_bytes,
                    file_name="Lich_Thi_Tai_Nguyen_Trinh.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.warning("⚠️ Không tìm thấy lịch thi nào tổ chức tại địa điểm Nguyễn Trình.")
