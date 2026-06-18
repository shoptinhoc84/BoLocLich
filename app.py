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
st.markdown('<div class="sub-title">Thuật toán quét dòng thông minh - Chống treo ứng dụng</div>', unsafe_allow_html=True)
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
    # --- BƯỚC 1: LÀM SẠCH VÀ CHUẨN HÓA TOÀN BỘ VĂN BẢN VỀ DẠNG DÒNG THUẦN ---
    raw_lines = raw_text.split('\n')
    cleaned_lines = []
    
    for line in raw_lines:
        # Khử sạch các dấu ngoặc kép rác, dấu phẩy phân tách ô bảng lỗi của OCR
        line_clean = line.replace('"', '').replace(' , ', ' ').strip()
        line_clean = re.sub(r'^[\s,\|]+|[\s,\|]+$', '', line_clean)
        if line_clean:
            cleaned_lines.append(line_clean)
            
    final_list = []
    stt_counter = 1
    n = len(cleaned_lines)
    
    # --- BƯỚC 2: THUẬT TOÁN QUÈT MỎ NEO (ANCHOR SCANNING) THEO NGÀY THI ---
    for i in range(n):
        current_line = cleaned_lines[i]
        
        # Tìm dòng chứa ngày thi dạng dd/mm/yyyy hoặc dd/m/yyyy
        date_match = re.search(r'\b(\d{2}/\d{1,2}/\d{4})\b', current_line)
        
        if date_match:
            ngay_thi = date_match.group(1)
            
            # Khởi tạo vùng tìm kiếm thông tin xung quanh ngày thi
            co_so_candidates = []
            hang_candidates = []
            qty_candidates = []
            dia_diem_text = ""
            
            # 1. Tìm Cơ sở đào tạo bằng cách quét ngược từ dòng ngày thi lên trên (Tối đa 4 dòng)
            for j in range(max(0, i-4), i):
                line_upper = cleaned_lines[j]
                if any(k in line_upper.lower() for k in ["trung tâm", "công ty", "trường cao đẳng", "trường đhsp"]):
                    co_so_candidates.append(line_upper)
            
            # Chọn cơ sở đào tạo gần dòng ngày thi nhất ở phía trên
            co_so_dao_tao = co_so_candidates[-1] if co_so_candidates else "Trung tâm Nguyễn Trình"
            
            # 2. Tìm Địa điểm và Hạng/Số lượng bằng cách quét xuôi xuống dưới từ dòng ngày thi (Tối đa 4 dòng)
            for j in range(i, min(n, i+5)):
                line_lower = cleaned_lines[j]
                
                # Tìm kiếm chuỗi Địa chỉ / Địa điểm
                if "địa chỉ:" in line_lower.lower() or "trung tâm giáo dục" in line_lower.lower() or "sân tập" in line_lower.lower():
                    if not dia_diem_text:
                        dia_diem_text = line_lower
                    else:
                        dia_diem_text += " " + line_lower
                        
                # Tìm kiếm thông tin Hạng xe
                hang_match = re.findall(r'(Hạng\s+[A-Z0-9,\s\-\/]+)', line_lower, re.IGNORECASE)
                for h in hang_match:
                    if h.strip() not in hang_candidates:
                        hang_candidates.append(h.strip())
                        
                # Tìm số lượng học viên (Số nguyên tách biệt từ 2-3 chữ số)
                numbers = re.findall(r'\b\d{2 organiz,3}\b|\b\d{2,3}\b', line_lower)
                for num in numbers:
                    # Bỏ qua nếu số đó trùng với ngày hoặc tháng thi
                    day_part = ngay_thi.split('/')[0]
                    month_part = ngay_thi.split('/')[1]
                    if num != day_part and num != month_part and num != "14": # bỏ qua số đường 14/9
                        if num not in qty_candidates:
                            qty_candidates.append(num)

            # --- BƯỚC 3: BỘ LỌC ĐIỀU KIỆN RIÊNG CHO NGUYỄN TRÌNH ---
            combined_context = (co_so_dao_tao + " " + dia_diem_text).lower()
            if "nguyễn trình" in combined_context or "nguyễn trinh" in combined_context:
                
                # Làm sạch dữ liệu đầu ra văn bản hành chính
                co_so_final = re.sub(r'^[-\s\.]+', '', co_so_dao_tao).strip()
                co_so_final = re.sub(r'\s+', ' ', co_so_final)
                
                hang_final = " | ".join(hang_candidates) if hang_candidates else "Hạng A1"
                so_luong_final = " | ".join(qty_candidates) if qty_candidates else "Đang cập nhật"
                
                if not dia_diem_text:
                    dia_diem_text = "Trung tâm Sát hạch lái xe Nguyễn Trình (Vĩnh Long)"
                
                dia_diem_final = re.sub(r'\s+', ' ', dia_diem_text).strip()
                # Khử lỗi ký tự hiển thị mã Latinh tự sinh từ tệp hành chính
                dia_diem_final = re.sub(r'\$\\hat\{A\}p\$', 'Ấp', dia_diem_final)
                dia_diem_final = re.sub(r'\$\dot\{A\}p\$', 'Ấp', dia_diem_final)
                
                final_list.append({
                    "STT": stt_counter,
                    "Ngày thi": ngay_thi,
                    "Cơ sở đào tạo": co_so_final,
                    "Hạng thi": hang_final,
                    "Số lượng (Học viên)": so_luong_final,
                    "Địa điểm tổ chức": dia_diem_final
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
        with st.spinner("🔄 Đang chạy thuật toán quét mỏ neo siêu tốc..."):
            raw_text = extract_text_from_pdf(uploaded_file)
    else:
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
        
    if raw_text.strip():
        parsed_data = parse_pdf_content(raw_text)
        
        if parsed_data:
            df_display = pd.DataFrame(parsed_data).drop(columns=["STT"])
            st.success(f"🎉 Thuật toán mới đã lọc thành công {len(parsed_data)} lịch thi của Nguyễn Trình!")
            
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
