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
    lines = [line.strip() for line in raw_text.split("\n") if line.strip()]
    data_rows = []
    current_row = None
    
    # Gom cụm các dòng: Mỗi khi gặp một ngày thi mới (dd/mm/yyyy) thì coi như bắt đầu một mục lịch thi mới
    for line in lines:
        date_match = re.search(r"(\d{2}/\d{1,2}/\d{4})", line)
        if date_match:
            if current_row:
                data_rows.append(current_row)
            ngay_thi = date_match.group(1)
            # Giữ lại toàn bộ text của dòng chứa ngày
            current_row = {
                "Ngày thi": ngay_thi,
                "Nội dung tích lũy": line 
            }
        else:
            if current_row:
                current_row["Nội dung tích lũy"] += " " + line
                
    if current_row:
        data_rows.append(current_row)
        
    final_list = []
    stt = 1
    
    # Duyệt qua từng cụm thông tin để bóc tách thông minh
    for row in data_rows:
        text_full = row["Nội dung tích lũy"]
        text_lower = text_full.lower()
        
        # Kiểm tra xem cụm này có chứa từ khóa Nguyễn Trình (không phân biệt hoa thường) hay không
        if "nguyễn trình" in text_lower:
            
            # 1. Xác định Cơ sở đào tạo thực tế dựa trên text tích lũy
            co_so = "Trung tâm Nguyễn Trình"
            if "hưng thịnh" in text_lower:
                co_so = "Trung tâm GDNN và SHLX Hưng Thịnh"
            elif "minh mẫn" in text_lower:
                co_so = "Nguyễn Trình & Mô tô Minh Mẫn"
            
            # 2. Tìm hạng thi bằng Regex (Ví dụ: Hạng A1, Hạng B2, Hạng C...)
            hang_match = re.search(r"(Hạng\s+[A-Z0-9,\s]+)", text_full, re.IGNORECASE)
            if hang_match:
                hang_thi = hang_match.group(1).strip()
            else:
                hang_thi = "Hạng A1, A"  # Giá trị mặc định nếu không tìm thấy chữ Hạng
            
            # 3. Tìm số lượng học viên thông minh hơn
            # Tìm tất cả các số có từ 2 đến 3 chữ số
            qty_candidates = re.findall(r"\b\d{2,3}\b", text_full)
            so_luong = "650" # Mặc định
            if qty_candidates:
                # Thường số lượng người thi sẽ đứng gần chữ "học viên", "thí sinh" hoặc nằm cuối cụm thông tin
                # Loại bỏ số ngày thi (nếu trùng) bằng cách lấy số cuối cùng không trùng với ngày/tháng/năm
                for num in reversed(qty_candidates):
                    if num not in row["Ngày thi"]:
                        so_luong = num
                        break

            dia_diem = "Sân sát hạch Nguyễn Trình (Châu Thành, Vĩnh Long)"
            
            final_list.append({
                "STT": stt,
                "Ngày thi": row["Ngày thi"],
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
        with st.spinner("🔄 Đang xử lý bóc tách dữ liệu thông minh..."):
            raw_text = extract_text_from_pdf(uploaded_file)
    else:
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
        
    if raw_text.strip():
        parsed_data = parse_pdf_content(raw_text)
        
        if parsed_data:
            df_display = pd.DataFrame(parsed_data).drop(columns=["STT"])
            st.success(f"🎉 Lọc thành công {len(parsed_data)} lịch thi liên quan đến Nguyễn Trình!")
            
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
                    file_name="Lich_Thi_Nguyen_Trinh_Ke_O.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.download_button(
                    label="🔵 Tải file WORD (.docx) - Chuẩn văn bản in ấn",
                    data=word_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh_Chuan.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.warning("⚠️ Không tìm thấy thông tin lịch thi liên quan đến Nguyễn Trình. Vui lòng kiểm tra lại file PDF.")
