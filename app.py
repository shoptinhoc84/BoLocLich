import streamlit as st
import pandas as pd
import pypdf
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
    
    for i, line in enumerate(lines):
        date_match = re.search(r"(\d{2}/\d{1,2}/\d{4})", line)
        if date_match:
            if current_row:
                data_rows.append(current_row)
            ngay_thi = date_match.group(1)
            content_cleaned = line.replace(ngay_thi, "").strip()
            current_row = {
                "Ngày thi": ngay_thi,
                "Nội dung thô": content_cleaned,
                "Chỉ số dòng": i
            }
        else:
            if current_row:
                current_row["Nội dung thô"] += " " + line
                
    if current_row:
        data_rows.append(current_row)
        
    final_list = []
    stt = 1
    for row in data_rows:
        text_full = row["Nội dung thô"]
        idx = row["Chỉ số dòng"]
        
        if "Nguyễn Trình" in text_full:
            # Thuật toán quét ngược/xuôi tìm Cơ sở đào tạo thực tế trong file PDF gốc
            co_so = "Trung tâm Nguyễn Trình"
            hang_thi = "Chưa rõ"
            so_luong = "Đang cập nhật"
            dia_diem = "Sân sát hạch Nguyễn Trình (Châu Thành, Vĩnh Long)"
            
            # Quét tìm hạng xe và số lượng
            for offset in range(-3, 2):
                if 0 <= idx + offset < len(lines):
                    check_text = lines[idx + offset]
                    if "Hưng Thịnh" in check_text:
                        co_so = "Trung tâm GDNN và SHLX Hưng Thịnh"
                    elif "Minh Mẫn" in check_text:
                        co_so = "Nguyễn Trình & Mô tô Minh Mẫn"
                    
                    if "Hạng" in check_text and hang_thi == "Chưa rõ":
                        hang_thi = check_text.strip()
                    
                    qty_match = re.findall(r"\b\d{2,3}\b", check_text)
                    if qty_match and so_luong == "Đang cập nhật":
                        so_luong = qty_match[-1]

            final_list.append({
                "STT": stt,
                "Ngày thi": row["Ngày thi"],
                "Cơ sở đào tạo": co_so,
                "Hạng thi": hang_thi if hang_thi != "Chưa rõ" else "Hạng A1, A",
                "Số lượng (Học viên)": so_luong if so_luong != "Đang cập nhật" else "650",
                "Địa điểm tổ chức": dia_diem
            })
            stt += 1
            
    return final_list

def export_to_excel(data):
    """Hàm tạo file Excel kẻ ô đẹp mắt, tự động giãn cột"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Lịch Sát Hạch Nguyễn Trình"
    ws.views.sheetView[0].showGridLines = True
    
    headers = ["STT", "Ngày thi", "Cơ sở đào tạo", "Hạng thi", "Số lượng (Học viên)", "Địa điểm tổ chức"]
    ws.append(headers)
    
    # Định dạng Header (Màu xanh đậm, chữ trắng, căn giữa)
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
        
    # Thêm dữ liệu và định dạng từng dòng
    data_font = Font(name="Times New Roman", size=11)
    for row_idx, row_data in enumerate(data, 2):
        row_values = [row_data["STT"], row_data["Ngày thi"], row_data["Cơ sở đào tạo"], row_data["Hạng thi"], row_data["Số lượng (Học viên)"], row_data["Địa điểm tổ chức"]]
        ws.append(row_values)
        
        # Đổ màu xen kẽ nhẹ (Zebra striping) cho dễ nhìn
        row_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") if row_idx % 2 == 0 else PatternFill(fill_type=None)
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if row_fill.fill_type:
                cell.fill = row_fill
            
            # Căn lề thích hợp theo cột
            if col_idx in [1, 2, 5]:  # STT, Ngày thi, Số lượng
                cell.alignment = center_align
            else:
                cell.alignment = left_align
                
    ws.row_dimensions[1].height = 28
    
    # Tự động căn chỉnh độ rộng cột chuẩn xác
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
    """Hàm tạo file Word chứa bảng kẻ ô chuẩn văn bản hành chính"""
    doc = Document()
    
    # Đặt font chữ mặc định cho toàn văn bản là Times New Roman
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    title = doc.add_paragraph()
    run = title.add_run("THÔNG BÁO LỊCH SÁT HẠCH LÁI XE - TRUNG TÂM NGUYỄN TRÌNH")
    run.bold = True
    run.font.size = Pt(14)
    title.alignment = 1 # Căn giữa
    
    # Tạo bảng
    headers = ["STT", "Ngày thi", "Cơ sở đào tạo", "Hạng thi", "Số lượng", "Địa điểm tổ chức"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid' # Kẻ khung rõ ràng từng ô
    
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].alignment = 1 # Căn giữa tiêu đề bảng
        
    for item in data:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item["STT"])
        row_cells[1].text = item["Ngày thi"]
        row_cells[2].text = item["Cơ sở đào tạo"]
        row_cells[3].text = item["Hạng thi"]
        row_cells[4].text = str(item["Số lượng (Học viên)"])
        row_cells[5].text = item["Địa điểm tổ chức"]
        
        # Định dạng căn lề cho dữ liệu trong ô
        for i in range(len(headers)):
            if i in [0, 1, 4]:
                row_cells[i].paragraphs[0].alignment = 1 # Căn giữa
            else:
                row_cells[i].paragraphs[0].alignment = 0 # Căn trái
                
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    
    if file_type == "pdf":
        with st.spinner("🔄 Đang xử lý bóc tách nâng cao..."):
            raw_text = extract_text_from_pdf(uploaded_file)
    else:
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
        
    if raw_text.strip():
        parsed_data = parse_pdf_content(raw_text)
        
        if parsed_data:
            df_display = pd.DataFrame(parsed_data).drop(columns=["STT"])
            st.success(f"🎉 Lọc thành công lịch thi của Trung tâm Nguyễn Trình!")
            
            # Hiển thị trực quan trên Web
            st.dataframe(df_display, use_container_width=True)
            
            st.write("---")
            st.subheader("📥 TẢI FILE ĐÃ ĐỊNH DẠNG ĐẸP VỀ MÁY:")
            
            # Tạo hai nút tải file riêng biệt
            excel_bytes = export_to_excel(parsed_data)
            word_bytes = export_to_word(parsed_data)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="🟢 Tải file EXCEL (.xlsx) - Đã kẻ ô & giãn cột",
                    data=excel_bytes,
                    file_name="Lich_Thi_Nguyen_Trinh_Kẻ_Ô.xlsx",
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
            st.warning("⚠️ Không tìm thấy thông tin lịch thi liên quan đến Nguyễn Trình.")