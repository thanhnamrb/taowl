import streamlit as st
import csv
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

# --- CÁC HÀM XỬ LÝ (GIỮ NGUYÊN NHƯ CŨ) ---
def replace_text_keep_style(paragraph, old_text, new_text):
    if old_text in paragraph.text:
        style_dict = {}
        if paragraph.runs:
            r = paragraph.runs[0]
            style_dict['name'] = r.font.name
            style_dict['size'] = r.font.size
            style_dict['bold'] = r.bold
        paragraph.text = paragraph.text.replace(old_text, new_text)
        for r in paragraph.runs:
            if 'name' in style_dict and style_dict['name']: r.font.name = style_dict['name']
            if 'size' in style_dict and style_dict['size']: r.font.size = style_dict['size']
            if 'bold' in style_dict and style_dict['bold'] is not None: r.bold = style_dict['bold']

def remove_empty_paragraph(cell):
    if len(cell.paragraphs) > 1 and cell.paragraphs[-1].text.strip() == "":
        p = cell.paragraphs[-1]._element
        p.getparent().remove(p)
        p._p = p._element = None

# --- GIAO DIỆN WEB DÀNH CHO MẸ ---
st.set_page_config(page_title="Công cụ tạo Word", page_icon="📝")

st.title("Phần mềm Tạo File Word - Dành cho Mẹ 💖")
st.write("Mẹ chỉ cần điền thông tin và dán bảng từ vựng vào đây nhé!")

# Các ô nhập liệu
filename = st.text_input("1. Tên file Word muốn tải về:", value="Tu_Vung_Unit_1.docx")
tieu_de = st.text_input("2. Tiêu đề trên cùng của file:", value="VOCAB BUILDER UNIT 1.1: DAY IN, DAY OUT")
raw_data = st.text_area("3. Dán danh sách từ vựng vào đây (Copy từ Excel):", height=250, 
                        value='No.,Word,Type,Pronunciation,Meaning\n1,cruise,"n, v",,\n,cruiser,n,,')

# Nút bấm chính
if st.button("🚀 BẤM VÀO ĐÂY ĐỂ TẠO FILE WORD", type="primary"):
    try:
        # Máy chủ web sẽ tự động đọc file template đã được giấu sẵn
        doc = Document("template.docx")
        
        # Thay tiêu đề
        for p in doc.paragraphs:
            replace_text_keep_style(p, "[TITLE]", tieu_de)

        # Xử lý dữ liệu
        f = io.StringIO(raw_data.strip())
        reader = csv.reader(f)
        next(reader) # Bỏ qua tiêu đề
        
        table = doc.tables[0]
        while len(table.rows) > 1:
            tbl = table._tbl
            tbl.remove(table.rows[1]._tr)

        parent_cells = None
        for row_data in reader:
            if not row_data or "".join(row_data).strip() == "":
                continue 
            while len(row_data) < 5:
                row_data.append("")
                
            row = table.add_row()
            is_new_family = bool(row_data[0].strip())
            
            row.cells[1].text = row_data[1].strip()
            row.cells[2].text = row_data[2].strip()
            row.cells[3].text = row_data[3].strip()
            row.cells[4].text = row_data[4].strip()
            
            if is_new_family:
                row.cells[0].text = row_data[0].strip()
                parent_cells = row.cells
            else:
                if parent_cells:
                    c0 = row.cells[0].merge(parent_cells[0])
                    remove_empty_paragraph(c0)

            for i, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for paragraph in cell.paragraphs:
                    if i == 0:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)

        # Chuẩn bị file để tải về ngay trên Web
        if not filename.endswith(".docx"):
            filename += ".docx"
            
        # Lưu file vào bộ nhớ đệm (BytesIO) thay vì lưu vào máy tính
        bio = io.BytesIO()
        doc.save(bio)
        
        st.success("🎉 Đã tạo xong! Mẹ bấm nút tải về ở bên dưới nhé.")
        
        # Nút tải file xuất hiện
        st.download_button(
            label="⬇️ TẢI FILE WORD VỀ MÁY",
            data=bio.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"Có lỗi xảy ra, mẹ kiểm tra lại dữ liệu dán vào nhé. (Mã lỗi: {e})")