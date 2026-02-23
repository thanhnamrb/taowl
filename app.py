import streamlit as st
import csv
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

# --- CÁC HÀM XỬ LÝ KỸ THUẬT ---

def replace_text_keep_style(paragraph, old_text, new_text):
    """Thay thế văn bản nhưng giữ nguyên định dạng (Font, Size, Bold) từ Template."""
    if old_text in paragraph.text:
        # Lưu lại định dạng của lượt chạy (run) đầu tiên làm mẫu
        style_dict = {}
        if paragraph.runs:
            r = paragraph.runs[0]
            style_dict['name'] = r.font.name
            style_dict['size'] = r.font.size
            style_dict['bold'] = r.bold
            style_dict['italic'] = r.italic
        
        # thực hiện thay thế nội dung
        paragraph.text = paragraph.text.replace(old_text, new_text)
        
        # Áp dụng lại định dạng cho tất cả các lượt chạy mới
        for r in paragraph.runs:
            if 'name' in style_dict and style_dict['name']: r.font.name = style_dict['name']
            if 'size' in style_dict and style_dict['size']: r.font.size = style_dict['size']
            if 'bold' in style_dict and style_dict['bold'] is not None: r.bold = style_dict['bold']
            if 'italic' in style_dict and style_dict['italic'] is not None: r.italic = style_dict['italic']

def remove_empty_paragraph(cell):
    """Loại bỏ dòng trống dư thừa sinh ra sau khi thực hiện gộp ô (merge)."""
    if len(cell.paragraphs) > 1 and cell.paragraphs[-1].text.strip() == "":
        p = cell.paragraphs[-1]._element
        p.getparent().remove(p)
        p._p = p._element = None

# --- GIAO DIỆN ỨNG DỤNG ---

st.set_page_config(page_title="Hệ thống Khởi tạo Tài liệu", page_icon="📄")

st.title("Công cụ Tự động hóa Danh sách Từ vựng")
st.write("Vui lòng nhập các thông số cấu hình và dữ liệu bảng để hệ thống khởi tạo tệp tin Word.")

# 1. Khu vực nhập thông số tệp
col1, col2 = st.columns([2, 1])
with col1:
    filename = st.text_input("Tên tệp tin khi tải về:", value="Vocab_List_Unit.docx")
with col2:
    # Tên tệp cần đảm bảo có đuôi .docx
    if not filename.endswith(".docx"):
        filename += ".docx"

# 2. Khu vực nhập Tiêu đề tài liệu
tieu_de = st.text_input("Tiêu đề hiển thị trong văn bản (thay thế cho [TITLE]):", 
                        value="VOCAB BUILDER UNIT 1.1: DAY IN, DAY OUT")

# 3. Khu vực nhập dữ liệu từ vựng
st.info("Lưu ý: Dán dữ liệu theo định dạng: No.,Word,Type,Pronunciation,Meaning. Nếu cột No. trống, hệ thống sẽ tự động gộp ô với hàng phía trên.")
raw_data = st.text_area("Dữ liệu từ vựng (CSV/Clipboard):", height=300, 
                        placeholder="Ví dụ:\n1,cruise,\"n, v\",,\n,cruiser,n,,")

# 4. Nút thực thi xử lý
if st.button("KHỞI TẠO TỆP TIN", type="primary"):
    if not raw_data.strip():
        st.error("Dữ liệu đầu vào không được để trống.")
    else:
        try:
            # Tải tệp mẫu từ thư mục gốc
            doc = Document("template.docx")
            
            # Thay thế Tiêu đề trong các đoạn văn bản (Paragraphs)
            for p in doc.paragraphs:
                replace_text_keep_style(p, "[TITLE]", tieu_de)

            # Xử lý dữ liệu văn bản thành dạng bảng
            f = io.StringIO(raw_data.strip())
            reader = csv.reader(f)
            
            try:
                # Bỏ qua dòng tiêu đề của dữ liệu dán vào (No., Word, Type...)
                next(reader) 
            except StopIteration:
                pass
            
            # Lấy bảng đầu tiên trong tệp mẫu
            table = doc.tables[0]
            
            # Xóa các hàng dữ liệu cũ (chỉ giữ lại hàng tiêu đề của bảng)
            while len(table.rows) > 1:
                tbl = table._tbl
                tbl.remove(table.rows[1]._tr)

            parent_cells = None
            
            for row_data in reader:
                if not row_data or "".join(row_data).strip() == "":
                    continue 
                
                # Đảm bảo dữ liệu có đủ 5 cột
                while len(row_data) < 5:
                    row_data.append("")
                
                # Thêm hàng mới vào bảng
                row = table.add_row()
                
                # Kiểm tra xem đây là từ mới (có STT) hay là từ thuộc Family (STT trống)
                is_new_entry = bool(row_data[0].strip())
                
                # Điền dữ liệu cho các cột Word, Type, Pronunciation, Meaning
                row.cells[1].text = row_data[1].strip()
                row.cells[2].text = row_data[2].strip()
                row.cells[3].text = row_data[3].strip()
                row.cells[4].text = row_data[4].strip()
                
                if is_new_entry:
                    # Điền số thứ tự và cập nhật hàng gốc (parent) để gộp sau này
                    row.cells[0].text = row_data[0].strip()
                    parent_cells = row.cells
                else:
                    # Nếu STT trống, thực hiện gộp ô cột số 0 với hàng gốc phía trên
                    if parent_cells:
                        merged_cell = row.cells[0].merge(parent_cells[0])
                        remove_empty_paragraph(merged_cell)

                # Định dạng phông chữ và căn lề cho hàng vừa thêm
                for i, cell in enumerate(row.cells):
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    for paragraph in cell.paragraphs:
                        # Cột No. căn giữa, các cột còn lại căn trái theo template
                        if i == 0:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Thiết lập phông chữ tiêu chuẩn Times New Roman
                        for run in paragraph.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)

            # Lưu tệp vào bộ nhớ tạm để chuẩn bị tải về
            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)
            
            st.success(f"Khởi tạo thành công tệp tin: {filename}")
            
            # Nút tải tệp tin dành cho người dùng
            st.download_button(
                label="TẢI TỆP TIN VỀ MÁY",
                data=output_stream,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except FileNotFoundError:
            st.error("Không tìm thấy tệp mẫu 'template.docx' trên máy chủ.")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")
