from __future__ import annotations

import csv
import io
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# ============================================================
# 1) DATA MODEL
# ============================================================

@dataclass(slots=True)
class VocabWord:
    word: str
    word_type: str = ""
    pronunciation: str = ""
    meaning: str = ""


@dataclass(slots=True)
class VocabFamily:
    number: str
    words: list[VocabWord] = field(default_factory=list)


@dataclass(slots=True)
class VocabDocument:
    unit: str
    heading_unit: str
    title: str
    document_type: str = "VOCAB BUILDER"
    families: list[VocabFamily] = field(default_factory=list)

    @property
    def unit_badge(self) -> str:
        return self.unit.strip() or "0"

    @property
    def display_heading_unit(self) -> str:
        if self.heading_unit.strip():
            return self.heading_unit.strip()
        return self.unit.split(".", 1)[0].strip() or self.unit

    @property
    def word_count(self) -> int:
        return sum(len(f.words) for f in self.families)

    @property
    def family_count(self) -> int:
        return len(self.families)


@dataclass(slots=True)
class TextStyle:
    font: str
    size: float
    color: str
    bold: bool = False
    italic: bool = False


@dataclass(slots=True)
class HeaderSegment:
    line: int
    text: str
    style: TextStyle


@dataclass(slots=True)
class ThemeConfig:
    badge: TextStyle
    document_type: TextStyle
    heading: TextStyle
    header_segments: list[HeaderSegment]


# ============================================================
# 2) CSV PARSER
# ============================================================

HEADER_ALIASES = {
    "no": {"no", "no.", "number", "stt", "stt."},
    "word": {"word", "vocabulary", "vocab"},
    "type": {"type", "word type", "pos", "part of speech"},
}


def looks_like_header(row: list[str]) -> bool:
    cells = [c.strip().lower() for c in row]
    return (
        len(cells) >= 3
        and cells[0] in HEADER_ALIASES["no"]
        and cells[1] in HEADER_ALIASES["word"]
        and cells[2] in HEADER_ALIASES["type"]
    )


def parse_vocab_csv(raw_data: str, *, unit: str, heading_unit: str, title: str, document_type: str) -> VocabDocument:
    rows = list(csv.reader(io.StringIO((raw_data or "").strip())))
    rows = [r for r in rows if r and "".join(r).strip()]
    if rows and looks_like_header(rows[0]):
        rows = rows[1:]

    doc = VocabDocument(
        unit=unit.strip(),
        heading_unit=heading_unit.strip(),
        title=title.strip(),
        document_type=(document_type or "VOCAB BUILDER").strip().upper(),
    )

    current: VocabFamily | None = None
    for row in rows:
        row = list(row)
        while len(row) < 5:
            row.append("")
        no, word, word_type, pronunciation, meaning = [x.strip() for x in row[:5]]
        if not word:
            continue
        if no:
            current = VocabFamily(no)
            doc.families.append(current)
        elif current is None:
            current = VocabFamily("")
            doc.families.append(current)
        current.words.append(VocabWord(word, word_type, pronunciation, meaning))
    return doc


# ============================================================
# 3) DOCX HELPERS
# ============================================================

def clean_hex(value: str, fallback: str = "000000") -> str:
    v = (value or "").strip().lstrip("#").upper()
    return v if len(v) == 6 and all(ch in "0123456789ABCDEF" for ch in v) else fallback


def apply_run_style(run, style: TextStyle) -> None:
    run.font.name = style.font
    run.font.size = Pt(float(style.size))
    run.font.bold = bool(style.bold)
    run.font.italic = bool(style.italic)
    run.font.color.rgb = RGBColor.from_string(clean_hex(style.color))
    rfonts = run._r.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), style.font)


def remove_text_runs_preserve_drawings(paragraph) -> None:
    for run in list(paragraph.runs):
        if not run._r.xpath(".//w:drawing | .//w:pict"):
            run._element.getparent().remove(run._element)


def set_single_run_paragraph(paragraph, text: str, style: TextStyle, *, align=None) -> None:
    remove_text_runs_preserve_drawings(paragraph)
    if align is not None:
        paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.keep_together = True
    run = paragraph.add_run(text)
    apply_run_style(run, style)


def set_title_block(doc: Document, data: VocabDocument, theme: ThemeConfig) -> None:
    title = doc.tables[0]

    badge = title.rows[0].cells[0]
    badge.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_single_run_paragraph(
        badge.paragraphs[0], data.unit_badge, theme.badge,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    type_cell = title.rows[0].cells[2]
    type_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_single_run_paragraph(
        type_cell.paragraphs[0], data.document_type, theme.document_type,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )

    main = title.rows[1].cells[2]
    main.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    while len(main.paragraphs) > 1:
        p = main.paragraphs[-1]._element
        p.getparent().remove(p)
    heading = f"UNIT {data.display_heading_unit}: {data.title.upper()}"
    set_single_run_paragraph(main.paragraphs[0], heading, theme.heading, align=WD_ALIGN_PARAGRAPH.LEFT)


def set_header(doc: Document, segments: list[HeaderSegment]) -> None:
    header = doc.sections[0].header
    while len(header.paragraphs) < 2:
        header.add_paragraph()

    for line_no in (1, 2):
        p = header.paragraphs[line_no - 1]
        remove_text_runs_preserve_drawings(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        items = [s for s in segments if s.line == line_no and s.text.strip()]
        for i, segment in enumerate(items):
            text = segment.text.strip() + (" " if i < len(items) - 1 else "")
            run = p.add_run(text)
            apply_run_style(run, segment.style)


def set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = tr_pr.find(qn("w:tblHeader"))
    if node is None:
        node = OxmlElement("w:tblHeader")
        tr_pr.append(node)
    node.set(qn("w:val"), "true")


def set_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = tr_pr.find(qn("w:cantSplit"))
    if node is None:
        node = OxmlElement("w:cantSplit")
        tr_pr.append(node)


def clear_vmerge(cell) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    vm = tcpr.find(qn("w:vMerge"))
    if vm is not None:
        tcpr.remove(vm)


def set_vmerge_tc(tc, *, restart: bool) -> None:
    tcpr = tc.get_or_add_tcPr()
    vm = tcpr.find(qn("w:vMerge"))
    if vm is None:
        vm = OxmlElement("w:vMerge")
        tcpr.append(vm)
    if restart:
        vm.set(qn("w:val"), "restart")
    else:
        vm.attrib.pop(qn("w:val"), None)


def set_family_keep_next(row, enabled: bool) -> None:
    for cell in row.cells:
        for p in cell.paragraphs:
            p.paragraph_format.keep_with_next = enabled
            p.paragraph_format.keep_together = True


def set_cell_text_preserve_prototype(cell, text: str, *, center: bool = False) -> None:
    while len(cell.paragraphs) > 1:
        p = cell.paragraphs[-1]._element
        p.getparent().remove(p)
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    if p.runs:
        p.runs[0].text = text or ""
        for r in list(p.runs[1:]):
            r._element.getparent().remove(r._element)
    else:
        r = p.add_run(text or "")
        apply_run_style(r, TextStyle("Times New Roman", 12, "000000", False))
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def populate_vocab_table(doc: Document, data: VocabDocument) -> None:
    table = doc.tables[1]
    if len(table.rows) < 3:
        raise RuntimeError("Template cần header + 2 prototype rows.")

    table.autofit = False
    set_repeat_header(table.rows[0])
    set_cant_split(table.rows[0])

    root_proto = table.rows[1]
    derivative_proto = table.rows[2]
    merge_groups: list[list] = []

    for family in data.families:
        family_rows = []
        for word_index, item in enumerate(family.words):
            source = root_proto if word_index == 0 else derivative_proto
            table._tbl.append(deepcopy(source._tr))
            row = table.rows[-1]
            set_cant_split(row)
            for c in row.cells:
                clear_vmerge(c)

            set_cell_text_preserve_prototype(row.cells[0], family.number if word_index == 0 else "", center=True)
            set_cell_text_preserve_prototype(row.cells[1], item.word)
            set_cell_text_preserve_prototype(row.cells[2], item.word_type)
            set_cell_text_preserve_prototype(row.cells[3], item.pronunciation)
            set_cell_text_preserve_prototype(row.cells[4], item.meaning)
            family_rows.append(row)

        for i, row in enumerate(family_rows):
            set_family_keep_next(row, i < len(family_rows) - 1)
        if len(family_rows) > 1:
            merge_groups.append(family_rows)

    table._tbl.remove(root_proto._tr)
    table._tbl.remove(derivative_proto._tr)

    # Important: apply merges only after ALL rows are filled.
    # Otherwise python-docx collapses later row.cells proxies into previous merged cells.
    for family_rows in merge_groups:
        set_vmerge_tc(family_rows[0]._tr.tc_lst[0], restart=True)
        for row in family_rows[1:]:
            set_vmerge_tc(row._tr.tc_lst[0], restart=False)


def render_vocab_docx(data: VocabDocument, *, template_path: Path, theme: ThemeConfig) -> bytes:
    doc = Document(str(template_path))
    set_title_block(doc, data, theme)
    set_header(doc, theme.header_segments)
    populate_vocab_table(doc, data)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


# ============================================================
# 4) STREAMLIT UI
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "lx_vocab_template.docx"

st.set_page_config(page_title="LingualXplore · Vocab Builder", page_icon="📘", layout="wide")
st.title("LingualXplore · VOCAB Builder")
st.caption("V2: badge 5.1 đúng vị trí, header chỉnh từng từ, bảng liền mạch theo file VOCAB gốc.")

with st.container(border=True):
    c1, c2, c3, c4 = st.columns([1, 1, 1.5, 2.4])
    with c1:
        unit = st.text_input("Sub-unit / badge", "5.1")
    with c2:
        heading_unit = st.text_input("Unit ở heading", "5")
    with c3:
        document_type = st.text_input("Document type", "VOCAB BUILDER")
    with c4:
        title = st.text_input("Title", "SPECIAL DAYS")
    filename = st.text_input("Tên file", "VOCAB BUILDER UNIT 5.1 - SPECIAL DAYS.docx")
    if not filename.lower().endswith(".docx"):
        filename += ".docx"

with st.expander("🎨 Badge + title", False):
    a, b, c = st.columns(3)
    with a:
        badge_font = st.text_input("Badge font", "Orbitron")
        badge_size = st.number_input("Badge size", 20.0, 80.0, 48.0, 1.0)
        badge_color = st.color_picker("Badge color", "#FFFFFF")
        badge_bold = st.checkbox("Badge bold", True)
    with b:
        type_font = st.text_input("Document type font", "Monument Extended")
        type_size = st.number_input("Document type size", 10.0, 40.0, 18.0, 0.5)
        type_color = st.color_picker("Document type color", "#BF4E14")
        type_bold = st.checkbox("Document type bold", False)
    with c:
        heading_font = st.text_input("Heading font", "Arial")
        heading_size = st.number_input("Heading size", 14.0, 50.0, 27.0, 0.5)
        heading_color = st.color_picker("Heading color", "#000000")
        heading_bold = st.checkbox("Heading bold", True)


def edit_header_word(key, default_text, default_font, default_color, default_bold):
    st.markdown(f"**{default_text}**")
    c1, c2, c3, c4, c5 = st.columns([1.2, 1.5, 1, 0.8, 0.7])
    with c1:
        text = st.text_input("Text", default_text, key=f"{key}_text", label_visibility="collapsed")
    with c2:
        font = st.text_input("Font", default_font, key=f"{key}_font", label_visibility="collapsed")
    with c3:
        color = st.color_picker("Color", default_color, key=f"{key}_color", label_visibility="collapsed")
    with c4:
        size = st.number_input("Size", 6.0, 30.0, 10.5, 0.5, key=f"{key}_size", label_visibility="collapsed")
    with c5:
        bold = st.checkbox("B", default_bold, key=f"{key}_bold")
    return text, font, size, color, bold


with st.expander("✍️ Header: chỉnh từng từ", False):
    st.caption("Dòng 1: Supplementary Exercises · Dòng 2: for KET Learners")
    seg1 = edit_header_word("supp", "Supplementary", "Arial", "#BF4E14", False)
    seg2 = edit_header_word("exercises", "Exercises", "Arial", "#BF4E14", False)
    seg3 = edit_header_word("for", "for", "Arial", "#BF4E14", False)
    seg4 = edit_header_word("ket", "KET", "Orbitron", "#156082", True)
    seg5 = edit_header_word("learners", "Learners", "Orbitron", "#156082", True)

st.info(
    "CSV 5 cột: No., Word, Type, Pronunciation, Meaning. "
    "No. trống = cùng word family. Bảng output là một table duy nhất và body plain như file gốc."
)
raw_data = st.text_area(
    "Vocabulary data", height=360,
    placeholder='No.,Word,Type,Pronunciation,Meaning\n1,special,"adj, n",,\n,specially,adv,,',
)

if st.button("KHỞI TẠO TÀI LIỆU", type="primary", use_container_width=True):
    if not TEMPLATE_PATH.exists():
        st.error(f"Thiếu master template: {TEMPLATE_PATH}")
        st.stop()
    if not raw_data.strip():
        st.error("Vocabulary data đang trống.")
        st.stop()

    data = parse_vocab_csv(
        raw_data, unit=unit, heading_unit=heading_unit, title=title, document_type=document_type
    )
    if not data.families:
        st.error("Không parse được word family nào.")
        st.stop()
    if data.families[0].number == "":
        st.error("Dòng từ đầu tiên phải có No./STT.")
        st.stop()

    def mkseg(line, values):
        text, font, size, color, bold = values
        return HeaderSegment(line, text, TextStyle(font, size, color.lstrip("#"), bold))

    theme = ThemeConfig(
        badge=TextStyle(badge_font, badge_size, badge_color.lstrip("#"), badge_bold),
        document_type=TextStyle(type_font, type_size, type_color.lstrip("#"), type_bold),
        heading=TextStyle(heading_font, heading_size, heading_color.lstrip("#"), heading_bold),
        header_segments=[
            mkseg(1, seg1), mkseg(1, seg2), mkseg(2, seg3), mkseg(2, seg4), mkseg(2, seg5)
        ],
    )

    try:
        output = render_vocab_docx(data, template_path=TEMPLATE_PATH, theme=theme)
    except Exception as exc:
        st.exception(exc)
        st.stop()

    st.success(f"Đã tạo {data.family_count} families / {data.word_count} vocabulary rows.")
    st.download_button(
        "TẢI FILE WORD", output, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary", use_container_width=True,
    )
