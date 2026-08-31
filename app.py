from __future__ import annotations

import csv
import io
import tempfile
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile

import streamlit as st
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# ===================== DATA MODEL =====================
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
    title: str
    document_type: str = "VOCAB BUILDER"
    heading_unit: str = ""
    families: list[VocabFamily] = field(default_factory=list)

    @property
    def word_count(self) -> int:
        return sum(len(f.words) for f in self.families)

    @property
    def family_count(self) -> int:
        return len(self.families)

    @property
    def unit_badge(self) -> str:
        # The user's corrected reference uses the full sub-unit in the badge: 5.1.
        return (self.unit or "").strip() or "0"

    @property
    def display_heading_unit(self) -> str:
        # 5.1 -> UNIT 5 by default, matching the corrected reference.
        if self.heading_unit.strip():
            return self.heading_unit.strip()
        return (self.unit or "").split(".", 1)[0].strip() or self.unit

# ===================== PARSER =====================
HEADER_ALIASES = {
    "no": {"no", "no.", "number", "stt", "stt."},
    "word": {"word", "vocabulary", "vocab"},
    "type": {"type", "word type", "pos", "part of speech"},
}


def _looks_like_header(row: list[str]) -> bool:
    cells = [c.strip().lower() for c in row]
    if len(cells) < 3:
        return False
    return (
        cells[0] in HEADER_ALIASES["no"]
        and cells[1] in HEADER_ALIASES["word"]
        and cells[2] in HEADER_ALIASES["type"]
    )


def parse_vocab_csv(
    raw_data: str,
    *,
    unit: str,
    title: str,
    document_type: str = "VOCAB BUILDER",
    heading_unit: str = "",
) -> VocabDocument:
    """Parse CSV/clipboard data into a semantic document model.

    Expected columns: No., Word, Type, Pronunciation, Meaning.
    A blank No. means the row belongs to the previous word family.
    """
    stream = io.StringIO((raw_data or "").strip())
    rows = list(csv.reader(stream))
    rows = [r for r in rows if r and "".join(r).strip()]

    if rows and _looks_like_header(rows[0]):
        rows = rows[1:]

    document = VocabDocument(
        unit=(unit or "").strip(),
        title=(title or "").strip(),
        document_type=(document_type or "VOCAB BUILDER").strip().upper(),
        heading_unit=(heading_unit or "").strip(),
    )

    current_family: VocabFamily | None = None

    for source_index, row in enumerate(rows, start=1):
        row = list(row)
        while len(row) < 5:
            row.append("")
        row = row[:5]

        no, word, word_type, pronunciation, meaning = (c.strip() for c in row)

        if not word:
            # The validator will catch meaningful malformed rows. Pure blank rows were removed above.
            continue

        if no:
            current_family = VocabFamily(number=no)
            document.families.append(current_family)
        elif current_family is None:
            # Preserve malformed input as a synthetic family so validation can explain it.
            current_family = VocabFamily(number="")
            document.families.append(current_family)

        current_family.words.append(
            VocabWord(
                word=word,
                word_type=word_type,
                pronunciation=pronunciation,
                meaning=meaning,
            )
        )

    return document

# ===================== THEME =====================
@dataclass(slots=True)
class TextStyle:
    font: str = "Arial"
    size: float = 10.5
    color: str = "000000"
    bold: bool = False
    italic: bool = False


@dataclass(slots=True)
class HeaderSegment:
    line: int
    text: str
    style: TextStyle


@dataclass(slots=True)
class TableTheme:
    # The geometry/spacing comes from the original pre-app VOCAB document.
    # These options only skin the table so it belongs to the LingualXplore theme.
    header_no_fill: str = "BF4E14"
    header_fill: str = "0F4761"
    header_text: TextStyle = field(
        default_factory=lambda: TextStyle("Arial", 10.5, "FFFFFF", True, False)
    )
    body_text: TextStyle = field(
        default_factory=lambda: TextStyle("Times New Roman", 12, "000000", False, False)
    )
    outer_border_color: str = "0F4761"
    outer_border_pt: float = 1.0
    inner_border_color: str = "B8C8D0"
    inner_border_pt: float = 0.5


@dataclass(slots=True)
class ThemeConfig:
    badge: TextStyle = field(
        default_factory=lambda: TextStyle("Orbitron", 48, "FFFFFF", True)
    )
    document_type: TextStyle = field(
        default_factory=lambda: TextStyle("Monument Extended", 18, "BF4E14", False)
    )
    heading: TextStyle = field(
        default_factory=lambda: TextStyle("Arial", 27, "000000", True)
    )
    header_segments: list[HeaderSegment] = field(default_factory=list)
    header_alignments: dict[int, str] = field(
        default_factory=lambda: {1: "right", 2: "right", 3: "right"}
    )
    # Extra body clearance under the header. Renderer automatically adds a little more
    # if a third header line is used.
    title_clearance_pt: float = 28.0
    table: TableTheme = field(default_factory=TableTheme)


def default_header_segments() -> list[HeaderSegment]:
    # Default is the user's second reference image.
    peach = "E7A07F"
    blue = "7FA9BC"
    return [
        HeaderSegment(1, "Supplementary", TextStyle("Arial", 10.5, peach, False)),
        HeaderSegment(1, "Vocabulary", TextStyle("Arial", 10.5, peach, False)),
        HeaderSegment(2, "for", TextStyle("Arial", 10.5, peach, False)),
        HeaderSegment(2, "KET", TextStyle("Orbitron", 10.5, blue, True)),
        HeaderSegment(2, "Learners", TextStyle("Orbitron", 10.5, blue, True)),
        HeaderSegment(2, "in", TextStyle("Arial", 10.5, peach, False)),
        HeaderSegment(2, "Orwell", TextStyle("Orbitron", 10.5, blue, True)),
        HeaderSegment(2, "Classes", TextStyle("Orbitron", 10.5, blue, True)),
    ]


def default_theme() -> ThemeConfig:
    return ThemeConfig(header_segments=default_header_segments())

# ===================== RENDERER =====================
def _hex(value: str, fallback: str = "000000") -> str:
    v = (value or "").strip().lstrip("#").upper()
    return v if len(v) == 6 and all(ch in "0123456789ABCDEF" for ch in v) else fallback


def _apply_run_style(run, style: TextStyle) -> None:
    run.font.name = style.font
    run.font.size = Pt(float(style.size))
    run.font.bold = bool(style.bold)
    run.font.italic = bool(style.italic)
    run.font.color.rgb = RGBColor.from_string(_hex(style.color))
    rfonts = run._r.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), style.font)


def _remove_text_runs_preserve_drawings(paragraph) -> None:
    for run in list(paragraph.runs):
        if not run._r.xpath(".//w:drawing | .//w:pict"):
            run._element.getparent().remove(run._element)


def _set_single_run_paragraph(paragraph, text: str, style: TextStyle, *, align=None) -> None:
    _remove_text_runs_preserve_drawings(paragraph)
    if align is not None:
        paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.keep_together = True
    run = paragraph.add_run(text)
    _apply_run_style(run, style)


def _set_title_block(doc: Document, document: VocabDocument, theme: ThemeConfig) -> None:
    title = doc.tables[0]

    badge = title.rows[0].cells[0]
    badge.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _set_single_run_paragraph(
        badge.paragraphs[0], document.unit_badge, theme.badge,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    type_cell = title.rows[0].cells[2]
    type_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _set_single_run_paragraph(
        type_cell.paragraphs[0], document.document_type.upper(), theme.document_type,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )

    main = title.rows[1].cells[2]
    main.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    while len(main.paragraphs) > 1:
        p = main.paragraphs[-1]._element
        p.getparent().remove(p)
    heading = f"UNIT {document.display_heading_unit}: {document.title.upper()}"
    _set_single_run_paragraph(main.paragraphs[0], heading, theme.heading, align=WD_ALIGN_PARAGRAPH.LEFT)


def _alignment(value: str):
    value = (value or "center").lower()
    return {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(value, WD_ALIGN_PARAGRAPH.CENTER)


def _set_header(doc: Document, theme: ThemeConfig) -> None:
    """Render 1-3 independently styled header lines.

    The logo drawing stays intact. Each HeaderSegment can be a word or a whole phrase,
    so the user can add/remove/reorder wording without changing Python.
    """
    header = doc.sections[0].header
    used_lines = max([s.line for s in theme.header_segments if s.text.strip()] or [1])
    used_lines = max(1, min(3, used_lines))

    while len(header.paragraphs) < used_lines:
        header.add_paragraph()

    # Clear text on the first 3 paragraphs but preserve floating drawings/logo.
    for line_no in range(1, 4):
        if len(header.paragraphs) < line_no:
            break
        p = header.paragraphs[line_no - 1]
        _remove_text_runs_preserve_drawings(p)
        p.alignment = _alignment(theme.header_alignments.get(line_no, "center"))
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

        items = [s for s in theme.header_segments if int(s.line) == line_no and s.text.strip()]
        for index, segment in enumerate(items):
            # A segment may be a word OR a phrase. This makes header editing flexible.
            text = segment.text.strip()
            if index < len(items) - 1:
                text += " "
            run = p.add_run(text)
            _apply_run_style(run, segment.style)

    # Hide any unused extra header text paragraphs while keeping drawing runs.
    for p in header.paragraphs[used_lines:3]:
        _remove_text_runs_preserve_drawings(p)


def _set_title_clearance(doc: Document, theme: ThemeConfig) -> None:
    """Prevent title artwork from entering the header zone.

    The V3 template contains a blank paragraph directly before the title table.
    We control its spacing here. A third header line automatically receives more room.
    """
    title_tbl = doc.tables[0]._tbl
    prev = title_tbl.getprevious()
    if prev is None or prev.tag != qn("w:p"):
        p = OxmlElement("w:p")
        title_tbl.addprevious(p)
        prev = p

    used_lines = max([s.line for s in theme.header_segments if s.text.strip()] or [1])
    extra = max(0, min(3, used_lines) - 2) * 14.0
    clearance = max(18.0, float(theme.title_clearance_pt) + extra)

    ppr = prev.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    # Keep this paragraph visually empty and use only after-spacing as clearance.
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), str(int(clearance * 20)))  # pt -> twips
    spacing.set(qn("w:line"), "20")
    spacing.set(qn("w:lineRule"), "exact")


def _set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = tr_pr.find(qn("w:tblHeader"))
    if node is None:
        node = OxmlElement("w:tblHeader")
        tr_pr.append(node)
    node.set(qn("w:val"), "true")


def _set_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    node = tr_pr.find(qn("w:cantSplit"))
    if node is None:
        node = OxmlElement("w:cantSplit")
        tr_pr.append(node)


def _set_family_keep_next(row, enabled: bool) -> None:
    for cell in row.cells:
        for p in cell.paragraphs:
            p.paragraph_format.keep_with_next = enabled
            p.paragraph_format.keep_together = True


def _clear_vmerge(cell) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    vm = tcpr.find(qn("w:vMerge"))
    if vm is not None:
        tcpr.remove(vm)


def _set_vmerge_tc(tc, *, restart: bool) -> None:
    tcpr = tc.get_or_add_tcPr()
    vm = tcpr.find(qn("w:vMerge"))
    if vm is None:
        vm = OxmlElement("w:vMerge")
        tcpr.append(vm)
    if restart:
        vm.set(qn("w:val"), "restart")
    else:
        vm.attrib.pop(qn("w:val"), None)




def _set_cell_no_wrap(cell, enabled: bool = True) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    node = tcpr.find(qn("w:noWrap"))
    if enabled and node is None:
        node = OxmlElement("w:noWrap")
        tcpr.append(node)
    elif not enabled and node is not None:
        tcpr.remove(node)



def _set_fixed_table_layout_and_no_column(table, min_no_inches: float = 0.50) -> None:
    """Force a fixed table grid and keep the No. column wide enough for one line.

    The original VOCAB geometry is preserved as much as possible: only the small
    amount needed by the No. column is borrowed from the Meaning column.
    """
    tblpr = table._tbl.tblPr
    layout = tblpr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblpr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid_cols = table._tbl.tblGrid.gridCol_lst
    if len(grid_cols) < 5:
        return

    widths = [int(col.w) for col in grid_cols]
    min_no_emu = int(Inches(float(min_no_inches)))
    if widths[0] < min_no_emu:
        delta = min_no_emu - widths[0]
        widths[0] = min_no_emu
        # Borrow only from Meaning so Word/Type/Pronunciation keep the original feel.
        widths[-1] = max(int(Inches(1.50)), widths[-1] - delta)

    for col, width in zip(grid_cols, widths):
        col.w = width

    # tcW in dxa/twips makes Word respect the same fixed geometry in every row.
    for row in table.rows:
        for i, cell in enumerate(row.cells[:len(widths)]):
            tcpr = cell._tc.get_or_add_tcPr()
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:type"), "dxa")
            tcw.set(qn("w:w"), str(max(1, round(widths[i] / 635))))
        if row.cells:
            _set_cell_no_wrap(row.cells[0], True)


def _set_cell_shading(cell, fill: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), _hex(fill, "FFFFFF"))


def _set_table_borders(table, theme: TableTheme) -> None:
    tblpr = table._tbl.tblPr
    borders = tblpr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblpr.append(borders)

    def set_edge(name: str, color: str, pt: float):
        edge = borders.find(qn(f"w:{name}"))
        if edge is None:
            edge = OxmlElement(f"w:{name}")
            borders.append(edge)
        edge.set(qn("w:val"), "single")
        edge.set(qn("w:sz"), str(max(1, int(round(float(pt) * 8)))))
        edge.set(qn("w:space"), "0")
        edge.set(qn("w:color"), _hex(color, "0F4761"))

    for name in ("top", "left", "bottom", "right"):
        set_edge(name, theme.outer_border_color, theme.outer_border_pt)
    for name in ("insideH", "insideV"):
        set_edge(name, theme.inner_border_color, theme.inner_border_pt)


def _set_cell_text_preserve_geometry(
    cell, text: str, *, style: TextStyle, center: bool = False
) -> None:
    """Replace text while leaving the original cell margins/row geometry untouched."""
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
        _apply_run_style(p.runs[0], style)
        for r in list(p.runs[1:]):
            r._element.getparent().remove(r._element)
    else:
        r = p.add_run(text or "")
        _apply_run_style(r, style)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _style_header_row(row, table_theme: TableTheme) -> None:
    labels = ["No.", "Word", "Type", "Pronunciation", "Meaning"]
    for i, (cell, label) in enumerate(zip(row.cells, labels)):
        _set_cell_shading(cell, table_theme.header_no_fill if i == 0 else table_theme.header_fill)
        _set_cell_no_wrap(cell, True)
        _set_cell_text_preserve_geometry(
            cell, label, style=table_theme.header_text, center=True
        )


def _populate_vocab_table(doc: Document, document: VocabDocument, table_theme: TableTheme) -> None:
    if len(doc.tables) < 2:
        raise RuntimeError("Master template phải có title table và VOCAB prototype table.")
    table = doc.tables[1]
    if len(table.rows) < 3:
        raise RuntimeError("VOCAB prototype table cần header + 2 prototype rows.")

    table.autofit = False
    _set_fixed_table_layout_and_no_column(table, 0.50)
    header = table.rows[0]
    _set_repeat_header(header)
    _set_cant_split(header)
    _style_header_row(header, table_theme)
    _set_table_borders(table, table_theme)

    root_proto = table.rows[1]
    derivative_proto = table.rows[2]
    merge_groups = []

    for family in document.families:
        family_rows = []
        for word_index, item in enumerate(family.words):
            source = root_proto if word_index == 0 else derivative_proto
            table._tbl.append(deepcopy(source._tr))
            row = table.rows[-1]
            _set_cant_split(row)
            for c in row.cells:
                _clear_vmerge(c)
                # Body remains clean white; only the table header carries strong brand color.
                _set_cell_shading(c, "FFFFFF")

            # IMPORTANT: the first word in a family is NOT bold/orange.
            # All body rows use the exact same style.
            _set_cell_text_preserve_geometry(
                row.cells[0], family.number if word_index == 0 else "",
                style=table_theme.body_text, center=True,
            )
            _set_cell_text_preserve_geometry(row.cells[1], item.word, style=table_theme.body_text)
            _set_cell_text_preserve_geometry(row.cells[2], item.word_type, style=table_theme.body_text)
            _set_cell_text_preserve_geometry(row.cells[3], item.pronunciation, style=table_theme.body_text)
            _set_cell_text_preserve_geometry(row.cells[4], item.meaning, style=table_theme.body_text)
            family_rows.append(row)

        for i, row in enumerate(family_rows):
            _set_family_keep_next(row, i < len(family_rows) - 1)
        if len(family_rows) > 1:
            merge_groups.append(family_rows)

    table._tbl.remove(root_proto._tr)
    table._tbl.remove(derivative_proto._tr)

    # One single continuous table. Apply vMerge only after all body text is final.
    for family_rows in merge_groups:
        _set_vmerge_tc(family_rows[0]._tr.tc_lst[0], restart=True)
        for row in family_rows[1:]:
            _set_vmerge_tc(row._tr.tc_lst[0], restart=False)


def render_vocab_docx(
    document: VocabDocument,
    *,
    template_path: str | Path,
    theme: ThemeConfig | None = None,
) -> io.BytesIO:
    theme = theme or default_theme()
    doc = Document(str(template_path))
    _set_header(doc, theme)
    _set_title_clearance(doc, theme)
    _set_title_block(doc, document, theme)
    _populate_vocab_table(doc, document, theme.table)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ===================== VALIDATION =====================
@dataclass(slots=True)
class ValidationMessage:
    level: str  # error | warning
    message: str


def validate_vocab(document: VocabDocument) -> list[ValidationMessage]:
    messages: list[ValidationMessage] = []

    if not document.unit:
        messages.append(ValidationMessage("error", "Unit không được để trống."))
    if not document.title:
        messages.append(ValidationMessage("error", "Title không được để trống."))
    if not document.families:
        messages.append(ValidationMessage("error", "Không tìm thấy dữ liệu từ vựng hợp lệ."))
        return messages

    seen_numbers: set[str] = set()
    seen_words: dict[str, int] = {}

    for family_index, family in enumerate(document.families, start=1):
        if not family.number:
            messages.append(
                ValidationMessage(
                    "error",
                    f"Family #{family_index} không có STT. Dòng đầu tiên phải có giá trị ở cột No.",
                )
            )
        elif family.number in seen_numbers:
            messages.append(
                ValidationMessage("warning", f"STT {family.number} xuất hiện nhiều hơn một lần.")
            )
        seen_numbers.add(family.number)

        if not family.words:
            messages.append(ValidationMessage("error", f"Family {family.number} không có từ."))

        for item in family.words:
            key = item.word.casefold()
            seen_words[key] = seen_words.get(key, 0) + 1
            if not item.word_type:
                messages.append(
                    ValidationMessage(
                        "warning", f"Từ “{item.word}” chưa có Type/part of speech."
                    )
                )

    duplicates = sorted(word for word, count in seen_words.items() if count > 1)
    if duplicates:
        preview = ", ".join(duplicates[:8])
        suffix = "…" if len(duplicates) > 8 else ""
        messages.append(ValidationMessage("warning", f"Có từ lặp: {preview}{suffix}"))

    return messages


def has_errors(messages: list[ValidationMessage]) -> bool:
    return any(m.level == "error" for m in messages)

# ===================== QA =====================
@dataclass(slots=True)
class QAResult:
    ok: bool
    checks: list[str]
    problems: list[str]


def _docx_text_fallback(path: Path) -> str:
    doc = Document(str(path))
    parts: list[str] = []
    parts.extend(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(parts)


def extract_docx_markdown(path: str | Path) -> str:
    """Use Microsoft's MarkItDown when installed; fall back locally for resilience."""
    path = Path(path)
    try:
        from markitdown import MarkItDown  # type: ignore

        result = MarkItDown().convert(str(path))
        return result.text_content or ""
    except Exception:
        return _docx_text_fallback(path)


def run_content_and_brand_qa(path: str | Path, expected: VocabDocument) -> QAResult:
    path = Path(path)
    checks: list[str] = []
    problems: list[str] = []

    text = extract_docx_markdown(path)
    expected_title_tokens = [expected.document_type, f"UNIT {expected.display_heading_unit}", expected.title]
    for token in expected_title_tokens:
        if token.casefold() in text.casefold():
            checks.append(f"Title token OK: {token}")
        else:
            problems.append(f"Thiếu title token: {token}")

    for family in expected.families:
        for item in family.words:
            if item.word.casefold() not in text.casefold():
                problems.append(f"Thiếu từ sau khi render: {item.word}")
    if not any(p.startswith("Thiếu từ") for p in problems):
        checks.append(f"Đủ {expected.word_count} vocabulary rows về mặt semantic.")

    # OOXML brand checks: these are structural, not visual.
    with ZipFile(path) as zf:
        names = set(zf.namelist())
        document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
        styles_xml = zf.read("word/styles.xml").decode("utf-8", errors="ignore")
        header_xml = "\n".join(
            zf.read(n).decode("utf-8", errors="ignore")
            for n in names
            if n.startswith("word/header") and n.endswith(".xml")
        )
        footer_xml = "\n".join(
            zf.read(n).decode("utf-8", errors="ignore")
            for n in names
            if n.startswith("word/footer") and n.endswith(".xml")
        )

    for token, label in [
        ("BF4E14", "orange BF4E14"),
        ("0F4761", "navy 0F4761"),
        ("Arial", "Arial"),
        ("Orbitron", "Orbitron"),
    ]:
        haystack = document_xml + styles_xml + header_xml + footer_xml
        if token in haystack:
            checks.append(f"Brand token OK: {label}")
        else:
            problems.append(f"Brand token không còn trong DOCX: {label}")

    if "Lingual" in header_xml or "image" in header_xml or "imagedata" in header_xml:
        checks.append("Header/logo-watermark structure còn tồn tại.")
    else:
        problems.append("Không phát hiện cấu trúc logo/watermark ở header.")

    if "PAGE" in footer_xml and "0345286842" in footer_xml:
        checks.append("Footer + PAGE field + contact OK.")
    else:
        problems.append("Footer mất PAGE field hoặc contact.")

    return QAResult(ok=not problems, checks=checks, problems=problems)

# ===================== STREAMLIT UI =====================
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "lx_vocab_template.docx"

st.set_page_config(page_title="LingualXplore · Vocab Builder", page_icon="📘", layout="wide")
st.title("LingualXplore · VOCAB Builder V3")
st.caption(
    "V3: title không đè header · header thêm/bớt từ tự do · bảng là MỘT table liền mạch, "
    "giữ geometry gốc nhưng được skin theo theme LingualXplore."
)

# -----------------------------------------------------------------------------
# DOCUMENT CONTENT
# -----------------------------------------------------------------------------
with st.container(border=True):
    c1, c2, c3, c4 = st.columns([1, 1, 1.5, 2.4])
    with c1:
        unit = st.text_input("Sub-unit / badge", value="5.1")
    with c2:
        heading_unit = st.text_input(
            "Unit ở heading", value="5",
            help="Badge có thể là 5.1 nhưng heading hiển thị UNIT 5: SPECIAL DAYS."
        )
    with c3:
        document_type = st.text_input("Document type", value="VOCAB BUILDER")
    with c4:
        title = st.text_input("Title", value="SPECIAL DAYS")

    filename = st.text_input("Tên file tải về", value="VOCAB BUILDER UNIT 5.1 - SPECIAL DAYS.docx")
    if not filename.lower().endswith(".docx"):
        filename += ".docx"

# -----------------------------------------------------------------------------
# TITLE THEME
# -----------------------------------------------------------------------------
with st.expander("🎨 Title block", expanded=False):
    a, b, c = st.columns(3)
    with a:
        st.markdown("**Badge 5.1**")
        badge_font = st.text_input("Font badge", "Orbitron", key="badge_font")
        badge_size = st.number_input("Size badge", 20.0, 80.0, 48.0, 1.0, key="badge_size")
        badge_color = st.color_picker("Màu badge", "#FFFFFF", key="badge_color")
        badge_bold = st.checkbox("Bold badge", True, key="badge_bold")
    with b:
        st.markdown("**VOCAB BUILDER**")
        type_font = st.text_input("Font document type", "Monument Extended", key="type_font")
        type_size = st.number_input("Size document type", 10.0, 40.0, 18.0, 0.5, key="type_size")
        type_color = st.color_picker("Màu document type", "#BF4E14", key="type_color")
        type_bold = st.checkbox("Bold document type", False, key="type_bold")
    with c:
        st.markdown("**UNIT 5: SPECIAL DAYS**")
        heading_font = st.text_input("Font heading", "Arial", key="heading_font")
        heading_size = st.number_input("Size heading", 14.0, 50.0, 27.0, 0.5, key="heading_size")
        heading_color = st.color_picker("Màu heading", "#000000", key="heading_color")
        heading_bold = st.checkbox("Bold heading", True, key="heading_bold")

    title_clearance = st.number_input(
        "Khoảng cách an toàn Header → Title (pt)", 18.0, 60.0, 28.0, 1.0,
        help="V3 đã sửa lỗi title đè header. Nếu bạn dùng header 3 dòng hoặc font rất lớn, tăng giá trị này."
    )

# -----------------------------------------------------------------------------
# FLEXIBLE HEADER
# -----------------------------------------------------------------------------
HEADER_DEFAULTS = [
    (1, "Supplementary", "Arial", 10.5, "#E7A07F", False, False),
    (1, "Vocabulary", "Arial", 10.5, "#E7A07F", False, False),
    (2, "for", "Arial", 10.5, "#E7A07F", False, False),
    (2, "KET", "Orbitron", 10.5, "#7FA9BC", True, False),
    (2, "Learners", "Orbitron", 10.5, "#7FA9BC", True, False),
    (2, "in", "Arial", 10.5, "#E7A07F", False, False),
    (2, "Orwell", "Orbitron", 10.5, "#7FA9BC", True, False),
    (2, "Classes", "Orbitron", 10.5, "#7FA9BC", True, False),
]

with st.expander("✍️ Header linh hoạt: thêm/bớt từ, font, màu từng segment", expanded=True):
    st.caption(
        "Mỗi segment có thể là một từ hoặc cả một cụm. Tăng/giảm 'Số segment' để thêm/bớt. "
        "Để bỏ một segment ở giữa, để Text trống. Bạn có thể dùng tối đa 3 dòng."
    )
    h1, h2, h3 = st.columns(3)
    with h1:
        align1 = st.selectbox("Căn dòng 1", ["left", "center", "right"], index=2)
    with h2:
        align2 = st.selectbox("Căn dòng 2", ["left", "center", "right"], index=2)
    with h3:
        align3 = st.selectbox("Căn dòng 3", ["left", "center", "right"], index=2)

    segment_count = int(st.number_input("Số segment", 1, 16, len(HEADER_DEFAULTS), 1))
    st.markdown("**Line · Text · Font · Color · Size · Bold · Italic**")
    header_values = []
    for i in range(segment_count):
        default = HEADER_DEFAULTS[i] if i < len(HEADER_DEFAULTS) else (2, "", "Arial", 10.5, "#E7A07F", False, False)
        line_d, text_d, font_d, size_d, color_d, bold_d, italic_d = default
        cols = st.columns([0.65, 2.0, 1.7, 1.05, 0.85, 0.55, 0.55])
        with cols[0]:
            line = st.selectbox("Line", [1, 2, 3], index=line_d - 1, key=f"h_line_{i}", label_visibility="collapsed")
        with cols[1]:
            text = st.text_input("Text", text_d, key=f"h_text_{i}", label_visibility="collapsed")
        with cols[2]:
            font = st.text_input("Font", font_d, key=f"h_font_{i}", label_visibility="collapsed")
        with cols[3]:
            color = st.color_picker("Color", color_d, key=f"h_color_{i}", label_visibility="collapsed")
        with cols[4]:
            size = st.number_input("Size", 6.0, 30.0, size_d, 0.5, key=f"h_size_{i}", label_visibility="collapsed")
        with cols[5]:
            bold = st.checkbox("B", bold_d, key=f"h_bold_{i}")
        with cols[6]:
            italic = st.checkbox("I", italic_d, key=f"h_italic_{i}")
        header_values.append((line, text, font, size, color, bold, italic))

# -----------------------------------------------------------------------------
# TABLE THEME — geometry stays from original file
# -----------------------------------------------------------------------------
with st.expander("🧩 Table theme — giữ dãn ô gốc, chỉ đổi skin", expanded=True):
    st.caption(
        "Row height/cell padding/tỷ lệ cột vẫn lấy từ file VOCAB trước app. "
        "Các tuỳ chọn dưới đây chỉ đổi màu, font và border."
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Header**")
        table_no_fill = st.color_picker("Fill ô No.", "#BF4E14")
        table_header_fill = st.color_picker("Fill 4 ô còn lại", "#0F4761")
        table_header_font = st.text_input("Font header bảng", "Arial")
        table_header_size = st.number_input("Size header bảng", 8.0, 18.0, 10.5, 0.5)
        table_header_color = st.color_picker("Màu chữ header", "#FFFFFF")
        table_header_bold = st.checkbox("Bold header bảng", True)
    with c2:
        st.markdown("**Body**")
        body_font = st.text_input("Font body", "Times New Roman")
        body_size = st.number_input("Size body", 9.0, 16.0, 12.0, 0.5)
        body_color = st.color_picker("Màu chữ body", "#000000")
        st.caption("Từ đầu word family dùng cùng style với mọi từ khác: không bold, không cam.")
    with c3:
        st.markdown("**Borders**")
        outer_border_color = st.color_picker("Outer border", "#0F4761")
        outer_border_pt = st.number_input("Outer border (pt)", 0.25, 2.0, 1.0, 0.25)
        inner_border_color = st.color_picker("Inner grid", "#B8C8D0")
        inner_border_pt = st.number_input("Inner grid (pt)", 0.25, 1.5, 0.5, 0.25)

st.info(
    "CSV 5 cột: No., Word, Type, Pronunciation, Meaning. No. trống = cùng word family. "
    "Output vẫn là MỘT bảng duy nhất; header bảng tự lặp khi sang trang."
)
raw_data = st.text_area(
    "Vocabulary data", height=360,
    placeholder='No.,Word,Type,Pronunciation,Meaning\n1,special,"adj, n",,\n,specially,adv,,',
)
run_qa = st.checkbox("Chạy Content + Brand QA", value=True)

if st.button("KHỞI TẠO TÀI LIỆU", type="primary", use_container_width=True):
    if not TEMPLATE_PATH.exists():
        st.error(f"Không tìm thấy template: {TEMPLATE_PATH}")
        st.stop()

    document = parse_vocab_csv(
        raw_data, unit=unit, heading_unit=heading_unit, title=title, document_type=document_type
    )
    messages = validate_vocab(document)
    for message in messages:
        (st.error if message.level == "error" else st.warning)(message.message)
    if has_errors(messages):
        st.stop()

    segments = [
        HeaderSegment(
            line=int(line),
            text=text,
            style=TextStyle(
                font=font, size=float(size), color=color.lstrip("#"),
                bold=bool(bold), italic=bool(italic),
            ),
        )
        for line, text, font, size, color, bold, italic in header_values
        if text.strip()
    ]

    theme = ThemeConfig(
        badge=TextStyle(badge_font, badge_size, badge_color.lstrip("#"), badge_bold),
        document_type=TextStyle(type_font, type_size, type_color.lstrip("#"), type_bold),
        heading=TextStyle(heading_font, heading_size, heading_color.lstrip("#"), heading_bold),
        header_segments=segments,
        header_alignments={1: align1, 2: align2, 3: align3},
        title_clearance_pt=title_clearance,
        table=TableTheme(
            header_no_fill=table_no_fill.lstrip("#"),
            header_fill=table_header_fill.lstrip("#"),
            header_text=TextStyle(
                table_header_font, table_header_size, table_header_color.lstrip("#"), table_header_bold
            ),
            body_text=TextStyle(body_font, body_size, body_color.lstrip("#"), False, False),
            outer_border_color=outer_border_color.lstrip("#"),
            outer_border_pt=outer_border_pt,
            inner_border_color=inner_border_color.lstrip("#"),
            inner_border_pt=inner_border_pt,
        ),
    )

    st.write(f"Đã nhận **{document.family_count} word families / {document.word_count} từ**.")
    try:
        output = render_vocab_docx(document, template_path=TEMPLATE_PATH, theme=theme)
    except Exception as exc:
        st.exception(exc)
        st.stop()

    if run_qa:
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp.write(output.getvalue())
            tmp_path = Path(tmp.name)
        try:
            result = run_content_and_brand_qa(tmp_path, document)
            if result.ok:
                st.success("QA semantic + brand structure: PASS")
            else:
                st.warning("QA có cảnh báo; file vẫn được tạo để bạn kiểm tra trực quan trong Word.")
            with st.expander("Chi tiết QA"):
                for item in result.checks:
                    st.write("✅", item)
                for item in result.problems:
                    st.write("❌", item)
        finally:
            tmp_path.unlink(missing_ok=True)

    st.download_button(
        "TẢI FILE WORD", output.getvalue(), file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary", use_container_width=True,
    )
