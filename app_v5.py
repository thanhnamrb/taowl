from __future__ import annotations

import html
from pathlib import Path

import pandas as pd
import streamlit as st

from v5_core import Theme, TextStyle, csv_to_rows, inspect_layout, render_docx, rows_to_document


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "lx_vocab_template.docx"
COLUMNS = ["No.", "Word", "Type", "Pronunciation", "Meaning"]


def _seed_frame() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["1", "special", "adj, n", "/ˈspeʃ.əl/", "đặc biệt"],
            ["", "specially", "adv", "/ˈspeʃ.əl.i/", "một cách đặc biệt"],
            ["2", "celebrate", "v", "/ˈsel.ə.breɪt/", "kỷ niệm"],
            ["", "celebration", "n", "/ˌsel.əˈbreɪ.ʃən/", "sự kỷ niệm"],
            ["3", "decorate", "v", "/ˈdek.ə.reɪt/", "trang trí"],
        ],
        columns=COLUMNS,
    )


def _ensure_state() -> None:
    defaults = {
        "v5_unit": "5.1",
        "v5_heading_unit": "5",
        "v5_section": "UNIT",
        "v5_document_type": "VOCAB BUILDER",
        "v5_title": "SPECIAL DAYS",
        "v5_filename": "VOCAB BUILDER UNIT 5.1 - SPECIAL DAYS V5.docx",
        "v5_phone": "0345286842",
        "v5_email": "email@yourcenter.com",
        "v5_body_font": "Times New Roman",
        "v5_body_size": 12.0,
        "v5_heading_font": "Montserrat",
        "v5_heading_size": 27.0,
        "v5_accent": "#BF4E14",
        "v5_navy": "#0F4761",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)
    if "v5_seed_rows" not in st.session_state:
        st.session_state.v5_seed_rows = _seed_frame()


def _rows_from_df(df: pd.DataFrame) -> list[dict]:
    clean = df.fillna("")
    return [
        {col: str(row.get(col, "")) for col in COLUMNS}
        for _, row in clean.iterrows()
        if str(row.get("Word", "")).strip()
    ]


def _preview_html(data, theme: Theme) -> str:
    flat_rows = []
    for family in data.families:
        for i, word in enumerate(family.words):
            flat_rows.append(
                (
                    family.number if i == 0 else "",
                    word.word,
                    word.word_type,
                    word.pronunciation,
                    word.meaning,
                )
            )

    rows_per_page = 15
    chunks = [flat_rows[i:i + rows_per_page] for i in range(0, len(flat_rows), rows_per_page)] or [[]]

    pages = []
    for page_no, chunk in enumerate(chunks, start=1):
        body = []
        for no, word, typ, pron, meaning in chunk:
            body.append(
                "<tr>"
                f"<td class='no'>{html.escape(no)}</td>"
                f"<td>{html.escape(word)}</td>"
                f"<td>{html.escape(typ)}</td>"
                f"<td>{html.escape(pron)}</td>"
                f"<td>{html.escape(meaning)}</td>"
                "</tr>"
            )

        page = f"""
        <section class="paper">
          <div class="paper-head">
            <div class="brand-small">Supplementary Vocabulary for KET Learners</div>
          </div>

          <div class="title-grid">
            <div class="badge">{html.escape(data.unit_badge)}</div>
            <div class="title-copy">
              <div class="doc-type">{html.escape(data.document_type)}</div>
              <div class="main-title">{html.escape(data.section_label)} {html.escape(data.display_heading_unit)}: {html.escape(data.title.upper())}</div>
            </div>
          </div>

          <table class="vocab">
            <colgroup>
              <col style="width:7.5%">
              <col style="width:23.5%">
              <col style="width:15.5%">
              <col style="width:21.5%">
              <col style="width:32%">
            </colgroup>
            <thead>
              <tr><th>No.</th><th>Word</th><th>Type</th><th>Pronunciation</th><th>Meaning</th></tr>
            </thead>
            <tbody>{''.join(body)}</tbody>
          </table>

          <div class="footer">
            <span>From Learners to Explorers</span>
            <span class="page-badge">{page_no}</span>
            <span class="contact">☎ {html.escape(theme.phone)}<br>✉ {html.escape(theme.email)}</span>
          </div>
        </section>
        """
        pages.append(page)

    return f"""
    <style>
      .preview-wrap {{
        background: #e8edf0;
        padding: 24px 12px;
        border-radius: 16px;
      }}
      .paper {{
        width: min(100%, 794px);
        aspect-ratio: 210 / 297;
        min-height: 1050px;
        margin: 0 auto 24px auto;
        background: white;
        box-shadow: 0 8px 28px rgba(18, 39, 49, .14);
        padding: 42px 48px 48px 48px;
        box-sizing: border-box;
        position: relative;
        font-family: Arial, sans-serif;
        color: #111;
        overflow: hidden;
      }}
      .paper-head {{
        min-height: 28px;
        display: flex;
        justify-content: flex-end;
        align-items: start;
        margin-bottom: 18px;
      }}
      .brand-small {{
        font-size: 12px;
        color: #7FA9BC;
        font-weight: 700;
      }}
      .title-grid {{
        display: grid;
        grid-template-columns: 88px 1fr;
        gap: 20px;
        align-items: center;
        margin-bottom: 22px;
      }}
      .badge {{
        background: {html.escape(theme.accent)};
        color: white;
        min-height: 82px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 4px;
        font-size: 34px;
        font-weight: 800;
      }}
      .doc-type {{
        color: {html.escape(theme.accent)};
        font-size: 18px;
        margin-bottom: 4px;
      }}
      .main-title {{
        font-size: 27px;
        line-height: 1.05;
        font-weight: 800;
      }}
      .vocab {{
        border-collapse: collapse;
        width: 100%;
        table-layout: fixed;
        font-family: "Times New Roman", serif;
        font-size: 13px;
      }}
      .vocab th {{
        background: {html.escape(theme.navy)};
        color: white;
        font-family: Arial, sans-serif;
        font-size: 12px;
        padding: 7px 5px;
        border: 1px solid {html.escape(theme.navy)};
      }}
      .vocab td {{
        border: 1px solid {html.escape(theme.navy)};
        padding: 6px 6px;
        vertical-align: middle;
        overflow-wrap: anywhere;
      }}
      .vocab td.no {{
        background: {html.escape(theme.accent)};
        color: white;
        font-family: Arial, sans-serif;
        text-align: center;
        font-weight: 700;
      }}
      .footer {{
        position: absolute;
        left: 48px;
        right: 48px;
        bottom: 24px;
        display: grid;
        grid-template-columns: 1fr 40px 1fr;
        align-items: center;
        font-size: 11px;
        color: {html.escape(theme.navy)};
      }}
      .page-badge {{
        width: 28px;
        height: 28px;
        border-radius: 50%;
        margin: auto;
        display: flex;
        align-items: center;
        justify-content: center;
        background: {html.escape(theme.accent)};
        color: white;
        font-weight: 800;
      }}
      .contact {{
        text-align: right;
      }}
      @media (max-width: 850px) {{
        .paper {{
          min-height: 820px;
          padding: 28px 30px 40px 30px;
        }}
        .footer {{left: 30px; right: 30px;}}
      }}
    </style>
    <div class="preview-wrap">{''.join(pages)}</div>
    """


def main() -> None:
    st.set_page_config(
        page_title="LingualXplore · Vocab Builder V5 Preview",
        page_icon="🧪",
        layout="wide",
    )
    _ensure_state()

    st.markdown(
        """
        <style>
          .block-container {max-width: 1500px; padding-top: 1rem;}
          [data-testid="stSidebar"] {background: #F7F9FA;}
          .v5-banner {
            padding: 10px 14px; border-radius: 12px;
            background: #FFF4EE; border: 1px solid #F2C8B2;
            color: #6C2A0B; margin-bottom: 12px;
          }
        </style>
        <div class="v5-banner"><b>V5 Preview</b> · renderer mới, fixed-width table, pagination ổn định hơn, page badge không còn floating anchor.</div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.markdown("## Tài liệu")
        st.selectbox("MODULE / UNIT", ["UNIT", "MODULE"], key="v5_section")
        st.text_input("Sub-unit / badge", key="v5_unit")
        st.text_input("Số ở heading", key="v5_heading_unit")
        st.text_input("Document type", key="v5_document_type")
        st.text_input("Title", key="v5_title")
        st.text_input("Tên file", key="v5_filename")

        st.divider()
        st.markdown("## Theme nhanh")
        st.color_picker("Accent", key="v5_accent")
        st.color_picker("Navy", key="v5_navy")
        st.selectbox(
            "Body font",
            ["Times New Roman", "Arial", "Aptos", "Calibri", "Cambria"],
            key="v5_body_font",
        )
        st.number_input("Body size", 9.0, 16.0, 0.5, key="v5_body_size")
        st.selectbox(
            "Heading font",
            ["Montserrat", "Arial", "Aptos", "Calibri"],
            key="v5_heading_font",
        )
        st.number_input("Heading size", 18.0, 36.0, 1.0, key="v5_heading_size")
        st.text_input("Điện thoại", key="v5_phone")
        st.text_input("Email", key="v5_email")

    theme = Theme(
        accent=st.session_state.v5_accent.lstrip("#"),
        navy=st.session_state.v5_navy.lstrip("#"),
        body=TextStyle(st.session_state.v5_body_font, float(st.session_state.v5_body_size), "000000"),
        heading=TextStyle(st.session_state.v5_heading_font, float(st.session_state.v5_heading_size), "000000", True),
        phone=st.session_state.v5_phone,
        email=st.session_state.v5_email,
    )

    tab_edit, tab_preview, tab_export = st.tabs(["✏️ Editor", "👁️ Preview A4", "⬇️ Xuất Word"])

    with tab_edit:
        st.markdown("### Editor dạng bảng")
        st.caption("Thêm/xóa dòng trực tiếp. Để trống No. nếu dòng đó thuộc cùng word family với dòng phía trên.")
        edited = st.data_editor(
            st.session_state.v5_seed_rows,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "No.": st.column_config.TextColumn("No.", width="small"),
                "Word": st.column_config.TextColumn("Word", width="medium", required=True),
                "Type": st.column_config.TextColumn("Type", width="medium"),
                "Pronunciation": st.column_config.TextColumn("Pronunciation", width="medium"),
                "Meaning": st.column_config.TextColumn("Meaning", width="large"),
            },
            key="v5_table_editor",
        )

        with st.expander("Dán CSV cũ"):
            raw = st.text_area(
                "CSV 5 cột",
                height=180,
                placeholder="No.,Word,Type,Pronunciation,Meaning",
                key="v5_csv_import",
            )
            if st.button("Nạp CSV vào editor", use_container_width=True):
                rows = csv_to_rows(raw)
                if rows:
                    st.session_state.v5_seed_rows = pd.DataFrame(rows, columns=COLUMNS)
                    st.session_state.pop("v5_table_editor", None)
                    st.rerun()
                else:
                    st.warning("Không tìm thấy dòng dữ liệu hợp lệ.")

    rows = _rows_from_df(edited)
    data = rows_to_document(
        rows,
        unit=st.session_state.v5_unit,
        title=st.session_state.v5_title,
        heading_unit=st.session_state.v5_heading_unit,
        section_label=st.session_state.v5_section,
        document_type=st.session_state.v5_document_type,
    )

    with tab_preview:
        c1, c2 = st.columns(2)
        c1.metric("Word families", data.family_count)
        c2.metric("Vocabulary rows", data.word_count)
        st.caption("Preview này ưu tiên thao tác nhanh. File DOCX vẫn là nguồn chuẩn cuối cùng cho pagination của Microsoft Word.")
        st.markdown(_preview_html(data, theme), unsafe_allow_html=True)

    with tab_export:
        st.markdown("### Xuất bằng renderer V5")
        st.write(
            "V5 đặt bảng theo đúng printable width của section, không dùng keep-with-next cho cả family, "
            "không vertical-merge cột No., và page badge là inline drawing ở giữa footer."
        )

        if not TEMPLATE_PATH.exists():
            st.error(f"Không tìm thấy template: {TEMPLATE_PATH}")
            return
        if not data.title.strip():
            st.error("Title đang trống.")
            return
        if not data.families:
            st.warning("Hãy nhập ít nhất một từ trước khi xuất.")
            return

        try:
            docx_bytes = render_docx(data, template_path=TEMPLATE_PATH, theme=theme)
            layout_problems = inspect_layout(docx_bytes)
        except Exception as exc:
            st.exception(exc)
            return

        if layout_problems:
            st.warning("Layout QA còn cảnh báo:")
            for item in layout_problems:
                st.write("•", item)
        else:
            st.success("Layout QA: PASS — fixed table / no keepNext / inline PAGE badge.")

        filename = st.session_state.v5_filename.strip() or "VOCAB_BUILDER_V5.docx"
        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        st.download_button(
            "TẢI FILE WORD V5",
            docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

        with st.expander("V5 đang sửa gì?"):
            st.markdown(
                """
                - Table width = page width - left margin - right margin.
                - w:tblLayout cố định; từng gridCol và tcW dùng cùng một geometry.
                - Không vMerge cột No. qua nhiều row, tránh tương tác xấu với page break.
                - Page badge dùng wp:inline, không còn offset âm theo paragraph.
                """
            )


if __name__ == "__main__":
    main()
