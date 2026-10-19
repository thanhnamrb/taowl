"""Inspect original LingualXplore DOCX geometry and generated output in CI."""
from pathlib import Path
import io
import zipfile
import re
import importlib.util
from docx import Document
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[1]
def table_info(table, label):
    print(f"{label}: rows={len(table.rows)}, cols={len(table.columns)}, autofit={table.autofit}")
    print("grid inches", [round(c.width.inches,3) if c.width else None for c in table.columns])
    for i,row in enumerate(table.rows[:3]):
        print(" row",i,"height",row.height.pt if row.height else None,"rule",row.height_rule,
            "cells",[(str(c.text)[:70],round(c.width.inches,2) if c.width else None) for c in row.cells])

d=Document(ROOT/"templates"/"lx_vocab_template.docx")
s=d.sections[0]
print("TEMPLATE SECTION", {a: round(getattr(s,a).inches,3) if getattr(s,a) else None for a in ("page_width","page_height","top_margin","bottom_margin","left_margin","right_margin","header_distance","footer_distance")})
print("TEMPLATE BODY tables",len(d.tables),"paragraphs",len(d.paragraphs))
for i,t in enumerate(d.tables[:2]):table_info(t,f"TEMPLATE TABLE {i}")
print("TEMPLATE HEADER",len(s.header.paragraphs),"texts",[p.text for p in s.header.paragraphs])
print("TEMPLATE HEADER drawing elements",len(s.header._element.xpath(".//w:drawing")),"pict",len(s.header._element.xpath(".//w:pict")))
print("TEMPLATE FOOTER tables",len(s.footer.tables),"text",[p.text for p in s.footer.paragraphs])
for i,t in enumerate(s.footer.tables):table_info(t,f"TEMPLATE FOOTER TABLE {i}")
with zipfile.ZipFile(ROOT/"templates"/"lx_vocab_template.docx") as z:
    print("TEMPLATE MEDIA",[(n,len(z.read(n))) for n in z.namelist() if n.startswith("word/media/")])
    print("TEMPLATE PAGE fields", [n for n in z.namelist() if n.startswith("word/footer") and n.endswith(".xml")])
