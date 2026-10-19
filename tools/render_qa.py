"""Render real DOCX outputs to PDF/PNG and report font/geometry details.

Runs in GitHub Actions against the exact branded master DOCX. The PDFs and
screenshots are released as a CI artifact for visual quality review.
"""
from __future__ import annotations
import ast, io, subprocess, sys, types
from pathlib import Path
import fitz

ROOT=Path(__file__).resolve().parents[1]
DEST=ROOT/"qa_output"
DEST.mkdir(exist_ok=True)
source=(ROOT/"app.py").read_text(encoding="utf-8")
cut=source.index("# ===================== STREAMLIT UI =====================")
module=types.ModuleType("taowl_render_visual")
sys.modules[module.__name__]=module
exec(compile(source[:cut],str(ROOT/"app.py"),"exec"),module.__dict__)
N=module.__dict__
MASTER=ROOT/"templates/lx_vocab_template.docx"

def make(title: str, unit: str, count: int) -> str:
    rows=['No.,Word,Type,Pronunciation,Meaning']
    for i in range(1,count+1):
        rows.append(f'{i},vocabulary{i},n,/vəˈkæbjələri/,"Nghĩa tiếng Việt và phần giải thích dài số {i}"')
        rows.append(f',vocabulary{i}ly,adv,/ˈvɒkjʊ/,"Một cách tự nhiên, rõ ràng và phù hợp với ví dụ {i}"')
    doc=N['parse_vocab_csv']('\n'.join(rows),unit=unit,title=title,heading_unit=unit.split('.')[0])
    data=N['render_vocab_docx'](doc,template_path=MASTER).getvalue()
    path=DEST/(f'{unit.replace(".","_")}.docx')
    path.write_bytes(data)
    return path

for title,unit,count in [
    ('SPECIAL DAYS','5.1',6),
    ('PEOPLE & PLACES','7.1',8),
    ('SPECIAL DAYS AND CULTURAL FESTIVALS AROUND THE WORLD','8.2',12),
    ('WILDLIFE AND ENVIRONMENT','9.3',85),
]:
    path=make(title,unit,count)
    proc=subprocess.run(
        ['libreoffice','-env:UserInstallation=file:///tmp/taowl_lo_studio',
         '--headless','--convert-to','pdf:writer_pdf_Export','--outdir',str(DEST),str(path)],
        capture_output=True,text=True,timeout=120,
    )
    print('SOFFICE',unit,proc.returncode,proc.stdout.strip(),proc.stderr.strip()[-450:])
    if proc.returncode:
        raise RuntimeError(f'LibreOffice failed for {path.name}: {proc.stderr}')
    pdf=fitz.open(path.with_suffix('.pdf'))
    print('PDF',unit,'pages',len(pdf),'page size',pdf[0].rect,'expected rows',count*2)
    for i,page in enumerate(pdf):
        words=page.get_text("words")
        if not words:
            raise AssertionError(f'Blank page {i+1} in {unit}')
        footer_words=[w for w in words if w[3] > page.rect.height - 65]
        visible_page=[w for w in footer_words if w[4] == str(i+1)]
        if not visible_page:
            raise AssertionError(f'PAGE field {i+1} missing in {unit}')
        bbox=visible_page[-1]
        field_center=(bbox[0]+bbox[2])/2
        page_center=page.rect.width/2
        print(' FOOTER-PAGE',unit,i+1,'center',round(field_center,2),'expected',round(page_center,2))
        if abs(field_center-page_center) > 7:
            raise AssertionError(f'PAGE field displaced {field_center-page_center:.1f}pt in {unit}')
        if not any('email@yourcenter.com' in w[4] for w in footer_words):
            raise AssertionError(f'Missing footer email in {unit} page {i+1}')
        if not any('0345286842' in w[4] for w in footer_words):
            raise AssertionError(f'Missing footer phone in {unit} page {i+1}')
        if i==0:
            print(' PAGE-ONE-TEXT',repr(page.get_text()[:400]))
            for label in ['VOCAB','UNIT','PLACES','Learners','Explorers','0345286842']:
                hits=page.search_for(label)
                print('  bbox',label,[(round(x.x0),round(x.y0),round(x.x1),round(x.y1)) for x in hits[:5]])
            pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
            pix.save(DEST/f'{unit.replace(".","_")}-page1.png')
            foot=fitz.Rect(0,page.rect.height-110,page.rect.width,page.rect.height)
            page.get_pixmap(matrix=fitz.Matrix(2,2),clip=foot,alpha=False).save(DEST/f'{unit.replace(".","_")}-footer.png')
    if count>15:
        last=pdf[-1]
        last.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False).save(DEST/f'{unit.replace(".","_")}-last.png')
