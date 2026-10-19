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
    if unit == "5.1":
        # Use pedagogically realistic content for the screenshot shown in-app.
        rows = [
            "No.,Word,Type,Pronunciation,Meaning",
            "1,impress,v,/ɪmˈpres/,gây ấn tượng",
            ",impressive,adj,/ɪmˈpresɪv/,ấn tượng; nổi bật",
            "2,design,v,/dɪˈzaɪn/,thiết kế",
            ",designer,n,/dɪˈzaɪnə(r)/,nhà thiết kế",
            "3,protect,v,/prəˈtekt/,bảo vệ",
            ",protective,adj,/prəˈtektɪv/,có tính bảo vệ",
            "4,develop,v,/dɪˈveləp/,phát triển",
            ",development,n,/dɪˈveləpmənt/,sự phát triển",
            "5,harmony,n,/ˈhɑːməni/,sự hòa hợp",
            ",harmonious,adj,/hɑːˈməʊniəs/,hài hòa",
            "6,society,n,/səˈsaɪəti/,xã hội",
            ",social,adj,/ˈsəʊʃl/,thuộc về xã hội",
        ]
    if unit == "10.1":
        rows = ["No.,Word,Type,Pronunciation,Meaning"]
        for i in range(85):
            rows.append(("1" if i == 0 else "") + f",continuous{i},n,/kənˈtɪnjʊəs/,Nghĩa từ vựng thứ {i}")
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
    ('ONE VERY LONG WORD FAMILY','10.1',85),
    ('FOOD, HEALTH & CULTURE','12.1',5),
    ('EXTRAORDINARY BIOLOGICAL DIVERSITY ACROSS CONTINENTS','13.1',6),
    ('THIS-IS-A-VERY-LONG-UNBREAKABLE-WORD-IN-THE-TITLE','14.1',6),
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
    if unit in ("7.1", "12.1", "13.1", "14.1"):
        first = pdf[0]
        words = first.get_text("words")
        table_headers = [w for w in words if w[4] == "No." and w[1] > 100]
        if not table_headers:
            raise AssertionError(f"Missing vocabulary table for {unit}")
        table_start = table_headers[0][1]
        title_words = [w for w in words if w[0] >= 165 and 117 < w[1] < table_start - 4
                       and w[4].upper() not in ("VOCAB", "BUILDER")]
        title_lines = {}
        for w in title_words:
            key = round(w[1] / 4) * 4
            title_lines.setdefault(key, []).append(w)
            if w[2] > first.rect.width - 44:
                raise AssertionError(f"Title crosses right safe margin: {w} in {unit}")
        title_content = " ".join(w[4] for group in title_lines.values() for w in group)
        print("TITLE-REGRESSION", unit, repr(title_content), "line count", len(title_lines),
              "table start", table_start)
        if "PLACES" in title.upper() and unit == "7.1" and len(title_lines) != 1:
            raise AssertionError("PEOPLE & PLACES wraps into multiple lines")
        if not " ".join(title.upper().split()).replace(" ", "") in (
                "".join(w[4] for group in title_lines.values() for w in group).upper()):
            raise AssertionError(f"Title text missing or clipped on page for unit {unit}")

    if unit == '10.1':
        if '10.1' not in pdf[0].get_text():
            raise AssertionError('The 10.1 badge wrapped instead of staying on one line')
        for i, page in enumerate(pdf):
            if i >= len(pdf)-1:
                break
            text_rows = [
                w for w in page.get_text("words")
                if (w[4].startswith("continuous") or w[4].startswith("Nghĩa"))
            ]
            if not text_rows:
                raise AssertionError(f'Word family missing on PDF page {i+1}')
            last_text_bottom = max(w[3] for w in text_rows)
            print('FAMILY-CONTINUATION-PAGE',i+1,'body bottom',round(last_text_bottom,1))
            if last_text_bottom < 650:
                raise AssertionError(f'Wasted content space on page {i+1}: last row ends at {last_text_bottom:.1f}')

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


# Verify the *actual PDF*, not only XML, for each contact visibility mode.
for show_phone, show_email in ((False, False), (False, True), (True, False)):
    theme=N['default_theme']()
    theme.footer.show_phone=show_phone
    theme.footer.show_email=show_email
    test_doc=N['parse_vocab_csv'](
        '1,protect,v,/prəˈtekt/,bảo vệ\n,protection,n,/prəˈtekʃn/,sự bảo vệ',
        unit='1.1',title='HIDDEN CONTACT TEST',
    )
    tag=f"phone-{int(show_phone)}-email-{int(show_email)}"
    filename=DEST/(tag+'.docx')
    filename.write_bytes(N['render_vocab_docx'](
        test_doc,template_path=MASTER,theme=theme,
    ).getvalue())
    proc=subprocess.run(
        ['libreoffice','-env:UserInstallation=file:///tmp/taowl_lo_studio',
         '--headless','--convert-to','pdf:writer_pdf_Export','--outdir',str(DEST),str(filename)],
        capture_output=True,text=True,timeout=120,
    )
    if proc.returncode:
        raise RuntimeError(f"Hidden-contact render failed: {proc.stderr}")
    pdf=fitz.open(filename.with_suffix('.pdf'))
    text=pdf[0].get_text()
    print("FOOTER-VISIBILITY",tag,"phone",("0345286842" in text),"email",("email@yourcenter.com" in text))
    if ("0345286842" in text) != show_phone or ("email@yourcenter.com" in text) != show_email:
        raise AssertionError("Contact visibility in rendered PDF does not match settings")
    if not show_phone and not show_email:
        page=pdf[0]
        footer_region=fitz.Rect(0,page.rect.height-115,page.rect.width,page.rect.height)
        page.get_pixmap(matrix=fitz.Matrix(2,2),clip=footer_region,alpha=False).save(DEST/"footer-without-contacts.png")
