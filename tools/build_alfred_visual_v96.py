from pathlib import Path
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from zipfile import ZipFile, ZIP_DEFLATED
import shutil
import tempfile

ROOT=Path(__file__).resolve().parents[1]
SOURCE=ROOT/'base'/'STGND - Oficio.zip'
ARCHIVE=ROOT/'base'/'originales'/'STGND - Oficio_original.zip'
PURPLE='4C2B46'; ORANGE='E97700'; GREY='686868'; LIGHT='F4F1F3'; BLACK='222222'

def set_font(run,name='Aptos',size=11,color=BLACK,bold=None):
    run.font.name=name
    run.font.size=Pt(size)
    run.font.color.rgb=RGBColor.from_string(color)
    if bold is not None: run.bold=bold
    rpr=run._r.get_or_add_rPr()
    rfonts=rpr.rFonts
    if rfonts is None:
        rfonts=OxmlElement('w:rFonts'); rpr.insert(0,rfonts)
    for attr in ['ascii','hAnsi','eastAsia','cs']:
        rfonts.set(qn('w:'+attr),name)

def p_border(p,side='bottom',color=ORANGE,size=8,space=3):
    pPr=p._p.get_or_add_pPr()
    pBdr=pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr=OxmlElement('w:pBdr'); pPr.append(pBdr)
    old=pBdr.find(qn('w:'+side))
    if old is not None: pBdr.remove(old)
    b=OxmlElement('w:'+side)
    b.set(qn('w:val'),'single'); b.set(qn('w:sz'),str(size))
    b.set(qn('w:space'),str(space)); b.set(qn('w:color'),color)
    pBdr.append(b)

def set_cell_shading(cell,fill):
    tcPr=cell._tc.get_or_add_tcPr()
    shd=tcPr.find(qn('w:shd'))
    if shd is None:
        shd=OxmlElement('w:shd'); tcPr.append(shd)
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),fill)

def set_cell_border(cell,**edges):
    tcPr=cell._tc.get_or_add_tcPr()
    borders=tcPr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders=OxmlElement('w:tcBorders'); tcPr.append(borders)
    for edge,opts in edges.items():
        tag='w:'+edge; el=borders.find(qn(tag))
        if el is None: el=OxmlElement(tag); borders.append(el)
        for k,v in opts.items(): el.set(qn('w:'+k),str(v))

def set_margins(doc,top,bottom,left,right):
    for s in doc.sections:
        s.top_margin=Mm(top); s.bottom_margin=Mm(bottom)
        s.left_margin=Mm(left); s.right_margin=Mm(right)
        s.header_distance=Mm(6); s.footer_distance=Mm(6)

def style_all_runs(doc):
    for p in doc.paragraphs:
        for r in p.runs: set_font(r,'Aptos',11,BLACK)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs: set_font(r,'Aptos',10.5,BLACK)

def style_heading(p,size=11):
    for r in p.runs: set_font(r,'Aptos',size,PURPLE,True)
    p.paragraph_format.space_before=Pt(8)
    p.paragraph_format.space_after=Pt(4)
    p_border(p,'bottom',ORANGE,6,2)

def style_meta(p,size=10.5):
    for r in p.runs: set_font(r,'Aptos',size,BLACK)
    p.paragraph_format.space_after=Pt(2)

def style_logo_paras(doc):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p._p.xpath('.//w:drawing'):
                        p.paragraph_format.space_after=Pt(3)
                        p_border(p,'bottom',ORANGE,5,2)
                        return

def improve_oficio(doc):
    # Horizontal margins stay at 30 mm because official tab stops depend on them.
    # Only vertical whitespace and typography are tightened to keep one-page output.
    set_margins(doc,12,12,30,30)
    style_all_runs(doc); style_logo_paras(doc)
    ps=doc.paragraphs
    for i,p in enumerate(ps):
        txt=p.text.strip()
        if i<=5:
            for r in p.runs:
                set_font(r,'Aptos',10.5,PURPLE if txt.startswith(('A','DE')) else BLACK,True)
            p.paragraph_format.space_after=Pt(1.5)
        elif 'Sin otro particular' in txt or 'Agradeciendo' in txt:
            for r in p.runs: set_font(r,'Aptos',11,BLACK)
            p.paragraph_format.space_before=Pt(6)
        elif txt.startswith('Distribución') or txt.startswith('- '):
            for r in p.runs: set_font(r,'Aptos',9,GREY)
            p.paragraph_format.space_after=Pt(0)
        else:
            for r in p.runs: set_font(r,'Aptos',11,BLACK)
            p.paragraph_format.line_spacing=1.05
            p.paragraph_format.space_after=Pt(3)
    if doc.tables:
        t=doc.tables[0]
        for ri,row in enumerate(t.rows):
            for ci,cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    for r in p.runs:
                        set_font(r,'Aptos',10.5,PURPLE if ci==0 else BLACK,ci==0)
            if ri==0:
                for cell in row.cells:
                    set_cell_border(cell,bottom={'val':'single','sz':'6','color':ORANGE,'space':'0'})

def improve_agenda(doc):
    set_margins(doc,11,11,16,16); style_all_runs(doc)
    ps=doc.paragraphs
    if len(ps)>1:
        for r in ps[1].runs: set_font(r,'Aptos Display',16,PURPLE,False)
        ps[1].alignment=WD_ALIGN_PARAGRAPH.CENTER
        p_border(ps[1],'bottom',ORANGE,7,3)
    for idx in [7,9,14]:
        if idx<len(ps): style_heading(ps[idx],11)
    for i,p in enumerate(ps):
        if i not in [1,7,9,14]: style_meta(p,10.5 if i<7 else 11)

def improve_acta(doc):
    set_margins(doc,10,10,14,14); style_all_runs(doc)
    ps=doc.paragraphs
    if ps:
        for r in ps[0].runs: set_font(r,'Aptos Display',15,PURPLE,False)
        ps[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
        p_border(ps[0],'bottom',ORANGE,7,3)
    for p in ps:
        txt=p.text.strip().upper()
        if txt in {'OBJETIVOS DE LA REUNIÓN:','AGENDA DE LA REUNION','DOCUMENTOS PREPARATORIOS','DESARROLLO DE LA REUNIÓN:','ACUERDOS','APROBACIÓN DEL ACTA'}:
            style_heading(p,10.5)
        else:
            style_meta(p,10)
    for t in doc.tables:
        for ri,row in enumerate(t.rows):
            for cell in row.cells:
                if ri==0:
                    set_cell_shading(cell,PURPLE)
                    for p in cell.paragraphs:
                        for r in p.runs: set_font(r,'Aptos',8.5,'FFFFFF',True)
                else:
                    for p in cell.paragraphs:
                        for r in p.runs: set_font(r,'Aptos',9,BLACK)
            if ri==0:
                for cell in row.cells:
                    set_cell_border(cell,bottom={'val':'single','sz':'8','color':ORANGE,'space':'0'})

def improve_informe(doc):
    set_margins(doc,12,12,16,16); style_all_runs(doc)
    for p in doc.paragraphs:
        txt=p.text.strip()
        if txt in {'Antecedentes','Introducción','Análisis','Consideraciones Finales.'}:
            for r in p.runs: set_font(r,'Aptos',11,PURPLE,True)
        elif txt.lower().startswith('contenido'):
            for r in p.runs: set_font(r,'Aptos Display',15,PURPLE,False)
            p_border(p,'bottom',ORANGE,7,3)
        elif txt:
            for r in p.runs: set_font(r,'Aptos',11,BLACK)
            p.paragraph_format.line_spacing=1.08
            p.paragraph_format.space_after=Pt(4)

def improve_minuta(doc):
    set_margins(doc,12,12,16,16); style_all_runs(doc)
    ps=doc.paragraphs
    for i,p in enumerate(ps):
        txt=p.text.strip()
        if i in (0,13,21):
            for r in p.runs: set_font(r,'Aptos Display',15,PURPLE,False)
            p_border(p,'bottom',ORANGE,7,3)
            p.paragraph_format.space_before=Pt(4)
            p.paragraph_format.space_after=Pt(7)
        elif ':' in txt and i<18:
            for r in p.runs: set_font(r,'Aptos',10.5,BLACK)
            p.paragraph_format.space_after=Pt(2)
        elif txt:
            for r in p.runs: set_font(r,'Aptos',11,BLACK)
            p.paragraph_format.line_spacing=1.08
            p.paragraph_format.space_after=Pt(4)

def soften_signature_placeholder(path):
    tmp=path.with_suffix('.tmp.docx')
    with ZipFile(path,'r') as zin, ZipFile(tmp,'w',ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename.endswith('.xml'):
                try:
                    text=data.decode('utf-8')
                    # Preserve geometry/object type. Only replace legacy Office colors.
                    text=text.replace('729FCF','F4F1F3').replace('3465A4','D8CFD6').replace('E36C0A','E97700')
                    data=text.encode('utf-8')
                except UnicodeDecodeError:
                    pass
            zout.writestr(item,data)
    tmp.replace(path)

def main():
    if not SOURCE.exists():
        raise SystemExit(f'Missing source: {SOURCE}')
    ARCHIVE.parent.mkdir(parents=True,exist_ok=True)
    if not ARCHIVE.exists():
        shutil.copy2(SOURCE,ARCHIVE)

    with tempfile.TemporaryDirectory() as td:
        td=Path(td); src=td/'src'; out=td/'out'
        src.mkdir(); out.mkdir()
        with ZipFile(SOURCE) as z: z.extractall(src)

        specs=[
            ('STGND - Oficio.docx',improve_oficio),
            ('STGND - Oficio circular.docx',improve_oficio),
            ('STGND - Agenda.docx',improve_agenda),
            ('STGND - Acta.docx',improve_acta),
            ('STGND - Informe.docx',improve_informe),
            ('STGND - Minuta.docx',improve_minuta),
        ]
        for name,fn in specs:
            doc=Document(src/name); fn(doc); doc.save(out/name); soften_signature_placeholder(out/name)

        purchase='Ficha de solicitud de compras.docx'
        if (src/purchase).exists():
            shutil.copy2(src/purchase,out/purchase)

        tmpzip=td/'alfred.zip'
        with ZipFile(tmpzip,'w',ZIP_DEFLATED) as z:
            for p in sorted(out.glob('*.docx')): z.write(p,p.name)
        shutil.copy2(tmpzip,SOURCE)
        print(f'Updated {SOURCE}')
        print(f'Original preserved at {ARCHIVE}')

if __name__=='__main__':
    main()
