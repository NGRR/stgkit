from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import base64, shutil, tempfile

ROOT=Path(__file__).resolve().parents[1]
DIST=ROOT/'dist'
BUILD=ROOT/'build'/'com_stgdocumentos_0.2.4'
SKELETON=ROOT/'tools'/'stgdocumentos_024_skeleton.b64'

CATALOG_FILES=[
    'STGND_Informe_Maestro_v9.docx',
    'STGND_Minuta_Actividad_Maestro_v9.docx',
    'STGND_Oficio_Maestro_v9.docx',
    'STGND_Oficio_Circular_Maestro_v9.docx',
    'STGND_Programa_Actividad_Maestro_v9.docx',
    'STGND_Checklist_Actividad_Maestro_v9.docx',
    'STGND_Invitacion_Email_Maestro_v9.docx',
    'STGND_Documento_Graficos_Maestro_v9.docx',
    'STGND_Datos_Graficos_Maestro_v10.xlsx',
    'STGND_Presentacion_v3_2026.pptx',
]
SOURCE_MAP={
    'STGND_Informe_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Informe_Maestro_v9.docx',
    'STGND_Minuta_Actividad_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Minuta_Actividad_Maestro_v9.docx',
    'STGND_Oficio_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Oficio_Maestro_v9.docx',
    'STGND_Oficio_Circular_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Oficio_Circular_Maestro_v9.docx',
    'STGND_Programa_Actividad_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Programa_Actividad_Maestro_v9.docx',
    'STGND_Checklist_Actividad_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Checklist_Actividad_Maestro_v9.docx',
    'STGND_Invitacion_Email_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Invitacion_Email_Maestro_v9.docx',
    'STGND_Documento_Graficos_Maestro_v9.docx': ROOT/'formatos/docx_maestro/STGND_Documento_Graficos_Maestro_v9.docx',
    'STGND_Datos_Graficos_Maestro_v10.xlsx': ROOT/'formatos/xlsx/STGND_Datos_Graficos_Maestro_v10.xlsx',
    'STGND_Presentacion_v3_2026.pptx': ROOT/'formatos/pptx/STGND_Presentacion_v3_2026.pptx',
}
ALFRED_NAMES=[
    'STGND - Oficio.docx','STGND - Oficio circular.docx','STGND - Agenda.docx',
    'STGND - Acta.docx','STGND - Informe.docx','STGND - Minuta.docx'
]

def main():
    if BUILD.exists(): shutil.rmtree(BUILD)
    BUILD.mkdir(parents=True)
    raw=base64.b64decode(SKELETON.read_text(encoding='ascii'))
    with tempfile.NamedTemporaryFile(suffix='.zip') as tmp:
        tmp.write(raw); tmp.flush()
        with ZipFile(tmp.name) as z: z.extractall(BUILD)

    catalog=BUILD/'admin/masters/catalog'; alfred=BUILD/'admin/masters/alfred'
    catalog.mkdir(parents=True,exist_ok=True); alfred.mkdir(parents=True,exist_ok=True)
    for name,src in SOURCE_MAP.items():
        if not src.exists(): raise SystemExit(f'Missing source {src}')
        shutil.copy2(src,catalog/name)

    with tempfile.TemporaryDirectory() as td:
        td=Path(td)
        with ZipFile(ROOT/'base/STGND - Oficio.zip') as z: z.extractall(td)
        for name in ALFRED_NAMES:
            src=td/name
            if not src.exists(): raise SystemExit(f'Missing Alfred entry {name}')
            shutil.copy2(src,alfred/name)
        shutil.copy2(td/'STGND - Agenda.docx',catalog/'STGND - Agenda.docx')
        shutil.copy2(td/'STGND - Acta.docx',catalog/'STGND - Acta.docx')

    logo=ROOT/'assets/logo-stgnd.png'
    if not logo.exists(): raise SystemExit('Missing assets/logo-stgnd.png')
    (BUILD/'media/images').mkdir(parents=True,exist_ok=True)
    shutil.copy2(logo,BUILD/'media/images/logo-stgnd.png')

    expected=12+6
    masters=list((BUILD/'admin/masters').rglob('*.*'))
    if len(masters)!=expected:
        raise SystemExit(f'Expected {expected} master files, found {len(masters)}')

    DIST.mkdir(exist_ok=True)
    out=DIST/'com_stgdocumentos_0.2.4.zip'
    if out.exists(): out.unlink()
    with ZipFile(out,'w',ZIP_DEFLATED) as z:
        for p in sorted(BUILD.rglob('*')):
            if p.is_file(): z.write(p,p.relative_to(BUILD).as_posix())
    print(out)
    print('masters',len(masters),'bytes',out.stat().st_size)

if __name__=='__main__':
    main()
