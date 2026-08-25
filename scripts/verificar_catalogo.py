"""Auditoría de formatos funcionales, referencias y gráficos Office."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
COLORS = {"4C2B46", "E97700"}

# secciones, tablas XML, cantidad de gráficos Office, fila dinámica, términos obligatorios
REQUIREMENTS = {
    "STGND_Informe_Maestro_v9.docx": (2, 2, 3, False, ["Control documental", "HALLAZGOS", "Gráfico de torta"]),
    "STGND_Informe_Fondo_Color_Maestro_v9.docx": (2, 2, 3, False, ["Control documental", "HALLAZGOS", "Gráfico de torta"]),
    "STGND_Minuta_Actividad_Maestro_v9.docx": (1, 2, 0, True, ["MINUTA DE ACTIVIDAD", "Acuerdos y seguimiento", "Síntesis"]),
    "STGND_Informativo_Normativo_Maestro_v9.docx": (2, 2, 0, False, ["Control normativo", "Ámbito de aplicación", "Vigencia"]),
    "STGND_Oficio_Maestro_v9.docx": (1, 2, 0, False, ["OFICIO N.°", "MAT.", "REF.", "FIRMA ELECTRÓNICA", "Distribución"]),
    "STGND_Oficio_Circular_Maestro_v9.docx": (1, 2, 0, False, ["OFICIO CIRCULAR N.°", "MAT.", "REF.", "FIRMA ELECTRÓNICA", "Distribución"]),
    "STGND_Agenda_Reunion_Maestro_v9.docx": (1, 2, 0, True, ["AGENDA DE REUNIÓN", "Objetivo de la reunión", "Documentos preparatorios"]),
    "STGND_Acta_Reunion_Maestro_v9.docx": (1, 3, 0, True, ["ACTA", "ACUERDOS Y SEGUIMIENTO", "PRÓXIMA REUNIÓN", "FIRMA ELECTRÓNICA"]),
    "STGND_Programa_Actividad_Maestro_v9.docx": (1, 1, 0, True, ["PROGRAMA", "Persona Expositora", "Tiempo"]),
    "STGND_Checklist_Actividad_Maestro_v9.docx": (2, 3, 0, True, ["Control de tareas", "Observaciones"]),
    "STGND_Invitacion_Email_Maestro_v9.docx": (1, 3, 0, False, ["Invitación institucional", "{{cuerpo_invitacion}}", "{{cta}}"]),
    "STGND_Ficha_Solicitud_Compras_Maestro_v9.docx": (2, 2, 0, False, ["Solicitud de compra o contratación", "Ítem presupuestario", "ID PAC"]),
    "STGND_Documento_Graficos_Maestro_v9.docx": (2, 4, 10, False, ["Ficha de datos", "Archivo relacionado", "Gráfico de torta", "Flujo de decisión"]),
}

SIMPLE = {
    "STGND_Oficio_Maestro_v9.docx", "STGND_Oficio_Circular_Maestro_v9.docx",
    "STGND_Agenda_Reunion_Maestro_v9.docx", "STGND_Acta_Reunion_Maestro_v9.docx",
    "STGND_Minuta_Actividad_Maestro_v9.docx", "STGND_Programa_Actividad_Maestro_v9.docx",
    "STGND_Invitacion_Email_Maestro_v9.docx",
}
FORMAL_CORRESPONDENCE = {"STGND_Oficio_Maestro_v9.docx", "STGND_Oficio_Circular_Maestro_v9.docx"}

errors: list[str] = []
checks = 0


def check(condition: bool, message: str) -> None:
    global checks
    checks += 1
    if not condition:
        errors.append(message)


def qn(name: str) -> str:
    return f"{{{W}}}{name}"


def verify_docx(path: Path) -> None:
    expected_sections, expected_tables, expected_charts, expects_dynamic, terms = REQUIREMENTS[path.name]
    with zipfile.ZipFile(path) as package:
        check(package.testzip() is None, f"{path.name}: paquete dañado")
        names = package.namelist()
        check(len(names) == len(set(names)), f"{path.name}: contiene partes OOXML duplicadas")
        check("word/document.xml" in names and "word/styles.xml" in names, f"{path.name}: faltan partes Word")
        document = package.read("word/document.xml")
        root = ET.fromstring(document)
        text = " | ".join(node.text or "" for node in root.iter(qn("t")))
        sections = len(root.findall(".//w:sectPr", NS))
        tables = len(root.findall(".//w:tbl", NS))
        charts = [name for name in names if name.startswith("word/charts/") and name.endswith(".xml") and "/_rels/" not in name]
        check(sections == expected_sections, f"{path.name}: secciones {sections} != {expected_sections}")
        check(tables == expected_tables, f"{path.name}: tablas {tables} != {expected_tables}")
        check(len(charts) == expected_charts, f"{path.name}: gráficos Office {len(charts)} != {expected_charts}")
        check(("{{dynamic_rows}}" in text) == expects_dynamic, f"{path.name}: fila dinámica incorrecta")
        for term in terms:
            check(term in text, f"{path.name}: falta {term}")
        if expected_sections == 1:
            check("Título de la sección" not in text and "Portadilla" not in text, f"{path.name}: conserva secciones editoriales innecesarias")
            check("HALLAZGOS" not in text, f"{path.name}: conserva hallazgos innecesarios")

        styles = ET.fromstring(package.read("word/styles.xml"))
        normal = styles.find(".//w:style[@w:styleId='Normal']", NS)
        size = normal.find("w:rPr/w:sz", NS) if normal is not None else None
        fonts = normal.find("w:rPr/w:rFonts", NS) if normal is not None else None
        check(size is not None and size.get(qn("val")) == "22", f"{path.name}: Normal no usa 11 pt")
        check(fonts is not None and fonts.get(qn("ascii")) == "Aptos", f"{path.name}: Normal no usa Aptos")
        # Office localiza los styleId integrados (p. ej. Ttulo1), pero conserva
        # los nombres semánticos OOXML; la comparación ignora mayúsculas.
        style_by_name = {}
        for node in styles.findall(".//w:style", NS):
            name = node.find("w:name", NS)
            if name is not None and name.get(qn("val")):
                style_by_name[name.get(qn("val")).casefold()] = node
        used_styles = {
            node.get(qn("val"))
            for node in root.findall(".//w:pPr/w:pStyle", NS)
            if node.get(qn("val"))
        }
        if path.name in FORMAL_CORRESPONDENCE:
            editorial_ids = {
                node.get(qn("styleId")) for name, node in style_by_name.items()
                if name in {"title", "heading 1", "heading 2", "heading 3"}
            }
            check(not (used_styles & editorial_ids), f"{path.name}: usa estilos editoriales improcedentes")
            colors = [node.get(qn("val"), "").upper() for node in root.findall(".//w:rPr/w:color", NS)]
            check(colors and all(color in {"000000", "AUTO"} for color in colors), f"{path.name}: contiene texto de color no formal")
        else:
            for semantic in ("title", "heading 1", "heading 2", "heading 3", "normal", "caption"):
                check(semantic in style_by_name, f"{path.name}: falta estilo semántico {semantic}")
            title_id = style_by_name["title"].get(qn("styleId"))
            check(title_id in used_styles, f"{path.name}: ningún párrafo usa el estilo Title")
            check(style_by_name["normal"].get(qn("default")) == "1", f"{path.name}: Normal no es el estilo de párrafo predeterminado")
            for level, style_name in enumerate(("heading 1", "heading 2", "heading 3")):
                node = style_by_name.get(style_name)
                outline = node.find("w:pPr/w:outlineLvl", NS) if node is not None else None
                check(outline is not None and outline.get(qn("val")) == str(level), f"{path.name}: {style_name} sin nivel TOC")
        check(b"{{" in document, f"{path.name}: no contiene campos configurables")
        simulated = re.sub(rb"\{\{\s*[^}]+?\s*\}\}", b"Texto guia", document)
        check(b"{{" not in simulated, f"{path.name}: sustitución simulada deja marcadores")
        if "FIRMA ELECTRÓNICA" in text:
            floating_signature = any("FIRMADIGITAL" in "".join(node.itertext()) for node in root.iter() if node.tag.endswith("}anchor"))
            check(not floating_signature and text.count("FIRMADIGITAL") == 1, f"{path.name}: firma electrónica flotante o duplicada")

        if path.name not in SIMPLE:
            upper = b"".join(package.read(name) for name in names if name.endswith(".xml")).decode("utf-8", "ignore").upper()
            for color in COLORS:
                check(color in upper, f"{path.name}: falta color institucional {color}")


def verify_presentation() -> None:
    path = ROOT / "formatos/pptx/STGND_Presentacion_Maestro_v9.pptx"
    with zipfile.ZipFile(path) as package:
        check(package.testzip() is None, f"{path.name}: paquete dañado")
        names = package.namelist()
        check(len(names) == len(set(names)), f"{path.name}: contiene partes OOXML duplicadas")
        charts = [name for name in names if name.startswith("ppt/charts/chart") and name.endswith(".xml")]
        check(len(charts) >= 4, f"{path.name}: faltan ejemplos de gráficos Office")
        all_xml = b"".join(package.read(name) for name in names if name.endswith(".xml"))
        check("OPCIÓN DE PORTADA · MORADA" in all_xml.decode("utf-8", "ignore"), f"{path.name}: falta portada morada")
        check("OPCIÓN DE PORTADA · BLANCA" in all_xml.decode("utf-8", "ignore"), f"{path.name}: falta portada blanca")
        white_logo = (ROOT / "assets/logo-stgnd-blanco.png").read_bytes()
        media = [package.read(name) for name in names if name.startswith("ppt/media/")]
        check(b"Gracias" in all_xml and white_logo in media, f"{path.name}: cierre sin logotipo blanco de contraste")
        rotated = 0
        for name in names:
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                root = ET.fromstring(package.read(name))
                rotated += sum(1 for node in root.iter() if node.tag.endswith("}xfrm") and node.get("rot"))
        check(rotated == 0, f"{path.name}: conserva {rotated} objetos rotados")


def verify_workbook() -> None:
    path = ROOT / "formatos/xlsx/STGND_Datos_Graficos_Maestro_v9.xlsx"
    with zipfile.ZipFile(path) as package:
        check(package.testzip() is None, f"{path.name}: paquete dañado")
        check(len(package.namelist()) == len(set(package.namelist())), f"{path.name}: contiene partes OOXML duplicadas")
        charts = [name for name in package.namelist() if name.startswith("xl/charts/") and name.endswith(".xml")]
        check(len(charts) >= 10, f"{path.name}: faltan ejemplos de gráficos Excel nativos")
        check(b"barChart" in b"".join(package.read(name) for name in charts), f"{path.name}: no contiene gráfico de Office")


def main() -> int:
    schema = json.loads((ROOT / "schema/catalogo_documental_v9.json").read_text(encoding="utf-8"))
    documents = schema["documents"]
    check(schema["version"] == "9.3-biblioteca-visual", "Versión de esquema inesperada")
    check(len(documents) == 14, "El esquema debe registrar 14 formatos")
    check(len({doc["id"] for doc in documents}) == len(documents), "Hay identificadores duplicados")
    for doc in documents:
        check((ROOT / doc["file"]).is_file(), f"No existe {doc['file']}")

    paths = sorted((ROOT / "formatos/docx_maestro").glob("*.docx"))
    check(len(paths) == len(REQUIREMENTS) == 13, "Deben existir 13 DOCX, incluidas dos portadas de informe")
    check({path.name for path in paths} == set(REQUIREMENTS), "La carpeta DOCX contiene archivos no controlados")
    for path in paths:
        verify_docx(path)
    verify_presentation()
    verify_workbook()

    app = (ROOT / "assets/js/app.js").read_text(encoding="utf-8")
    html = (ROOT / "index.html").read_text(encoding="utf-8")
    for invariant in ("templateVariants", "portada_informe", "headerImage", "imagen_cabecera", "replaceHeaderImage", "expandDynamicRows", "expandSections", "unresolved.length", "ciudad_upper"):
        check(invariant in app, f"El gestor no implementa {invariant}")
    for doc_id, media_path in (("programa", "word/media/image1.png"), ("email", "word/media/image1.png")):
        definition = re.search(r"\{id:'" + doc_id + r"'.+?\},", app)
        check(definition is not None and f"headerImage:'{media_path}'" in definition.group(0), f"{doc_id}: imagen de cabecera no configurable")
        template = next(doc for doc in documents if doc["id"] == doc_id)
        with zipfile.ZipFile(ROOT / template["file"]) as package:
            check(media_path in package.namelist(), f"{doc_id}: falta el recurso de cabecera {media_path}")
    for simple_id in ("minuta", "agenda", "acta", "programa"):
        definition = re.search(r"\{id:'" + re.escape(simple_id) + r"'.+?\},", app)
        check(definition is not None and "sections:true" not in definition.group(0), f"{simple_id}: expone secciones improcedentes")
    check("formatos/docx/" not in app, "El gestor referencia plantillas v8")
    check('lang="es"' in html and 'charset="utf-8"' in html, "HTML sin idioma o UTF-8")
    for reference in ("STGND - Oficio.zip", "ProgramaHito_lanzamiento.docx", "Invitacion_Hito_Guia.docx"):
        check((ROOT / "base" / reference).is_file(), f"Falta referencia {reference}")

    if errors:
        print(f"FALLO: {len(errors)} de {checks} comprobaciones")
        for error in errors:
            print(f"- {error}")
        return 1
    print(f"OK: {checks} comprobaciones; 13 DOCX funcionales y gráficos Office nativos verificados.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
