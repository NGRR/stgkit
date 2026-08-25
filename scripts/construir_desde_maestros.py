"""Construye cada DOCX como variante directa de los maestros de ``base``.

La topología OOXML y todos los recursos del maestro permanecen intactos. Sólo se
editan texto, metadatos y el estilo Normal para cuerpo Aptos 11.
"""
from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "base"
OUT = ROOT / "formatos" / "docx_maestro"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
NS = {"w": W, "cp": CP, "dc": DC}


@dataclass(frozen=True)
class Document:
    filename: str
    master: str
    kind: str
    theme: str
    subtitle: str
    summary: str
    callout: str
    context: str
    section_title: str
    section_summary: str
    section_callout: str
    section_body: str
    rows: bool = False
    sections: bool = False


DOCS = [
    Document("STGND_Informe_Maestro_v9.docx", "DocumentoMaestro2.docx", "Informe", "{{tema}}", "{{bajada}}", "{{resumen}}", "{{idea_clave}}", "{{cuerpo}}", "Título de la sección", "{{resumen_seccion}}", "{{idea_clave_seccion}}", "{{texto_seccion}}", sections=True),
    Document("STGND_Informe_Fondo_Color_Maestro_v9.docx", "DocumentoMaestro.docx", "Informe", "{{tema}}", "{{bajada}}", "{{resumen}}", "{{idea_clave}}", "{{cuerpo}}", "Título de la sección", "{{resumen_seccion}}", "{{idea_clave_seccion}}", "{{texto_seccion}}", sections=True),
    Document("STGND_Minuta_Actividad_Maestro_v9.docx", "DocumentoMaestro2.docx", "Minuta", "ACTIVIDAD", "Documento de coordinación, registro y seguimiento de actividad.", "{{resumen}}", "{{idea_clave}}", "{{cuerpo}}", "Título de la sección", "{{resumen_seccion}}", "{{idea_clave_seccion}}", "{{texto_seccion}}", sections=True),
    Document("STGND_Informativo_Normativo_Maestro_v9.docx", "DocumentoMaestro2.docx", "Informativo", "NORMATIVA", "Normas, orientaciones y criterios de funcionamiento institucional.", "{{resumen}}", "{{idea_clave}}", "{{cuerpo}}", "Título de la sección", "{{resumen_seccion}}", "{{idea_clave_seccion}}", "{{texto_seccion}}", sections=True),
    Document("STGND_Oficio_Maestro_v9.docx", "DocumentoMaestro.docx", "Oficio", "CORRESPONDENCIA", "Comunicación institucional formal.", "{{cuerpo}}", "Destinatario: {{destinatario}} · {{cargo_destinatario}}", "Institución: {{institucion_destinataria}} · Responsable: {{responsable}}", "Datos del oficio", "Código {{codigo}} · Fecha {{fecha}} · Estado {{estado}}", "Documento preparado para revisión y firma institucional.", "Complete referencias, anexos, distribución y firma cuando corresponda."),
    Document("STGND_Oficio_Circular_Maestro_v9.docx", "DocumentoMaestro.docx", "Oficio circular", "CORRESPONDENCIA", "Comunicación institucional dirigida a múltiples destinatarios.", "{{cuerpo}}", "Destinatarios: {{destinatarios}}", "Responsable: {{responsable}} · Estado: {{estado}}", "Distribución y control", "{{destinatarios}}", "Código {{codigo}} · Fecha {{fecha}}", "Complete referencias, anexos y distribución interna."),
    Document("STGND_Agenda_Reunion_Maestro_v9.docx", "DocumentoMaestro.docx", "Agenda", "REUNIONES", "Planificación de sesión, reunión o mesa de trabajo.", "{{objetivo}}", "Responsable: {{responsable}}", "Fecha {{fecha}} · Estado {{estado}}", "Programa de la reunión", "Bloques horarios, temas y responsables.", "Código {{codigo}}", "{{dynamic_rows}}", rows=True),
    Document("STGND_Acta_Reunion_Maestro_v9.docx", "DocumentoMaestro.docx", "Acta", "REUNIONES", "Registro de sesión, acuerdos, responsables y plazos.", "{{desarrollo}}", "Responsable del acta: {{responsable}}", "Fecha {{fecha}} · Estado {{estado}}", "Acuerdos y seguimiento", "Registro editable de decisiones y responsables.", "Código {{codigo}}", "{{dynamic_rows}}", rows=True),
    Document("STGND_Programa_Actividad_Maestro_v9.docx", "DocumentoMaestro.docx", "Programa", "ACTIVIDADES", "Programa horario para jornadas, talleres e hitos institucionales.", "{{descripcion_actividad}}", "Responsable: {{responsable}}", "Fecha {{fecha}} · Estado {{estado}}", "Programa de la actividad", "Horario, contenido, responsable y duración.", "Código {{codigo}}", "{{dynamic_rows}}", rows=True),
    Document("STGND_Checklist_Actividad_Maestro_v9.docx", "DocumentoMaestro.docx", "Checklist", "ACTIVIDADES", "Control operativo antes, durante y después de la actividad.", "Marque cada tarea y complete responsable, plazo y observaciones.", "Responsable general: {{responsable}}", "Fecha de control: {{fecha}} · Estado {{estado}}", "Control de tareas", "Registro editable para el seguimiento operativo.", "Código {{codigo}}", "{{dynamic_rows}}", rows=True),
    Document("STGND_Invitacion_Email_Maestro_v9.docx", "DocumentoMaestro.docx", "Invitación", "COMUNICACIONES", "Pieza base para comunicación institucional por correo electrónico.", "{{cuerpo_invitacion}}", "Asunto: {{asunto_email}}", "Preencabezado: {{preheader}}", "Llamado a la acción", "{{cta}}", "Enlace de destino", "{{url_cta}}"),
    Document("STGND_Ficha_Solicitud_Compras_Maestro_v9.docx", "DocumentoMaestro.docx", "Ficha", "ADMINISTRACIÓN", "Antecedentes para solicitud de contratación o adquisición.", "{{motivo_contratacion}}", "Unidad requirente: {{unidad_requirente}}", "Monto estimado: {{monto_estimado}}", "Datos de la solicitud", "Convenio marco: {{convenio_marco}}", "Identificación y control", "ID PAC: {{id_pac}} · Responsable: {{responsable}} · Fecha: {{fecha}}"),
    Document("STGND_Documento_Graficos_Maestro_v9.docx", "DocumentoMaestro2.docx", "Informe gráfico", "{{tema}}", "Documento con visualizaciones y hoja Excel relacionada.", "{{bajada}}", "{{lectura_grafico}}", "Fuente: {{fuente_grafico}}", "Evidencia y lectura de resultados", "{{lectura_grafico}}", "ARCHIVO RELACIONADO", "STGND_Datos_Graficos_Maestro_v9.xlsx · {{fuente_grafico}}"),
]

ROLE_BLOCKS = {
    "Informe": [
        {"title": "Control documental", "rows": [["Código", "{{codigo}}"], ["Fecha", "{{fecha}}"], ["Estado", "{{estado}}"], ["Responsable", "{{responsable}}"]]},
    ],
    "Minuta": [
        {"title": "Ficha de la actividad", "rows": [["Actividad", "{{titulo}}"], ["Fecha y horario", "{{fecha}} · {{hora_inicio}}–{{hora_termino}}"], ["Lugar / modalidad", "{{lugar}} · {{modalidad}}"], ["Participantes", "{{participantes}}"], ["Responsable", "{{responsable}}"]]},
        {"title": "Acuerdos y seguimiento", "headers": ["Acuerdo", "Responsable", "Plazo", "Estado"], "dynamic": True},
    ],
    "Informativo": [
        {"title": "Control normativo", "rows": [["Tipo de instrumento", "{{tipo_norma}}"], ["Ámbito de aplicación", "{{ambito_aplicacion}}"], ["Fundamento", "{{fundamento}}"], ["Vigencia", "{{vigencia}}"], ["Responsable", "{{responsable}}"], ["Estado", "{{estado}}"]]},
    ],
    "Oficio": [
        {"title": "Identificación del oficio", "rows": [["N.º", "{{numero_documento}}"], ["ANT.", "{{antecedente}}"], ["MAT.", "{{materia}}"], ["Lugar y fecha", "{{ciudad}}, {{fecha}}"]]},
        {"title": "Destinatario", "rows": [["A", "{{destinatario}}"], ["Cargo", "{{cargo_destinatario}}"], ["Institución", "{{institucion_destinataria}}"]]},
        {"title": "Cuerpo y cierre", "rows": [["Texto", "{{cuerpo}}"], ["Firma", "{{remitente}} · {{cargo_remitente}}"], ["Distribución", "{{distribucion}}"]]},
    ],
    "Oficio circular": [
        {"title": "Identificación del oficio circular", "rows": [["N.º", "{{numero_documento}}"], ["MAT.", "{{materia}}"], ["DE", "{{remitente}} · {{cargo_remitente}}"], ["A", "{{destinatarios}}"], ["Lugar y fecha", "{{ciudad}}, {{fecha}}"]]},
        {"title": "Contenido y distribución", "rows": [["Texto", "{{cuerpo}}"], ["Distribución", "{{distribucion}}"]]},
    ],
    "Agenda": [
        {"title": "Ficha de la reunión", "rows": [["Objetivo", "{{objetivo}}"], ["Fecha", "{{fecha}}"], ["Horario", "{{hora_inicio}}–{{hora_termino}}"], ["Lugar / modalidad", "{{lugar}} · {{modalidad}}"], ["Participantes", "{{participantes}}"], ["Responsable", "{{responsable}}"]]},
        {"title": "Tabla de agenda", "headers": ["Hora", "Tema / actividad", "Responsable", "Duración"], "dynamic": True},
    ],
    "Acta": [
        {"title": "Identificación de la sesión", "rows": [["Fecha y horario", "{{fecha}} · {{hora_inicio}}–{{hora_termino}}"], ["Lugar / modalidad", "{{lugar}} · {{modalidad}}"], ["Participantes", "{{participantes}}"], ["Objetivo", "{{objetivo}}"], ["Responsable del acta", "{{responsable}}"]]},
        {"title": "Acuerdos", "headers": ["Acuerdo", "Responsable", "Plazo", "Estado"], "dynamic": True},
        {"title": "Cierre", "rows": [["Síntesis", "{{desarrollo}}"], ["Próxima reunión", "{{proxima_reunion}}"]]},
    ],
    "Programa": [
        {"title": "Ficha de la actividad", "rows": [["Actividad", "{{titulo}}"], ["Fecha", "{{fecha}}"], ["Lugar / modalidad", "{{lugar}} · {{modalidad}}"], ["Público objetivo", "{{audiencia}}"], ["Responsable", "{{responsable}}"], ["Descripción", "{{descripcion_actividad}}"]]},
        {"title": "Programa horario", "headers": ["Hora", "Actividad / contenido", "Responsable", "Duración"], "dynamic": True},
    ],
    "Checklist": [
        {"title": "Identificación de la actividad", "rows": [["Actividad", "{{titulo}}"], ["Fecha", "{{fecha}}"], ["Lugar / modalidad", "{{lugar}} · {{modalidad}}"], ["Responsable general", "{{responsable}}"]]},
        {"title": "Control de tareas", "headers": ["Estado", "Tarea", "Responsable", "Plazo", "Observaciones"], "dynamic": True},
    ],
    "Invitación": [
        {"title": "Estructura del correo", "rows": [["Para", "{{destinatarios_email}}"], ["Asunto", "{{asunto_email}}"], ["Preencabezado", "{{preheader}}"], ["Mensaje", "{{cuerpo_invitacion}}"], ["Fecha y horario", "{{fecha}} · {{hora_inicio}}"], ["Lugar / modalidad", "{{lugar}} · {{modalidad}}"], ["CTA", "{{cta}}"], ["URL", "{{url_cta}}"]]},
    ],
    "Ficha": [
        {"title": "Solicitud de compra o contratación", "rows": [["Unidad requirente", "{{unidad_requirente}}"], ["Bien o servicio", "{{item_solicitado}}"], ["Cantidad", "{{cantidad}}"], ["Motivo", "{{motivo_contratacion}}"], ["Monto estimado", "{{monto_estimado}}"], ["Ítem presupuestario", "{{partida_presupuestaria}}"], ["Convenio marco", "{{convenio_marco}}"], ["ID PAC", "{{id_pac}}"], ["Fecha requerida", "{{fecha_requerida}}"], ["Responsable", "{{responsable}}"]]},
    ],
    "Informe gráfico": [
        {"title": "Ficha de datos", "rows": [["Período", "{{periodo_datos}}"], ["Fuente", "{{fuente_grafico}}"], ["Responsable", "{{responsable}}"], ["Lectura principal", "{{lectura_grafico}}"], ["Archivo relacionado", "STGND_Datos_Graficos_Maestro_v9.xlsx"]]},
    ],
}

ROLE_LABELS = {
    "Minuta": ("Síntesis", "ACUERDO / IDEA CLAVE", "Antecedentes de la actividad"),
    "Informativo": ("Objeto", "CRITERIO NORMATIVO", "Fundamento y alcance"),
    "Oficio": ("Materia", "PROPÓSITO DEL OFICIO", "Contenido"),
    "Oficio circular": ("Materia", "PROPÓSITO DE LA CIRCULAR", "Contenido"),
    "Agenda": ("Objetivo", "RESULTADO ESPERADO", "Antecedentes de la reunión"),
    "Acta": ("Síntesis de la sesión", "ACUERDO / DECISIÓN", "Desarrollo de la sesión"),
    "Programa": ("Descripción", "PROPÓSITO DE LA ACTIVIDAD", "Antecedentes de la actividad"),
    "Checklist": ("Criterio de uso", "CONTROL OPERATIVO", "Alcance del control"),
    "Invitación": ("Mensaje principal", "LLAMADO A LA ACCIÓN", "Datos de la actividad"),
    "Ficha": ("Justificación", "NECESIDAD INSTITUCIONAL", "Antecedentes de la solicitud"),
}

ROLE_TOKENS = {
    "Minuta": ["OBJETIVO", "PARTICIPANTES", "ACUERDOS", "RESPONSABLES", "PLAZOS", "SEGUIMIENTO"],
    "Informativo": ["OBJETO", "ÁMBITO", "FUNDAMENTO", "DISPOSICIONES", "VIGENCIA", "RESPONSABLE"],
    "Oficio": ["ANTECEDENTE", "MATERIA", "DESTINATARIO", "CONTENIDO", "FIRMA", "DISTRIBUCIÓN"],
    "Oficio circular": ["MATERIA", "DESTINATARIOS", "CONTENIDO", "VIGENCIA", "FIRMA", "DISTRIBUCIÓN"],
    "Agenda": ["OBJETIVO", "HORARIO", "TEMAS", "RESPONSABLES", "INSUMOS", "SEGUIMIENTO"],
    "Acta": ["ASISTENCIA", "TEMAS", "ACUERDOS", "RESPONSABLES", "PLAZOS", "SEGUIMIENTO"],
    "Programa": ["OBJETIVO", "PÚBLICO", "HORARIO", "ACTIVIDADES", "RESPONSABLES", "RECURSOS"],
    "Checklist": ["ANTES", "DURANTE", "DESPUÉS", "RESPONSABLE", "PLAZO", "OBSERVACIÓN"],
    "Invitación": ["DESTINATARIOS", "ASUNTO", "MENSAJE", "FECHA", "CTA", "ENLACE"],
    "Ficha": ["NECESIDAD", "UNIDAD", "PRESUPUESTO", "MERCADO", "PAC", "APROBACIÓN"],
}


def parse(xml: bytes) -> etree._Element:
    return etree.fromstring(xml, etree.XMLParser(remove_blank_text=False))


def text_nodes(root: etree._Element) -> list[etree._Element]:
    return root.xpath("//w:t", namespaces=NS)


def exact(root: etree._Element, value: str) -> list[etree._Element]:
    return [node for node in text_nodes(root) if node.text == value]


def set_all(root: etree._Element, old: str, new: str) -> None:
    matches = exact(root, old)
    if not matches:
        raise ValueError(f"No se encontró texto maestro: {old}")
    for node in matches:
        node.text = new


def set_ordered(root: etree._Element, old: str, values: list[str]) -> None:
    matches = exact(root, old)
    if len(matches) != len(values):
        raise ValueError(f"Se esperaban {len(values)} ocurrencias de {old!r}; se encontraron {len(matches)}")
    for node, value in zip(matches, values):
        node.text = value


def first(root: etree._Element, old: str, new: str) -> None:
    matches = exact(root, old)
    if not matches:
        raise ValueError(f"No se encontró texto maestro: {old}")
    matches[0].text = new


def set_editable_size(root: etree._Element) -> None:
    tokens = ("{{resumen", "{{idea", "{{cuerpo", "{{texto", "{{dynamic", "{{objetivo", "{{desarrollo", "{{descripcion", "{{motivo", "{{lectura", "{{fuente")
    for node in text_nodes(root):
        if not node.text or not any(token in node.text for token in tokens):
            continue
        run = node.getparent()
        if run.tag != f"{{{W}}}r":
            continue
        properties = run.find(f"{{{W}}}rPr")
        if properties is None:
            properties = etree.Element(f"{{{W}}}rPr")
            run.insert(0, properties)
        for tag in ("sz", "szCs"):
            size = properties.find(f"{{{W}}}{tag}")
            if size is None:
                size = etree.SubElement(properties, f"{{{W}}}{tag}")
            size.set(f"{{{W}}}val", "22")


def w_element(tag: str, **attributes: str) -> etree._Element:
    node = etree.Element(f"{{{W}}}{tag}")
    for key, value in attributes.items():
        node.set(f"{{{W}}}{key}", value)
    return node


def role_paragraph(text: str, *, title: bool = False, header: bool = False, label: bool = False) -> etree._Element:
    paragraph = w_element("p")
    properties = w_element("pPr")
    spacing = w_element("spacing", before="180" if title else "0", after="100" if title else "40", line="276", lineRule="auto")
    properties.append(spacing)
    paragraph.append(properties)
    run = w_element("r")
    run_properties = w_element("rPr")
    fonts = w_element("rFonts", ascii="Aptos Display" if title else "Aptos", hAnsi="Aptos Display" if title else "Aptos", eastAsia="Aptos", cs="Aptos")
    run_properties.append(fonts)
    run_properties.append(w_element("sz", val="30" if title else "22"))
    run_properties.append(w_element("szCs", val="30" if title else "22"))
    if title or header or label:
        run_properties.append(w_element("b"))
    run_properties.append(w_element("color", val="FFFFFF" if header else "4C2B46" if title or label else "222222"))
    run.append(run_properties)
    text_node = w_element("t")
    text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    run.append(text_node)
    paragraph.append(run)
    return paragraph


def role_cell(text: str, *, header: bool = False, label: bool = False) -> etree._Element:
    cell = w_element("tc")
    properties = w_element("tcPr")
    properties.append(w_element("tcW", w="2400", type="dxa"))
    margins = w_element("tcMar")
    for side in ("top", "left", "bottom", "right"):
        margins.append(w_element(side, w="110", type="dxa"))
    properties.append(margins)
    if header:
        properties.append(w_element("shd", fill="4C2B46"))
    elif label:
        properties.append(w_element("shd", fill="EFEDEF"))
    cell.append(properties)
    cell.append(role_paragraph(text, header=header, label=label))
    return cell


def role_table(block: dict[str, object]) -> etree._Element:
    headers = block.get("headers") or ["Campo", "Contenido"]
    rows = block.get("rows") or []
    table = w_element("tbl")
    properties = w_element("tblPr")
    properties.append(w_element("tblW", w="5000", type="pct"))
    properties.append(w_element("jc", val="center"))
    properties.append(w_element("tblLayout", type="fixed"))
    borders = w_element("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(w_element(side, val="single", sz="4", color="D9D5D8"))
    properties.append(borders)
    table.append(properties)
    grid = w_element("tblGrid")
    width = str(9600 // len(headers))
    for _ in headers:
        grid.append(w_element("gridCol", w=width))
    table.append(grid)

    header_row = w_element("tr")
    header_row.append(w_element("trPr"))
    for value in headers:
        header_row.append(role_cell(str(value), header=True))
    table.append(header_row)

    for values in rows:
        row = w_element("tr")
        row_properties = w_element("trPr")
        row_properties.append(w_element("cantSplit"))
        row.append(row_properties)
        for index, value in enumerate(values):
            row.append(role_cell(str(value), label=index == 0 and len(headers) == 2))
        table.append(row)

    if block.get("dynamic"):
        row = w_element("tr")
        row.append(w_element("trPr"))
        for index in range(len(headers)):
            row.append(role_cell("{{dynamic_rows}}" if index == 0 else ""))
        table.append(row)
    return table


def insert_role_blocks(root: etree._Element, spec: Document, anchor_text: str) -> None:
    blocks = ROLE_BLOCKS.get(spec.kind, [])
    if not blocks:
        return
    matches = exact(root, anchor_text)
    if not matches:
        raise ValueError(f"No se encontró ancla funcional para {spec.kind}: {anchor_text}")
    paragraph = matches[-1]
    while paragraph is not None and paragraph.tag != f"{{{W}}}p":
        paragraph = paragraph.getparent()
    if paragraph is None:
        raise ValueError(f"El ancla funcional de {spec.kind} no está en un párrafo")
    parent = paragraph.getparent()
    position = parent.index(paragraph) + 1
    for block in blocks:
        title = role_paragraph(str(block["title"]), title=True)
        table = role_table(block)
        parent.insert(position, title)
        parent.insert(position + 1, table)
        position += 2


def customize_role_vocabulary(root: etree._Element, spec: Document) -> None:
    labels = ROLE_LABELS.get(spec.kind)
    if labels:
        summary_label, callout_label, context_label = labels
        set_all(root, "Resumen", summary_label)
        set_all(root, "HIPÓTESIS / IDEA CLAVE", callout_label)
        set_all(root, "Antecedentes", context_label)
    tokens = ROLE_TOKENS.get(spec.kind)
    if tokens:
        originals = ["IDEAS", "PREGUNTA", "EVIDENCIA", "ANÁLISIS", "IMPACTO", "DATOS "]
        for old, new in zip(originals, tokens):
            set_all(root, old, new)


def trim_irrelevant_findings(root: etree._Element, spec: Document) -> None:
    if spec.kind in {"Informe", "Informe gráfico"}:
        return
    body = root.find(f".//{{{W}}}body")
    children = list(body)
    start = None
    for index, child in enumerate(children):
        content = "".join(child.xpath(".//w:t/text()", namespaces=NS))
        if "HALLAZGOS" in content and "03" in content:
            start = index
            break
    if start is None:
        raise ValueError(f"No se encontró página de hallazgos para retirar en {spec.kind}")
    for child in children[start:]:
        if child.tag != f"{{{W}}}sectPr":
            body.remove(child)


def customize_document(source: bytes, spec: Document) -> bytes:
    root = parse(source)
    set_all(root, "TEMA", spec.theme)
    set_all(root, "Informe", spec.kind)
    set_all(root, "Anual de actividades", "{{titulo}}")
    set_all(root, "Subtítulo descriptivo o bajada editorial que contextualiza el alcance del documento.", spec.subtitle)
    set_all(root, "2025- ", "{{anio}}" if spec.sections or spec.kind == "Informe gráfico" else "{{fecha}}")
    set_all(root, "2026", "")
    if exact(root, "Codigo de documento"):
        set_all(root, "Codigo de documento", "{{codigo}}")
    else:
        set_all(root, "Codigo", "{{codigo}}")
        set_all(root, " de documento", "")
    set_all(root, "Título del documento", "{{titulo}}")
    set_ordered(root, "Subtítulo o descripción breve del objetivo y alcance del documento.", [spec.subtitle, spec.section_summary])
    summary_text = "Utilice este espacio para una síntesis de 120–180 palabras. La plantilla privilegia márgenes amplios, jerarquías precisas y una lectura editorial limpia. El naranja funciona como señal y no como superficie dominante."
    set_ordered(root, summary_text, [spec.summary, spec.section_summary])
    set_all(root, "Utilice este espacio para una síntesis de 120–180 palabras. La plantilla privilegia márgenes amplios, jerarquías precisas y una lectura editorial limpia. El naranja funciona como señal y no como superficie dominante", spec.section_summary)
    set_ordered(root, "“Una formulación breve que concentre la tesis, hallazgo o criterio de lectura.”", [f"“{spec.callout}”", f"“{spec.section_callout}”"])
    first(root, "Desarrolle el contexto en párrafos breves. Para documentos extensos, use los estilos Título 1, Título 2 y Título 3 incluidos en el archivo. Evite densidades mayores a 65–75 caracteres por línea.", spec.context)
    set_all(root, "Segunda sección o primera del documento con portadilla", spec.section_title)
    set_all(root, "TEXTO META O PALABRA CLAVE SENCILLA DEL TEMA O CUALQUIER OTRA DEFINICIÓN", "CONTENIDO · RESPONSABLES · PLAZOS · EVIDENCIA")
    if exact(root, "Titulo sección"):
        set_all(root, "Titulo sección", spec.section_title)
    else:
        set_all(root, "Titulo", spec.section_title)
        set_all(root, " sección", "")
    section_flow = spec.section_body if spec.sections else "Utilice la estructura funcional incluida a continuación y complete cada campo aplicable."
    set_all(root, "Desarrolle el contexto en párrafos breves. Para documentos extensos", section_flow)
    set_all(root, ", use los estilos Título 1, Título 2 y Título 3 incluidos en el archivo. Evite densidades mayores a 65–75 caracteres por línea.", "")
    set_all(root, "Planteamiento", "Desarrollo")
    detail_body = "Continúe aquí el desarrollo de la sección con los estilos del documento maestro." if spec.sections else "Complete los bloques funcionales del formato y elimine las filas que no correspondan."
    set_all(root, "Texto de trabajo editable. La estructura está dimensionada para mantener líneas de lectura estables y evitar problemas de desborde. Puede duplicar bloques, insertar citas o reemplazar este contenido por texto definitivo.", detail_body)
    if not spec.sections:
        # Los documentos operativos mantienen la portadilla y página interior del maestro,
        # pero usan contenido específico en lugar de marcadores de sección.
        for node in text_nodes(root):
            if node.text:
                node.text = node.text.replace("{{resumen_seccion}}", spec.section_summary).replace("{{idea_clave_seccion}}", spec.section_callout).replace("{{texto_seccion}}", spec.section_body)
    insert_role_blocks(root, spec, detail_body)
    customize_role_vocabulary(root, spec)
    trim_irrelevant_findings(root, spec)
    set_editable_size(root)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def normalize_styles(source: bytes) -> bytes:
    root = parse(source)
    normal = root.xpath("//w:style[@w:styleId='Normal']", namespaces=NS)[0]
    properties = normal.find(f"{{{W}}}rPr")
    if properties is None:
        properties = etree.SubElement(normal, f"{{{W}}}rPr")
    fonts = properties.find(f"{{{W}}}rFonts")
    if fonts is None:
        fonts = etree.SubElement(properties, f"{{{W}}}rFonts")
    for attribute in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(f"{{{W}}}{attribute}", "Aptos")
    for tag in ("sz", "szCs"):
        size = properties.find(f"{{{W}}}{tag}")
        if size is None:
            size = etree.SubElement(properties, f"{{{W}}}{tag}")
        size.set(f"{{{W}}}val", "22")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def core_properties(source: bytes, title: str) -> bytes:
    root = parse(source)
    nodes = root.xpath("//dc:title", namespaces=NS)
    if nodes:
        nodes[0].text = title
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def metrics(data: bytes) -> tuple[int, int, int, int, int]:
    root = parse(data)
    body = root.find(f".//{{{W}}}body")
    serialized = data.decode("utf-8", "ignore")
    return len(body), len(root.xpath("//w:tbl", namespaces=NS)), len(root.xpath("//w:sectPr", namespaces=NS)), serialized.count("wp:anchor"), len(root.xpath("//w:drawing", namespaces=NS))


def build(spec: Document) -> None:
    master_path = BASE / spec.master
    with zipfile.ZipFile(master_path) as package:
        entries = [(info, package.read(info.filename)) for info in package.infolist()]
        master_document = package.read("word/document.xml")
        expected_names = [info.filename for info, _ in entries]
    replacements = {
        "word/document.xml": customize_document(master_document, spec),
        "word/styles.xml": None,
        "docProps/core.xml": None,
    }
    entry_map = dict((info.filename, data) for info, data in entries)
    replacements["word/styles.xml"] = normalize_styles(entry_map["word/styles.xml"])
    replacements["docProps/core.xml"] = core_properties(entry_map["docProps/core.xml"], spec.kind)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as target:
        for info, data in entries:
            target.writestr(info, replacements.get(info.filename) or data)
    with zipfile.ZipFile(io.BytesIO(output.getvalue())) as check:
        if check.testzip() is not None:
            raise ValueError(f"Paquete inválido: {spec.filename}")
        if check.namelist() != expected_names:
            raise ValueError(f"Cambió la lista de partes: {spec.filename}")
        built_document = check.read("word/document.xml")
    master_shape = metrics(master_document)
    built_shape = metrics(built_document)
    minimum_tables = len(ROLE_BLOCKS.get(spec.kind, [])) + 1
    if built_shape[2:] != master_shape[2:] or built_shape[1] < minimum_tables:
        raise ValueError(f"La ampliación funcional dañó la topología maestra: {spec.filename}; maestro={master_shape}; salida={built_shape}")
    target_path = OUT / spec.filename
    temporary = target_path.with_suffix(".docx.tmp")
    temporary.write_bytes(output.getvalue())
    temporary.replace(target_path)
    print(f"OK {spec.filename} <- {spec.master} topology={metrics(built_document)}")


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    for spec in DOCS:
        build(spec)
    from ajustar_formatos_funcionales import main as ajustar

    ajustar()
    print(f"Construidos desde maestros y referencias institucionales: {len(DOCS)} DOCX")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
