# Catálogo STGND Documentos v9.3

Catálogo local de instrumentos corporativos construidos desde los maestros gráficos y las referencias institucionales incluidas en `base/`.

## Versión online

El catálogo se publica como sitio estático en `https://ngrr.github.io/stgkit/`. La generación ocurre en el navegador: las plantillas se descargan desde el mismo sitio, se procesan localmente y el DOCX resultante no se envía a un servidor.

## Inicio

Ejecute `iniciar_catalogo.cmd`. El servicio queda limitado a `127.0.0.1` y se detiene al presionar Enter en su consola. El catálogo no usa automatización de Word.

## Criterio documental

La identidad maestra se conserva, pero cada formato contiene sólo los componentes necesarios para cumplir su función:

- Informe: mantiene la estructura extensa, secciones agregables, tres gráficos Office y permite escoger portada blanca o portada con fondo de color.
- Oficio y oficio circular: derivan de los documentos oficiales de `base/STGND - Oficio.zip`; usan número, MAT., REF., fecha, destinatario, remitente, cuerpo, cierre y distribución.
- Acta: una sección, sin portada ni portadillas; contiene identificación, participantes, desarrollo, acuerdos y próxima reunión.
- Agenda: una sección, sin portada ni portadillas; contiene datos de sesión, objetivo, agenda dinámica y documentos preparatorios.
- Minuta: una sección, sin portada, secciones ni gráficos; contiene datos de actividad, objetivo, contenido y seguimiento.
- Programa: deriva de `base/ProgramaHito_lanzamiento.docx`; una página con imagen de cabecera reemplazable, identificación de actividad y tabla horaria dinámica.
- Invitación: deriva de `base/Invitacion_Hito_Guia.docx`, conserva su lógica de pieza institucional para correo y permite reemplazar la imagen de cabecera.
- Documento gráfico: contiene diez gráficos Office nativos, tabla comparativa, esquema de proceso, libros Excel incrustados y vínculo al libro de datos.
- Libro de datos: incluye dashboard y galería de diez gráficos Excel nativos vinculados a sus datos.
- Presentación: ofrece portada morada y blanca, elimina objetos rotados, incorpora gráficos PowerPoint nativos y usa logotipo blanco de contraste en el cierre.

Los documentos de una sección no incluyen portadas, portadillas, secciones editoriales ni ejemplos gráficos improcedentes. Los DOCX editoriales incluyen `Title`, `Heading 1` a `Heading 3`, `Normal` y `Caption`; los encabezados tienen nivel de esquema para navegación y tablas de contenido. Oficio y oficio circular son excepciones formales: conservan los estilos de su referencia oficial y todo su texto se mantiene negro. Las firmas electrónicas usan un bloque en línea de ancho controlado, sin objetos flotantes.

## Arquitectura

- `base/`: maestros y referencias institucionales autoritativas.
- `formatos/docx_maestro/`: 13 plantillas DOCX; dos corresponden a las variantes de portada del informe.
- `formatos/xlsx/` y `formatos/pptx/`: complementos activos.
- `scripts/construir_desde_maestros.py`: reconstrucción completa.
- `scripts/ajustar_formatos_funcionales.py`: reducción funcional, referencias y gráficos Office.
- `scripts/verificar_catalogo.py`: auditoría estructural.
- `scripts/test_generador.js`: prueba de generación en memoria.

## Reconstrucción y verificación

```powershell
python scripts\construir_desde_maestros.py
python scripts\verificar_catalogo.py
& 'C:\Program Files\Adobe\Adobe Creative Cloud Experience\libs\node.exe' scripts\test_generador.js
```
