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


## Formatos Alfred

La categoría **Alfred** utiliza directamente los siete documentos contenidos en `base/STGND - Oficio.zip`. El catálogo extrae el DOCX correspondiente en el navegador y sustituye únicamente los campos configurables de contenido.

Regla de implementación: no se remaquetan los documentos de origen, no se crean tablas nuevas, no se agregan cuadros de texto ni figuras y se preservan márgenes, firma digital, logotipos, tablas, casillas y recursos gráficos existentes. Esto mantiene la compatibilidad con el gestor institucional Alfred y la integridad de formatos como oficio, oficio circular, agenda, acta, informe, minuta y ficha de compras.


## v9.5 — revisión institucional

- Se elimina de la interfaz el documento normativo y la ficha de solicitud de compras.
- La categoría Correspondencia pasa a llamarse **Oficios**.
- Agenda y Acta de la categoría Reuniones usan directamente los formatos fuente compatibles con Alfred, que ya incorporan el logotipo institucional, evitando una segunda remaquetación.
- Las tarjetas Alfred incluyen dos acciones visibles: **Configurar** y **Descargar formato**.
- La categoría Alfred incluye un botón para descargar el paquete fuente completo `base/STGND - Oficio.zip`.
- El paquete fuente permanece íntegro aunque algunos documentos ya no se muestren como tarjetas del catálogo.
- La actualización visual de Excel y PPT se mantiene como una revisión separada de los binarios Office.


## v9.6 — Alfred visual mínimo

La colección Alfred recibe una capa gráfica compatible sin reconstruir su arquitectura documental:

- Aptos como tipografía de trabajo y cuerpo de 11 pt donde corresponde.
- Púrpura `#4C2B46` para jerarquías y naranja `#E97700` como señal, siguiendo el lenguaje visual del kit.
- Reglas de párrafo y sombreado de celdas existentes en lugar de cuadros de texto o figuras flotantes nuevas.
- Agenda y Acta mantienen su logotipo institucional; Acta estiliza únicamente la tabla de asistencia ya existente.
- Oficio y Oficio circular conservan los tabuladores y márgenes laterales institucionales de 30 mm; se reduce sólo el espacio vertical para evitar el salto innecesario de distribución a una segunda página.
- Los bloques `FIRMADIGITAL` mantienen exactamente su objeto y geometría; sólo se suaviza el azul heredado a un gris-púrpura neutro.
- Informe y Minuta conservan sus gráficos y secuencia; se normalizan tipografía, interlineado y jerarquías.
- No se crean tablas nuevas, no se agregan textos flotantes y no se modifica el mapeo usado por el configurador.
- El formato de compras continúa fuera del catálogo visible y se conserva sin modificaciones dentro del paquete por compatibilidad histórica.
- El paquete fuente anterior se preserva en `base/originales/STGND - Oficio_original.zip`.

La construcción es reproducible mediante `tools/build_alfred_visual_v96.py`.
