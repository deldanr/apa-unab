---
name: "apa-unab"
description: "Genera y valida documentos Word (.docx) en APA 7ª edición según la UNAB (Tech MBA): portada, títulos, citas, referencias, tablas y apéndices. Úsala al crear o revisar trabajos para la UNAB."
---

# Documentos APA 7 para la UNAB (Tech MBA)

Produce un `.docx` con estilos reales de Word y lo valida. Usa dos scripts de Python (`build_docx.py` y `validate_docx.py`, requieren `python-docx`; para revisar el resultado, LibreOffice y `pdftoppm`).

**Dónde están los scripts:** en el documento del Project «Tech MBA UNAB» llamado `claude/apa-unab-skill-completa.md` (léelo con `project_read`; contiene el ejemplo mínimo y los dos scripts en bloques de código `python`, que se guardan tal cual como `build_docx.py` y `validate_docx.py`). Si ese documento no está disponible en la sesión, dilo al usuario, pídele que lo adjunte y, mientras tanto, aplica las reglas de este archivo a mano y no afirmes que validaste el formato.

## Prioridad de reglas
1. Instrucción explícita del profesor o rúbrica del curso.
2. Guía de la Biblioteca UNAB «Estilo APA: citas y referencias» (Unidad DCI), en lo que define explícitamente.
3. APA 7ª edición (guía «Normas APA 7ma edición», no oficial; la oficial está en apastyle.apa.org).
4. Los ejemplos de la guía UNAB con rasgos de la 6ª edición no se copian: citan «(APA, 2010)», incluyen ciudad de la editorial, «Recuperado el» en páginas estables y «Autor.» como editorial.

Decisiones tomadas con el usuario: APA 7; la lista se titula «Referencias» (no «Referencias Bibliográficas»); salida en .docx.

## Logo de portada
El Tech MBA pertenece a la Facultad de Economía y Negocios (FEN): por defecto la portada lleva el logo «Universidad Andrés Bello | Facultad de Economía y Negocios» (PNG 812x385, fondo blanco). Está guardado en el Project como `claude/apa-unab-logo-fen-base64.md` (bloque base64): léelo con `project_read`, decodifica el bloque con `base64 -d` a `logo-fen.png` y pon `"logo": "logo-fen.png"` en `datos.json`. El ancho por defecto es 3 pulgadas (`"logo_ancho"` lo cambia). Con logo no se necesita `marca`. Solo si el usuario pide otro logo (UNAB vertical u horizontal, sin facultad) debe adjuntarlo. Si el logo no está disponible, generar la portada sin logo y avisarlo.

## Flujo
1. Reunir título, autor(es), curso, profesor, fecha, texto y fuentes con datos completos (autor, año, título, editorial, revista, URL o DOI).
2. Escribir `cuerpo.md` y `datos.json` (formato abajo). No inventar datos de una fuente: si falta autor, año o editorial, preguntar o usar «s.f.» / entrada por título.
3. Guardar los scripts y ejecutar `python3 build_docx.py cuerpo.md datos.json salida.docx`.
4. Validar con `python3 validate_docx.py salida.docx`; corregir todo ERROR y revisar cada AVISO (algunos son heurísticos).
5. Revisar visualmente: `soffice --headless --convert-to pdf salida.docx` y `pdftoppm -png -r 70 salida.pdf p`; mirar portada, una página de texto, la tabla y las referencias.
6. Entregar el `.docx` e indicar lo que quedó pendiente.

## Formato (guía APA 7)
- Papel carta, márgenes de 2,54 cm.
- Una sola fuente: Times New Roman 12 (por defecto), Calibri 11, Arial 11 o Georgia 11. La UNAB no define tipografía para trabajos.
- Interlineado doble en todo (resumen, títulos, citas en bloque, referencias, número y título de tablas y figuras, notas); sin espacio extra entre párrafos ni líneas en blanco de relleno.
- Alineación a la izquierda (no justificar, salvo que el profesor lo pida). Sangría de 1,27 cm en la primera línea de cada párrafo, con estilo de párrafo, nunca con espacios.
- Número de página automático, arriba a la derecha; la portada es la página 1. Trabajos de estudiante sin título abreviado.
- Portada: título en negrita, línea en blanco, autor, afiliación (Universidad Andrés Bello), programa, curso, profesor y fecha, todo centrado.
- Orden: portada, resumen (solo si se pide; 150–250 palabras, un párrafo sin sangría, «Palabras clave:» en cursiva), texto con el título repetido arriba (centrado, negrita), Referencias, apéndices.
- Títulos con capitalización de título, sin numerar, niveles en orden: 1 centrado en negrita; 2 izquierda en negrita; 3 izquierda en negrita cursiva; 4 con sangría, negrita, termina en punto y el texto sigue en la misma línea; 5 igual al 4 pero en negrita cursiva.
- Tablas: «Tabla N» en negrita, título en cursiva debajo, solo líneas horizontales, encabezados centrados, «Nota.» en cursiva si hace falta; citarlas por número en el texto.
- Figuras: «Figura N» en negrita, título en cursiva, «Nota.» con fuente o licencia si no son propias.
- Apéndices: página nueva cada uno, después de Referencias; «Apéndice» si hay uno, «Apéndice A», «B»… si hay varios; tablas y figuras dentro llevan la letra (Tabla A1).

## Citas (guía UNAB + APA 7)
- Autor-fecha; toda cita está en Referencias y viceversa, con ortografía y años idénticos.
- Narrativa: Luna (2020). Parentética: (Luna, 2020). 1 autor: Luna. 2 autores: Salas y D'Agostino. 3 o más: Martín et al. desde la primera cita.
- Autor de grupo con sigla: primera cita (Organización Panamericana de la Salud [OPS], 2003) u Organización Panamericana de la Salud (OPS, 2003); luego (OPS, 2003) u OPS (2003).
- Varias fuentes en un paréntesis: orden alfabético, con punto y coma. Mismo autor y año: 2020a, 2020b.
- Cita textual de menos de 40 palabras: entre comillas, en el párrafo, con autor, año y página. De 40 o más: bloque aparte, sangría de 1,27 cm, sin comillas; el punto va antes de (Autor, año, p. x).
- Parafraseo: sin comillas, con autor y año; la página es recomendable.
- [sic] para errores de la fuente, […] para omisiones, [texto] para añadidos propios, [énfasis agregado] tras cursiva propia.
- Fuente secundaria: (Ayala, 1983, como se citó en Sánchez, 2009); en Referencias solo la consultada. Comunicaciones personales: solo en el texto.
- Imágenes o tablas de otros: atribución de derechos de autor en la nota y entrada en Referencias.

## Referencias
- Página nueva, «Referencias» en negrita y centrado; orden alfabético letra por letra por el apellido del primer autor; sangría francesa de 1,27 cm; doble espacio.
- Autor, fecha, título y fuente. Hasta 20 autores («y» antes del último en español, «&» en inglés); con más de 20, los 19 primeros, puntos suspensivos y el último.
- Sin ciudad de la editorial; si la editorial es el mismo autor, se omite. DOI como https://doi.org/…; sin punto final tras URL o DOI. Sin fecha: s.f. Sin autor: el título pasa a su lugar.
- Revistas: *Revista*, *volumen*(número), páginas (sin las palabras «volumen» ni «número»). Libros: título en cursiva, edición entre paréntesis. Capítulos: «En I. Apellido (Ed.), *Libro* (pp. x–y).»
- `build_docx.py` soporta: libro, capitulo, revista, periodico, web, informe, tesis, video, podcast, webinar, presentacion, social, ley. Para otro tipo, escribir la referencia a mano con estos principios y avisarlo.

## Formato de entrada
`datos.json`: `portada` (titulo, autores[], afiliacion, programa, curso, profesor, fecha); opcionales `resumen`, `palabras_clave[]`, `fuente`, `logo` (ruta a PNG/JPG), `logo_ancho` (pulgadas, por defecto 3), `marca` (true agrega una línea en rojo institucional #A90429) y `fuentes[]`. Cada fuente lleva `id`, `tipo`, `fecha` («2003», «2019-11-16» o «s.f.»), `titulo` y, según el tipo: `autores` («Apellido, I.»), `corporativo` y `sigla`, `editorial`, `edicion` («2.ª»), `revista`, `volumen`, `numero`, `paginas`, `doi`, `url`, `editores` y `libro` (capítulos), `sitio`, `usuario`, `plataforma`, `tipo_tesis`, `institucion`, `repositorio`, `periodico`, `numero_informe`, `rol` («Ed.», «Coords.»), `idioma` («en» usa «&»), `recuperado`.

`cuerpo.md`:
- `#` a `#####` son los niveles de título 1 a 5 (el 4 y el 5 se unen al párrafo siguiente). Párrafos separados por línea en blanco; `*cursiva*`, `**negrita**`, listas con `- `.
- `> texto` es cita en bloque, con la cita al final: `> texto. [@id|pp. 3-4]`.
- Citas: `[@id]` o `[@id|p. 5]` parentética; `[@a; @b]` varias; `[n@id]` o `[n@id|p. 5]` narrativa (el autor sale del marcador: no escribirlo antes); `[loc|p. 541]` solo página; `[cp|N. Apellido, comunicación personal, 3 de mayo de 2026]`.
- `::tabla Título` seguido de filas `| a | b |` (la primera es el encabezado) y `::nota texto`; `::figura ruta | Título`; `::apendice Título`; `::pagina`.
- Solo se incluyen en Referencias las fuentes citadas.

## Límites
- La UNAB no entrega en estos documentos portada, tipografía ni márgenes propios: si el usuario consigue una plantilla o rúbrica del programa, esa manda. El logo de la portada es el de la Facultad de Economía y Negocios (ver «Logo de portada»); cualquier otro logo debe adjuntarlo el usuario (institucional.unab.cl, sección logos). Colores institucionales: azul #03102C (Pantone 296) y rojo #A90429 (Pantone 187); APA no usa color, por eso `marca` es opcional y viene desactivada.
- Los avisos del validador sobre citas y referencias son heurísticos (autor corporativo, apellidos compuestos): revisar a mano.
- LibreOffice sustituye Times New Roman por una fuente equivalente al exportar a PDF; en Word se ve con la fuente real.