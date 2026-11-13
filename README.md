# apa-unab

Skill para Claude que genera y valida documentos Word (`.docx`) en formato **APA 7.ª edición** según la Universidad Andrés Bello (Tech MBA): portada, títulos, citas, referencias, tablas, figuras y apéndices.

## Contenido

- `apa-unab/SKILL.md` — instrucciones de la skill (reglas de formato, citas, referencias, formato de entrada y flujo de trabajo).

Los scripts `build_docx.py` y `validate_docx.py` (Python + `python-docx`) y el logo de portada **no se incluyen** en este repositorio: la skill los espera desde un Project de Claude. Sin ellos, la skill aplica las reglas a mano y no afirma haber validado el formato.

## Instalación

Copia la carpeta `apa-unab/` en el directorio de skills de tu cliente (por ejemplo `~/.claude/skills/apa-unab/` en Claude Code) o súbela como skill en claude.ai.

## Aviso de responsabilidad

Este proyecto se entrega **«tal cual es» («as is»)**, sin garantías de ningún tipo, expresas ni implícitas, incluidas las de exactitud, idoneidad para un fin determinado o cumplimiento de los requisitos de un curso, profesor o institución.

- **El uso es de exclusiva responsabilidad de quien lo utiliza.** Revisa siempre el documento resultante antes de entregarlo; ni el autor ni los colaboradores responden por calificaciones, rechazos, errores de formato o de citación, ni por ningún daño derivado de su uso.
- Las instrucciones del profesor, la rúbrica del curso y las guías oficiales de la UNAB y de la APA prevalecen sobre esta skill.
- Usar esta herramienta no exime de las normas de integridad académica de la institución; el contenido del trabajo es responsabilidad de su autor.
- Proyecto independiente y no oficial: **no está afiliado, patrocinado ni avalado** por la Universidad Andrés Bello ni por la American Psychological Association. Los nombres y marcas mencionados pertenecen a sus titulares.

## Licencia

[MIT](LICENSE).
