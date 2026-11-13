# apa-unab

Skill para Claude que genera y valida documentos Word (`.docx`) en **APA 7.ª edición** según la Universidad Andrés Bello (Tech MBA): portada con logo, títulos, citas, referencias, tablas, figuras y apéndices. Todo viene incluido en este repositorio: reglas, scripts, logos y un ejemplo.

## Instalación (copiar y pegar)

### Opción 1: pídeselo a tu agente

Pega esto en Claude Code (u otro agente con acceso a terminal):

```text
Instrucciones para tu agente: "Descarga e instala el skill APA-UNAB desde https://github.com/deldanr/apa-unab. Clona el repositorio, copia la carpeta apa-unab/ a ~/.claude/skills/apa-unab/, instala las dependencias con `python3 -m pip install -r ~/.claude/skills/apa-unab/requirements.txt` y `python3 -m playwright install chromium`, y verifica la instalación generando y validando el ejemplo de ~/.claude/skills/apa-unab/examples/."
```

### Opción 2: comandos (macOS / Linux)

```bash
git clone https://github.com/deldanr/apa-unab.git /tmp/apa-unab-repo \
  && mkdir -p ~/.claude/skills \
  && cp -R /tmp/apa-unab-repo/apa-unab ~/.claude/skills/apa-unab \
  && python3 -m pip install -r ~/.claude/skills/apa-unab/requirements.txt \
  && python3 -m playwright install chromium
```

Para instalarla solo en un proyecto, usa `<tu-proyecto>/.claude/skills/apa-unab` en lugar de `~/.claude/skills/apa-unab`.

### Opción 3: Windows (PowerShell)

```powershell
git clone https://github.com/deldanr/apa-unab.git $env:TEMP\apa-unab-repo
New-Item -ItemType Directory -Force "$HOME\.claude\skills" | Out-Null
Copy-Item -Recurse -Force "$env:TEMP\apa-unab-repo\apa-unab" "$HOME\.claude\skills\apa-unab"
python -m pip install -r "$HOME\.claude\skills\apa-unab\requirements.txt"
python -m playwright install chromium
```

### Opción 4: claude.ai (subir como skill)

```bash
git clone https://github.com/deldanr/apa-unab.git /tmp/apa-unab-repo \
  && cd /tmp/apa-unab-repo && zip -r apa-unab.zip apa-unab
```

Sube `apa-unab.zip` como skill desde la configuración de Skills de claude.ai. Esta opción necesita un entorno que pueda ejecutar Python con `python-docx`; si no lo tiene, la skill aplica las reglas a mano y no afirma haber validado el formato.

### Verificar la instalación

```bash
cd /tmp && python3 ~/.claude/skills/apa-unab/scripts/build_docx.py \
  ~/.claude/skills/apa-unab/examples/cuerpo.md ~/.claude/skills/apa-unab/examples/datos.json ejemplo.docx \
  && python3 ~/.claude/skills/apa-unab/scripts/validate_docx.py ejemplo.docx
```

Debe terminar con `0 errores, 0 avisos`.

### Actualizar o desinstalar

```bash
# actualizar
git clone https://github.com/deldanr/apa-unab.git /tmp/apa-unab-repo && rm -rf ~/.claude/skills/apa-unab && cp -R /tmp/apa-unab-repo/apa-unab ~/.claude/skills/apa-unab
# desinstalar
rm -rf ~/.claude/skills/apa-unab
```

## Cómo usarla

Dile a tu agente, por ejemplo: «Con la skill apa-unab, genera un .docx de mi trabajo (adjunto el texto y las fuentes)» o «Valida este .docx con la skill apa-unab». La skill pide los datos de portada y de las fuentes, genera el Word, lo valida y lo revisa visualmente.

## Requisitos

| Qué | Para qué | Instalación |
|---|---|---|
| Python 3.9+ y `python-docx` | Generar y validar el `.docx` | `python3 -m pip install -r requirements.txt` |
| Playwright + Chromium | Solo el fallback de logos | `python3 -m playwright install chromium` |
| LibreOffice y `pdftoppm` | Opcional: revisar el resultado como PDF/imagen | `brew install --cask libreoffice` y `brew install poppler` (macOS) |

## Logos y fallback

Los logos vienen en `apa-unab/assets/logos/`:

- `logo-fen.png`: Universidad Andrés Bello | Facultad de Economía y Negocios (el de la portada por defecto).
- `logo-unab-horizontal.png` y `logo-unab-vertical.png`: logo institucional.

Si falta alguno, `build_docx.py` lo descarga automáticamente desde https://institucional.unab.cl usando Playwright. También se puede hacer a mano:

```bash
python3 ~/.claude/skills/apa-unab/scripts/fetch_logos.py            # descarga los que falten
python3 ~/.claude/skills/apa-unab/scripts/fetch_logos.py --forzar   # vuelve a descargar todos
python3 ~/.claude/skills/apa-unab/scripts/fetch_logos.py --listar   # muestra los enlaces de imagen de la página (depuración)
```

Si tampoco se pueden descargar, la portada se genera sin logo y se avisa. En `datos.json`, `"logo"` acepta `"fen"` (por defecto), `"unab-horizontal"`, `"unab-vertical"`, la ruta a otra imagen o `false`.

## Contenido del repositorio

```text
apa-unab/
├── SKILL.md              instrucciones de la skill (reglas, formato de entrada, flujo)
├── requirements.txt      dependencias de Python
├── scripts/
│   ├── build_docx.py     genera el .docx desde cuerpo.md + datos.json
│   ├── validate_docx.py  valida un .docx contra APA 7 / UNAB
│   └── fetch_logos.py    descarga los logos de institucional.unab.cl (Playwright)
├── assets/logos/         logo-fen.png, logo-unab-horizontal.png, logo-unab-vertical.png
└── examples/             datos.json y cuerpo.md de ejemplo
```

## Aviso de responsabilidad

Este proyecto se entrega **«tal cual es» («as is»)**, sin garantías de ningún tipo, expresas ni implícitas, incluidas las de exactitud, idoneidad para un fin determinado o cumplimiento de los requisitos de un curso, profesor o institución.

- **El uso es de exclusiva responsabilidad de quien lo utiliza.** Revisa siempre el documento resultante antes de entregarlo; ni el autor ni los colaboradores responden por calificaciones, rechazos, errores de formato o de citación, ni por ningún daño derivado de su uso.
- Las instrucciones del profesor, la rúbrica del curso y las guías oficiales de la UNAB y de la APA prevalecen sobre esta skill.
- Usar esta herramienta no exime de las normas de integridad académica de la institución; el contenido del trabajo es responsabilidad de su autor.
- Proyecto independiente y no oficial: **no está afiliado, patrocinado ni avalado** por la Universidad Andrés Bello ni por la American Psychological Association. Los nombres y marcas mencionados pertenecen a sus titulares. Los logos de `apa-unab/assets/logos/` son de la UNAB, se descargaron de https://institucional.unab.cl y se incluyen solo por conveniencia; su uso debe respetar el manual de marca y los protocolos de la institución.

## Licencia

[MIT](LICENSE).
