#!/usr/bin/env python3
"""validate_docx.py - revisa un .docx contra APA 7 (UNAB). Uso: python3 validate_docx.py archivo.docx"""
import re, sys
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml.ns import qn

OK_FONTS = {"Times New Roman": 12, "Calibri": 11, "Arial": 11, "Georgia": 11, "Lucida Sans Unicode": 10}
F = []
def err(m): F.append(("ERROR", m))
def warn(m): F.append(("AVISO", m))

def style_attr(st, getter):
    while st is not None:
        v = getter(st)
        if v is not None: return v
        st = st.base_style
    return None
def default_font(doc):
    rf = doc.styles.element.find(".//" + qn("w:docDefaults") + "//" + qn("w:rFonts"))
    return rf.get(qn("w:ascii")) if rf is not None else None
def run_font(doc, p, r):
    return r.font.name or style_attr(p.style, lambda s: s.font.name) or default_font(doc)
def run_size(p, r):
    v = r.font.size or style_attr(p.style, lambda s: s.font.size)
    return v.pt if v else None
def spacing(p):
    return p.paragraph_format.line_spacing or style_attr(p.style, lambda s: s.paragraph_format.line_spacing)
def indent(p, attr):
    v = getattr(p.paragraph_format, attr)
    if v is None: v = style_attr(p.style, lambda s: getattr(s.paragraph_format, attr))
    return round(v.inches, 2) if v is not None else 0.0
def align(p):
    return p.alignment if p.alignment is not None else style_attr(p.style, lambda s: s.paragraph_format.alignment)

def main(path):
    doc = Document(path); sec = doc.sections[0]
    # --- pagina
    if abs(sec.page_width.inches - 8.5) > .02 or abs(sec.page_height.inches - 11) > .02: err("El papel no es tamaño carta (8,5 x 11 pulgadas).")
    for n in ("left_margin","right_margin","top_margin","bottom_margin"):
        if abs(getattr(sec, n).inches - 1) > .02: err("Margen %s = %.2f pulg (debe ser 1 pulg = 2,54 cm)." % (n, getattr(sec, n).inches))
    # --- numero de pagina
    hx = sec.header._element.xml
    if "PAGE" not in hx: err("El encabezado no tiene número de página automático (campo PAGE).")
    elif sec.header.paragraphs[0].alignment != AL.RIGHT: err("El número de página no está alineado a la derecha.")
    # --- fuentes y parrafos
    fonts = set(); body = []; started = False
    for p in doc.paragraphs:
        body.append(p)
        for r in p.runs:
            if r.text.strip():
                fonts.add((run_font(doc, p, r), run_size(p, r)))
    names = {n for n, _ in fonts}
    names.discard(None)
    if len(names) > 1: err("Se usa más de una fuente: %s." % ", ".join(sorted(map(str, names))))
    for n, s in fonts:
        if n is None: continue
        if n not in OK_FONTS: err("Fuente no permitida por APA 7: %s." % n)
        elif s and abs(s - OK_FONTS[n]) > .1 and not s in (8,9,10,11,12,14): err("Tamaño %s pt no corresponde a %s (%s pt)." % (s, n, OK_FONTS[n]))
        elif s and n in OK_FONTS and abs(s - OK_FONTS[n]) > .1: warn("Tamaño %s pt en %s (APA pide %s pt; solo se acepta menor en notas al pie/figuras)." % (s, n, OK_FONTS[n]))
    blanks = 0; first_text = True
    for i, p in enumerate(body):
        t = p.text.strip(); sn = p.style.name
        if not t:
            if p._p.xpath(".//w:drawing"): blanks = 0; continue
            blanks += 1
            if blanks > 1: err("Párrafos vacíos consecutivos (párrafo %d): no use saltos de línea extra." % i); blanks = 0
            continue
        blanks = 0
        if align(p) == AL.JUSTIFY: err("Texto justificado (párrafo %d): APA pide alineación a la izquierda." % i)
        sp = spacing(p) or 1.0
        if abs(float(sp) - 2.0) > .05 and not p._p.xpath(".//w:drawing") and sn != "Header":
            err("Interlineado %.2f en «%s…» (debe ser doble)." % (float(sp), t[:30]))
        if sn.startswith("Heading"):
            lvl = int(sn.split()[-1]); r0 = p.runs[0] if p.runs else None
            b = style_attr(p.style, lambda s: s.font.bold); it = style_attr(p.style, lambda s: s.font.italic)
            if not (b or (r0 and r0.bold)): err("Título nivel %d sin negrita: «%s»." % (lvl, t[:30]))
            if lvl == 1 and align(p) != AL.CENTER: err("Título nivel 1 debe ir centrado: «%s»." % t[:30])
            if lvl in (2, 3) and align(p) != AL.LEFT: err("Título nivel %d debe ir a la izquierda: «%s»." % (lvl, t[:30]))
            if lvl == 3 and not it: err("Título nivel 3 debe ir en cursiva: «%s»." % t[:30])
            if re.match(r"^\s*(\d+[\.\)]|[A-Z][\.\)]|\d+\.\d+)\s", t): warn("Título numerado «%s»: APA no numera títulos." % t[:30])
            if lvl > 1 and t.isupper(): warn("Título en mayúsculas sostenidas: «%s» (use capitalización de título)." % t[:30])
        elif sn == "Normal":
            fi = indent(p, "first_line_indent"); has_img = p._p.xpath(".//w:drawing")
            prev = body[i-1].text.strip() if i else ""
            if not has_img and prev != "Resumen" and len(t) > 80 and abs(fi - 0.5) > .02 and not t.startswith(("Nota.", "Palabras clave")) and align(p) != AL.CENTER and indent(p, "left_indent") == 0:
                warn("Párrafo sin sangría de 1,27 cm: «%s…» (permitido solo en resumen)." % t[:30])
            words = len(t.split())
            for q in re.findall(r"“([^”]+)”", t):
                if len(q.split()) >= 40: err("Cita de 40 o más palabras dentro del párrafo: debe ir en bloque (sin comillas, sangría 1,27 cm).")
        if sn == "Cita en bloque" and "“" in t: warn("Cita en bloque con comillas: se omiten en citas largas.")
    # --- referencias
    refs = []; in_ref = False
    for p in body:
        if p.style.name == "Heading 1": in_ref = p.text.strip().lower() in ("referencias", "referencias bibliográficas") and p.text.strip() == "Referencias"; 
        if p.style.name.startswith("Heading 1") and p.text.strip() != "Referencias" and in_ref is False: pass
        if in_ref and p.style.name == "Referencia": refs.append(p)
    hdr = [p.text.strip() for p in body if p.style.name == "Heading 1"]
    if "Referencias" not in hdr: err("No existe la página «Referencias» (título exacto, negrita y centrado).")
    keys = []
    for p in refs:
        t = p.text
        if abs(indent(p, "left_indent") - 0.5) > .02 or abs(indent(p, "first_line_indent") + 0.5) > .02: err("Referencia sin sangría francesa de 1,27 cm: «%s…»." % t[:30])
        if " & " in t: warn("«&» en referencia: en español APA/UNAB usa «y» (solo «&» si la fuente está en inglés): «%s…»." % t[:30])
        if "Recuperado" in t: warn("«Recuperado el»: APA 7 lo usa solo si el contenido cambia con el tiempo y no está archivado: «%s…»." % t[:30])
        m = re.match(r"^(.+?)\.?\s+\((\d{4}[a-z]?|s\.f\.[-a-z]*)[,)]", t)
        if not m: warn("No se pudo leer autor/fecha en la referencia: «%s…»." % t[:40]); continue
        keys.append((re.split(r"[,.\[]", m.group(1))[0].strip().lower(), m.group(2)))
    plain = [re.sub(r"[^a-z0-9]", "", p.text.lower()) for p in refs]
    if plain != sorted(plain): warn("La lista de referencias no está en orden alfabético letra por letra.")
    text = " ".join(p.text for p in body if p.style.name != "Referencia" and p.text.strip() != "Referencias")
    cites = set(); siglas = set()
    for g in re.findall(r"\(([^()]*?(?:\d{4}[a-z]?|s\.f\.)[^()]*?)\)", text):
        for part in g.split(";"):
            m = re.match(r"^\s*(.+?),?\s+(\d{4}[a-z]?|s\.f\.[-a-z]*)", part)
            if m:
                lab = re.split(r"[ ,\[]", m.group(1).strip())[0]
                if lab.isupper() and len(lab) > 1: siglas.add(lab.lower())
                cites.add((lab.lower(), m.group(2)))
    for m in re.finditer(r"([A-ZÁÉÍÓÚÑ][\wáéíóúñ'’\-]+)(?: et al\.| y [A-ZÁÉÍÓÚÑ][\wáéíóúñ'’\-]+)? \((?:[A-Z]+, )?(\d{4}[a-z]?|s\.f\.)", text):
        cites.add((m.group(1).lower(), m.group(2)))
    have = {(k[0].split()[0] if k[0] else "", k[1]) for k in keys}
    for c in sorted(cites - have):
        if c[0] not in ("p", "pp") and c[0] not in siglas: warn("Cita sin referencia (revisar): %s, %s." % c)
    for c in sorted(have - cites): warn("Referencia sin cita en el texto (revisar; puede ser sigla de autor corporativo): %s, %s." % c)
    # --- tablas
    for t in doc.tables:
        x = t._tbl.xml
        if 'w:insideV w:val="nil"' not in x and "insideV" in x: err("Tabla con bordes verticales: APA solo usa líneas horizontales.")
        if "tblBorders" not in x: warn("Tabla sin bordes superior/inferior definidos.")
    n_err = sum(1 for l, _ in F if l == "ERROR"); n_warn = len(F) - n_err
    for l, m in F: print("[%s] %s" % (l, m))
    print("\nResumen: %d errores, %d avisos | %d referencias, %d citas distintas." % (n_err, n_warn, len(keys), len(cites)))
    return 1 if n_err else 0

if __name__ == "__main__":
    if len(sys.argv) != 2: sys.exit(__doc__)
    sys.exit(main(sys.argv[1]))
