#!/usr/bin/env python3
"""build_docx.py - Markdown simplificado + datos.json -> .docx en APA 7 (UNAB).
Uso: python3 build_docx.py cuerpo.md datos.json salida.docx
"""
import json, os, re, subprocess, sys, unicodedata, datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONTS = {"Times New Roman": 12, "Calibri": 11, "Arial": 11, "Georgia": 11}
MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto",
         "septiembre","octubre","noviembre","diciembre"]
LINKED = ("libro","web","informe","tesis","video","podcast","webinar","presentacion","social","ley")

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGOS = {"fen": "logo-fen.png", "unab-horizontal": "logo-unab-horizontal.png", "unab-vertical": "logo-unab-vertical.png"}

def die(m): sys.exit("ERROR: " + m)

def resolve_logo(D, base):
    """Devuelve la ruta del logo o None. 'logo': 'fen' (defecto) | 'unab-horizontal' | 'unab-vertical' | ruta | false.
    Si falta el archivo incluido, lo descarga de institucional.unab.cl con scripts/fetch_logos.py (Playwright)."""
    v = D.get("logo", "fen")
    if v in (False, None, "", "ninguno"): return None
    if v in LOGOS:
        path = os.path.join(SKILL_DIR, "assets", "logos", LOGOS[v])
        if not os.path.exists(path):
            print("AVISO: falta %s; intentando descargarlo de institucional.unab.cl con Playwright..." % LOGOS[v], file=sys.stderr)
            subprocess.run([sys.executable, os.path.join(SKILL_DIR, "scripts", "fetch_logos.py"), "--solo", v])
        if os.path.exists(path): return path
        print("AVISO: no se pudo obtener el logo '%s'; la portada se genera sin logo." % v, file=sys.stderr)
        return None
    for cand in (v, os.path.join(base, v)):
        if os.path.exists(cand): return cand
    print("AVISO: no existe el archivo de logo '%s'; la portada se genera sin logo." % v, file=sys.stderr)
    return None
def norm(s):
    s = unicodedata.normalize("NFD", s.lower())
    return re.sub(r"[^a-z0-9]", "", "".join(c for c in s if unicodedata.category(c) != "Mn"))
def end(t): return t if t[-1:] in ".?!" else t + "."
def surname(a): return a.split(",")[0].strip()
def flip(a):
    s, _, i = a.partition(","); return (i.strip() + " " + s.strip()).strip()

def dates(f):
    d = str(f.get("fecha", "s.f."))
    if d == "s.f.": return "s.f.", "s.f."
    m = re.match(r"^(\d{4})(?:-(\d{2}))?(?:-(\d{2}))?$", d)
    if not m: die("fecha invalida en %s: %s" % (f.get("id"), d))
    y, mo, da = m.groups()
    full = y + (", " + (da.lstrip("0") + " de " if da else "") + MESES[int(mo)-1] if mo else "")
    return y, full

class Ctx:
    def __init__(self, fuentes, cited):
        self.F = {f["id"]: f for f in fuentes}
        self.cited = cited; self.seen = set(); self.suf = {}
        for k in cited:
            if k not in self.F: die("cita a fuente inexistente: " + k)
        groups = {}
        for k in cited:
            f = self.F[k]
            groups.setdefault((self.author_key(f), dates(f)[0]), []).append(k)
        for g in groups.values():
            if len(g) > 1:
                for i, k in enumerate(sorted(g, key=lambda k: norm(self.F[k]["titulo"]))):
                    self.suf[k] = "abcdefghijklmnopqrstuvwxyz"[i]
    def author_key(self, f): return norm("".join(f.get("autores", [])) or f.get("corporativo", "") or f["titulo"])
    def year(self, f):
        y = dates(f)[0]; s = self.suf.get(f["id"], "")
        return y + s if y != "s.f." else ("s.f." + ("-" + s if s else ""))
    def label(self, f, narrative):
        k = f["id"]; first = k not in self.seen; self.seen.add(k)
        if f.get("autores"):
            s = [surname(a) for a in f["autores"]]
            return [(s[0] if len(s) == 1 else s[0] + " y " + s[1] if len(s) == 2 else s[0] + " et al.", False)], None
        if f.get("corporativo"):
            n, sg = f["corporativo"], f.get("sigla")
            if not sg: return [(n, False)], None
            if first: return [(n, False)], sg + ("," if narrative else "")
            return [(sg, False)], None
        t = f.get("titulo_corto", f["titulo"])
        return ([(t, True)] if f["tipo"] in LINKED else [("“" + t + "”", False)]), None
    def one(self, key, loc, narrative):
        if key not in self.F: die("cita a fuente inexistente: " + key)
        f = self.F[key]; lab, sig = self.label(f, narrative); y = self.year(f)
        extra = y + (", " + loc if loc else "")
        if narrative:
            if sig: return lab + [(" (%s %s)" % (sig, extra), False)]
            return lab + [(" (%s)" % extra, False)]
        if sig is None and f.get("corporativo") and f.get("sigla") and lab[0][0] == f["corporativo"]:
            lab = [(f["corporativo"] + " [" + f["sigla"] + "]", False)]
        return lab + [(", " + extra, False)]
    def cite(self, tok):
        inner = tok[1:-1]
        if inner.startswith("n@"):
            k, _, loc = inner[2:].partition("|"); return self.one(k.strip(), loc.strip(), True)
        if inner.startswith("@"):
            parts = []
            for it in inner.split(";"):
                k, _, loc = it.strip().lstrip("@").partition("|")
                parts.append(self.one(k.strip(), loc.strip(), False))
            parts.sort(key=lambda r: norm("".join(t for t, _ in r)))
            out = [("(", False)]
            for i, r in enumerate(parts):
                if i: out.append(("; ", False))
                out += r
            return out + [(")", False)]
        if inner.startswith("loc|"): return [("(" + inner[4:] + ")", False)]
        if inner.startswith("cp|"): return [("(" + inner[3:] + ")", False)]
        die("marca de cita no reconocida: " + tok)

    # ---------- referencias ----------
    def authors(self, f):
        a = f.get("autores")
        if not a: return f.get("corporativo", "")
        if len(a) > 20: a = a[:19] + ["… " + a[-1]]; s = ", ".join(a[:-1]) + " " + a[-1]; r = s
        elif len(a) == 1: r = a[0]
        elif f.get("idioma") == "en": r = ", ".join(a[:-1]) + (", " if len(a) > 2 else " ") + "& " + a[-1]
        else: r = ", ".join(a[:-1]) + " y " + a[-1]
        return r
    def link(self, f):
        if f.get("doi"):
            d = f["doi"]; return d if d.startswith("http") else "https://doi.org/" + d
        return f.get("url", "")
    def reference(self, key):
        f = self.F[key]; t = f["tipo"]; R = []
        def add(*segs):
            for s in segs:
                if isinstance(s, tuple): R.append(s)
                elif s: R.append((s, False))
        A = self.authors(f)
        if f.get("rol"): A += " (%s)" % f["rol"]
        if f.get("usuario"): A += " [%s]" % f["usuario"]
        y, full = dates(f); d = full if t in ("web","periodico","video","podcast","social","ley") else y
        if y != "s.f." and self.suf.get(key): d = d.replace(y, y + self.suf[key], 1)
        if y == "s.f." and self.suf.get(key): d = "s.f.-" + self.suf[key]
        ti = f["titulo"]; it = t in LINKED
        if A:
            add(end(A) + " (%s). " % d)
            body = []
        else:
            add(); body = []
        ttl = (ti, True) if it else (end(ti) if False else ti)
        def title(extra=""):
            if it: add((ti, True)); add(extra + ". " if extra else ". ")
            else: add(end(ti) + " ")
        if not A:  # sin autor: el titulo pasa al lugar del autor
            if it: add((ti, True), ". (%s). " % d)
            else: add(end(ti) + " (%s). " % d)
        pub = f.get("editorial")
        if pub and pub == f.get("corporativo"): pub = None
        if t == "libro":
            if A: add((ti, True), (" (%s ed.)" % f["edicion"]) if f.get("edicion") else "", ". ")
            elif f.get("edicion"): R[-1] = (R[-1][0], R[-1][1]); add("(%s ed.). " % f["edicion"])
            add(end(pub) + " " if pub else "")
        elif t == "capitulo":
            ed = f.get("editores", [])
            en = ", ".join(flip(e) for e in ed[:-1]) + (" y " if len(ed) > 1 else "") + flip(ed[-1]) if ed else ""
            add(end(ti) + " ")
            add("En " + en + " (%s), " % ("Ed." if len(ed) == 1 else "Eds.") if ed else "En ", (f["libro"], True),
                " (%spp. %s). " % ((f["edicion"] + " ed., ") if f.get("edicion") else "", f["paginas"]))
            add(end(pub) + " " if pub else "")
        elif t == "revista":
            add(end(ti) + " ", (f["revista"], True))
            if f.get("volumen"): add(", ", (str(f["volumen"]), True))
            if f.get("numero"): add("(%s)" % f["numero"])
            if f.get("paginas"): add(", " + f["paginas"])
            add(". ")
        elif t == "periodico":
            add(end(ti) + " ", (f["periodico"], True), (", " + f["paginas"]) if f.get("paginas") else "", ". ")
        elif t == "web":
            if A: add((ti, True), ". ")
            s = f.get("sitio"); add(end(s) + " " if s and s != f.get("corporativo") else "")
            if f.get("recuperado"): add("Recuperado el %s de " % f["recuperado"])
        elif t == "informe":
            if A: add((ti, True), (" (%s)" % f["numero_informe"]) if f.get("numero_informe") else "", ". ")
            add(end(pub) + " " if pub else "")
        elif t == "tesis":
            if A: add((ti, True))
            add(" [%s, %s]. " % (f["tipo_tesis"], f["institucion"]))
            add(end(f["repositorio"]) + " " if f.get("repositorio") else "")
        elif t in ("video","podcast","webinar","presentacion","social"):
            medio = f.get("medio", {"video":"Video","podcast":"Audio en podcast","webinar":"Webinar",
                                    "presentacion":"Diapositivas","social":"Publicación"}[t])
            if A: add((ti, True))
            add(" [%s]. " % medio)
            sitio = f.get("plataforma") or f.get("sitio") or pub
            add(end(sitio) + " " if sitio else "")
        elif t == "ley":
            if A: add((ti, True), (" (%s)" % f["numero"]) if f.get("numero") else "", ". ")
            add(end(f["fuente"]) + " " if f.get("fuente") else "")
        else: die("tipo de fuente desconocido: " + t)
        lk = self.link(f)
        if lk: add(lk)
        if R and R[-1][0].endswith(" "): R[-1] = (R[-1][0].rstrip(), R[-1][1])
        return R

    def sort_key(self, k):
        f = self.F[k]; return (self.author_key(f), dates(f)[0], self.suf.get(k, ""))

# ---------------- documento ----------------
def set_fonts(style, name, size, bold=None, italic=None):
    f = style.font; f.name = name; f.size = Pt(size); f.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in list(rf.attrib):
        if "Theme" in a: del rf.attrib[a]
    for a in ("ascii","hAnsi","cs","eastAsia"): rf.set(qn("w:" + a), name)
    if bold is not None: f.bold = bold
    if italic is not None: f.italic = italic

def fmt(style, align=None, first=0.0, left=0.0, line=2.0, keep=False):
    pf = style.paragraph_format
    pf.space_before = pf.space_after = Pt(0); pf.line_spacing = line
    pf.first_line_indent = Inches(first); pf.left_indent = Inches(left)
    if align is not None: pf.alignment = align
    pf.keep_with_next = keep; pf.widow_control = True

def setup(doc, name):
    size = FONTS[name]
    dd = doc.styles.element.find(qn("w:docDefaults"))
    for rf in dd.iter(qn("w:rFonts")):
        for a in list(rf.attrib): del rf.attrib[a]
        for a in ("ascii","hAnsi","cs","eastAsia"): rf.set(qn("w:" + a), name)
    lang = dd.find(".//" + qn("w:lang"))
    if lang is not None: lang.set(qn("w:val"), "es-CL")
    S = doc.styles
    set_fonts(S["Normal"], name, size); fmt(S["Normal"], AL.LEFT, first=0.5)
    cfg = {1: (AL.CENTER, True, False), 2: (AL.LEFT, True, False), 3: (AL.LEFT, True, True)}
    for n, (al, b, i) in cfg.items():
        st = S["Heading %d" % n]; set_fonts(st, name, size, b, i); fmt(st, al, keep=True)
    set_fonts(S["Header"], name, size); fmt(S["Header"], AL.RIGHT, line=1.0)
    for nm, left, first in (("Cita en bloque", 0.5, 0.0), ("Referencia", 0.5, -0.5)):
        st = S.add_style(nm, 1); st.base_style = S["Normal"]; fmt(st, AL.LEFT, first, left)
    fmt(S["List Bullet"], AL.LEFT, 0.0, 0.5) if "List Bullet" in [s.name for s in S] else None
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Inches(1)
    sec.header_distance = Inches(0.5)
    p = sec.header.paragraphs[0]; p.style = S["Header"]; p.alignment = AL.RIGHT
    for typ, txt in (("begin", None), (None, " PAGE "), ("end", None)):
        r = p.add_run()
        if typ:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), typ)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt
        r._r.append(e)

TOK = re.compile(r"(\[(?:n@|@|loc\||cp\|)[^\]]*\]|\*\*[^*]+\*\*|\*[^*\s][^*]*\*)")
def inline(par, text, ctx, bold=False, italic=False):
    for part in TOK.split(text):
        if not part: continue
        if part.startswith("[") and TOK.fullmatch(part):
            for t, it in ctx.cite(part): r = par.add_run(t); r.italic = it or italic or None; r.bold = bold or None
        elif part.startswith("**"): r = par.add_run(part[2:-2]); r.bold = True; r.italic = italic or None
        elif part.startswith("*") and len(part) > 2 and TOK.fullmatch(part):
            r = par.add_run(part[1:-1]); r.italic = True; r.bold = bold or None
        else: r = par.add_run(part); r.bold = bold or None; r.italic = italic or None

def para(doc, style=None, align=None, first=None, before=False, keep=False, space=None):
    p = doc.add_paragraph(style=style)
    if align is not None: p.alignment = align
    if first is not None: p.paragraph_format.first_line_indent = Inches(first)
    if before: p.paragraph_format.page_break_before = True
    if keep: p.paragraph_format.keep_with_next = True
    if space: p.paragraph_format.space_before = Pt(space)
    return p

def border(el_parent, tag, **kw):
    b = OxmlElement("w:" + tag)
    for k, v in kw.items(): b.set(qn("w:" + k), str(v))
    el_parent.append(b)

def table(doc, rows, ctx):
    n = len(rows[0]); t = doc.add_table(rows=len(rows), cols=n); t.autofit = False
    tpr = t._tbl.tblPr; tb = OxmlElement("w:tblBorders")
    for side in ("top","bottom"): border(tb, side, val="single", sz=8, space=0, color="000000")
    for side in ("left","right","insideH","insideV"): border(tb, side, val="nil")
    tpr.append(tb)
    for i, row in enumerate(rows):
        tr = t.rows[i]
        if i == 0:
            h = OxmlElement("w:tblHeader"); tr._tr.get_or_add_trPr().append(h)
        for j in range(n):
            c = tr.cells[j]; c.width = Inches(6.5 / n)
            p = c.paragraphs[0]; pf = p.paragraph_format
            pf.first_line_indent = Inches(0); pf.line_spacing = 1.0
            p.alignment = AL.CENTER if i == 0 else AL.LEFT
            inline(p, row[j] if j < len(row) else "", ctx)
            if i == 0:
                tcb = OxmlElement("w:tcBorders"); border(tcb, "bottom", val="single", sz=6, space=0, color="000000")
                c._tc.get_or_add_tcPr().append(tcb)

def build(md_path, data_path, out):
    D = json.load(open(data_path, encoding="utf-8")); md = open(md_path, encoding="utf-8").read()
    name = D.get("fuente", "Times New Roman")
    if name not in FONTS: die("fuente no permitida: %s (use %s)" % (name, ", ".join(FONTS)))
    corpus = md + " " + D.get("resumen", "")
    cited = list(dict.fromkeys(re.findall(r"@([A-Za-z0-9_\-]+)", corpus)))
    ctx = Ctx(D.get("fuentes", []), cited)
    doc = Document(); setup(doc, name); P = D["portada"]
    # ---- portada ----
    logo = resolve_logo(D, os.path.dirname(os.path.abspath(data_path)))
    if logo:
        p = para(doc, align=AL.CENTER, first=0); p.paragraph_format.line_spacing = 1.0
        p.add_run().add_picture(logo, width=Inches(float(D.get("logo_ancho", 3.0))))
    if D.get("marca"):
        p = para(doc, first=0); p.paragraph_format.line_spacing = 1.0
        bd = OxmlElement("w:pBdr"); border(bd, "bottom", val="single", sz=12, space=1, color="A90429")
        p._p.get_or_add_pPr().append(bd)
    p = para(doc, align=AL.CENTER, first=0, space=72); r = p.add_run(P["titulo"]); r.bold = True
    autores = P["autores"] if isinstance(P["autores"], list) else [P["autores"]]
    lines = [", ".join(autores) if len(autores) < 3 else ", ".join(autores[:-1]) + " y " + autores[-1]]
    lines += [P.get("afiliacion", "Universidad Andrés Bello")]
    if P.get("programa"): lines.append(P["programa"])
    lines += [P.get("curso", ""), P.get("profesor", "")]
    hoy = datetime.date.today(); lines.append(P.get("fecha") or "%d de %s de %d" % (hoy.day, MESES[hoy.month-1], hoy.year))
    for i, l in enumerate([x for x in lines if x]):
        p = para(doc, align=AL.CENTER, first=0, space=24 if i == 0 else None); p.add_run(l)
    first_break = True
    # ---- resumen ----
    if D.get("resumen"):
        p = para(doc, "Heading 1", before=True); p.add_run("Resumen")
        p = para(doc, first=0); inline(p, D["resumen"], ctx)
        if D.get("palabras_clave"):
            p = para(doc, first=0.5); r = p.add_run("Palabras clave:"); r.italic = True
            p.add_run(" " + ", ".join(D["palabras_clave"]))
        first_break = False
    # ---- cuerpo ----
    p = para(doc, "Heading 1", before=True); p.add_run(P["titulo"])
    n_ap = len(re.findall(r"^::apendice", md, re.M)); ap_i = -1; ap_letter = ""
    cnt = {"Tabla": 0, "Figura": 0}; runin = None; prev_quote = False; pend_break = False
    def caption(kind, title):
        cnt[kind] += 1
        num = "%s%d" % (ap_letter, cnt[kind]) if ap_letter else str(cnt[kind])
        p = para(doc, first=0, keep=True); p.add_run("%s %s" % (kind, num)).bold = True
        p = para(doc, first=0, keep=True); p.add_run(title).italic = True
    def note(text):
        p = para(doc, first=0); p.add_run("Nota.").italic = True; inline(p, " " + text, ctx)
    blocks = re.split(r"\n\s*\n", md.strip())
    for b in blocks:
        lines = [l.rstrip() for l in b.split("\n") if l.strip()]
        if not lines: continue
        h = lines[0]; m = re.match(r"^(#{1,5}) (.+)$", h)
        if runin and not (m or h.startswith(("::", ">", "- ", "|"))):
            p = para(doc, first=0.5, before=pend_break); pend_break = False
            r = p.add_run(runin[0]); r.bold = True; r.italic = runin[1]; p.add_run(" ")
            inline(p, " ".join(lines), ctx); runin = None; prev_quote = False; continue
        if runin:
            p = para(doc, first=0.5); r = p.add_run(runin[0]); r.bold = True; r.italic = runin[1]; runin = None
        if m:
            lvl, txt = len(m.group(1)), m.group(2).strip(); prev_quote = False
            if lvl <= 3: p = para(doc, "Heading %d" % lvl); p.add_run(txt)
            else: runin = (end(txt), lvl == 5)
        elif h.startswith(">"):
            txt = " ".join(l.lstrip("> ").strip() for l in lines)
            p = para(doc, "Cita en bloque", first=0.5 if prev_quote else 0.0); inline(p, txt, ctx); prev_quote = True
        elif h.startswith("::tabla"):
            prev_quote = False; caption("Tabla", h[7:].strip())
            rows = [[c.strip() for c in l.strip().strip("|").split("|")] for l in lines[1:]
                    if l.startswith("|") and not re.match(r"^\|[\s\-:|]+\|?$", l)]
            table(doc, rows, ctx)
            for l in lines[1:]:
                if l.startswith("::nota"): note(l[6:].strip())
        elif h.startswith("::figura"):
            prev_quote = False; path, _, title = h[8:].partition("|"); caption("Figura", title.strip())
            p = para(doc, align=AL.CENTER, first=0); p.paragraph_format.line_spacing = 1.0
            p.add_run().add_picture(path.strip(), width=Inches(5.5))
            for l in lines[1:]:
                if l.startswith("::nota"): note(l[6:].strip())
        elif h.startswith("::nota"): note(h[6:].strip())
        elif h.startswith("::apendice"):
            ap_i += 1; ap_letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[ap_i] if n_ap > 1 else "A"; cnt = {"Tabla": 0, "Figura": 0}
            p = para(doc, "Heading 1", before=True); p.add_run("Apéndice" + (" " + ap_letter if n_ap > 1 else ""))
            p = para(doc, "Heading 1"); p.add_run(h[10:].strip())
        elif h.startswith("::pagina"): pend_break = True
        elif h.startswith("- "):
            for l in lines:
                p = para(doc, first=0); p.paragraph_format.left_indent = Inches(0.5); p.paragraph_format.first_line_indent = Inches(-0.25)
                inline(p, "• " + l[2:], ctx)
        else:
            p = para(doc, before=pend_break); pend_break = False; inline(p, " ".join(lines), ctx); prev_quote = False
    if runin: p = para(doc, first=0.5); r = p.add_run(runin[0]); r.bold = True; r.italic = runin[1]
    # ---- referencias (antes de apendices: se reubican) ----
    refs = sorted(ctx.cited, key=ctx.sort_key)
    sunused = [f["id"] for f in D.get("fuentes", []) if f["id"] not in ctx.cited]
    if sunused: print("AVISO: fuentes sin citar (no se incluyen):", ", ".join(sunused), file=sys.stderr)
    body = doc.element.body; sect = body[-1]; mark = len(body) - 1
    p = para(doc, "Heading 1", before=True); p.add_run("Referencias")
    for k in refs:
        p = para(doc, "Referencia")
        for t, it in ctx.reference(k): r = p.add_run(t); r.italic = it or None
    if n_ap:  # mover apendices tras las referencias
        els = list(body)[:-1]; idx = [i for i, e in enumerate(els) if e is p._p or False]
        ap_start = next((i for i, e in enumerate(els) if e.xpath("string(.)").startswith("Apéndice")
                         and e.find(".//" + qn("w:pageBreakBefore")) is not None), None)
        ref_start = next(i for i, e in enumerate(els) if e.xpath("string(.)") == "Referencias")
        if ap_start is not None and ap_start < ref_start:
            seg_ap, seg_ref = els[ap_start:ref_start], els[ref_start:]
            for e in seg_ap + seg_ref: body.remove(e)
            for e in seg_ref + seg_ap: sect.addprevious(e)
    doc.core_properties.title = P["titulo"]; doc.core_properties.author = autores[0]
    doc.save(out); print("OK:", out, "| referencias:", len(refs))

if __name__ == "__main__":
    if len(sys.argv) != 4: sys.exit(__doc__)
    build(*sys.argv[1:])
