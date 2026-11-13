#!/usr/bin/env python3
"""fetch_logos.py - descarga los logos oficiales desde https://institucional.unab.cl con Playwright.

Uso:
  python3 fetch_logos.py                 # descarga los que falten en ../assets/logos
  python3 fetch_logos.py --solo fen      # solo uno (fen, unab-horizontal, unab-vertical)
  python3 fetch_logos.py --forzar        # vuelve a descargar aunque existan
  python3 fetch_logos.py --listar        # muestra los enlaces de imagen que encuentra (depuración)

Requiere: pip install playwright && playwright install chromium
Códigos de salida: 0 = todo listo, 1 = faltó algún logo, 2 = Playwright no disponible.
"""
import argparse, pathlib, re, sys

URL = "https://institucional.unab.cl"
DEST = pathlib.Path(__file__).resolve().parent.parent / "assets" / "logos"
# nombre corto -> (archivo local, patrones del enlace en la página; se prueban en orden)
TARGETS = {
    "fen": ("logo-fen.png", [r"/fen-1\.png$", r"/fen-\d\.png$"]),
    "unab-horizontal": ("logo-unab-horizontal.png", [r"/UNAB_horizontal\.png$", r"/unab[-_]horizontal[^/]*\.png$"]),
    "unab-vertical": ("logo-unab-vertical.png", [r"/UNAB_vertical_s_palabras\.png$", r"/unab[-_]vertical[^/]*\.png$"]),
}
JS = """els => els.flatMap(e => [e.href, e.src, e.getAttribute('data-src'), e.getAttribute('data-lazy-src')])
                 .filter(Boolean)"""


def pick(links, patterns):
    for pat in patterns:
        hits = sorted({l for l in links if re.search(pat, l, re.I)}, reverse=True)  # el más reciente primero
        if hits:
            return hits[0]
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--solo", choices=sorted(TARGETS))
    ap.add_argument("--forzar", action="store_true")
    ap.add_argument("--listar", action="store_true")
    ap.add_argument("--dest", default=str(DEST))
    a = ap.parse_args()
    dest = pathlib.Path(a.dest); dest.mkdir(parents=True, exist_ok=True)
    wanted = {k: v for k, v in TARGETS.items() if not a.solo or k == a.solo}
    todo = {k: v for k, v in wanted.items() if a.forzar or a.listar or not (dest / v[0]).exists()}
    if not todo:
        print("Logos ya presentes en", dest); return 0
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright no está instalado. Ejecuta: pip install playwright && playwright install chromium", file=sys.stderr)
        return 2
    missing = 0
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch()
        except Exception as e:  # navegador no instalado
            print("No se pudo abrir Chromium (%s). Ejecuta: playwright install chromium" % str(e).splitlines()[0], file=sys.stderr)
            return 2
        ctx = browser.new_context()
        page = ctx.new_page()
        for intento in range(3):  # la página a veces no llega a "networkidle": basta con cargar el DOM
            try:
                page.goto(URL, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(4000)
                break
            except Exception as e:
                if intento == 2:
                    print("No se pudo abrir %s (%s)" % (URL, str(e).splitlines()[0]), file=sys.stderr)
                    browser.close(); return 1
        links = sorted(set(page.eval_on_selector_all("a[href], img", JS)))
        if a.listar:
            for l in links:
                if re.search(r"\.(png|jpe?g|svg)$", l, re.I): print(l)
            browser.close(); return 0
        for name, (fname, patterns) in todo.items():
            link = pick(links, patterns)
            if not link:
                print("No encontré el logo '%s' en %s (¿cambió la página?). Usa --listar para ver los enlaces." % (name, URL), file=sys.stderr)
                missing += 1; continue
            r = ctx.request.get(link, timeout=60000)
            body = r.body() if r.ok else b""
            if not body.startswith(b"\x89PNG") or len(body) < 2000:
                print("Descarga inválida para '%s' (%s, HTTP %s)" % (name, link, r.status), file=sys.stderr)
                missing += 1; continue
            (dest / fname).write_bytes(body)
            print("OK  %-16s -> %s  (%d KB)  %s" % (name, dest / fname, len(body) // 1024, link))
        browser.close()
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
