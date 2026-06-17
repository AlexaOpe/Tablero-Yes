# -*- coding: utf-8 -*-
"""
Construye la PLANTILLA de la foto (foto-plantilla.html) A PARTIR del panel,
para que el panel y la foto NUNCA se desincronicen.

Funciona en DOS lugares (detecta solo dónde están los archivos):
  - En tu PC:  HTML/publicar/  (+ assets/)
  - En el repo de GitHub: archivos planos en la raíz

Toma: taskpane.html, styles.css, app.js, chart.umd.min.js, logo-marca.png
Produce una foto AUTOCONTENIDA, de SOLO LECTURA, con dos marcadores que el botón
"Descargar foto" reemplaza con los datos vivos:
    window.__FACT__ = "__FACT_PLACEHOLDER__";   (BaseDeDatos)
    window.__TC__   = "__TC_PLACEHOLDER__";      (tabla de tipo de cambio)

NO toca el Excel. NO descarga nada. Solo lee archivos locales y escribe el HTML.

Uso:
    python Recursos/construir_foto.py          (en tu PC)
    python construir_foto.py                   (en el repo / GitHub Action)
"""
import base64, os, re, shutil, sys

HERE = os.path.dirname(os.path.abspath(__file__))
# Lugares donde buscar los archivos del panel, en orden de prioridad.
SEARCH = [
    os.getcwd(),                                       # repo plano (raíz)
    HERE,                                              # junto al script
    os.path.join(HERE, "..", "HTML", "publicar"),     # tu PC: Recursos -> HTML/publicar
    os.path.join(HERE, "HTML", "publicar"),
]


def find(name):
    """Devuelve la ruta del archivo 'name' mirando en SEARCH y en sub 'assets'."""
    for d in SEARCH:
        for sub in ("", "assets"):
            p = os.path.normpath(os.path.join(d, sub, name))
            if os.path.isfile(p):
                return p
    print("FALTA el archivo:", name)
    sys.exit(1)


def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


def main():
    task_path = find("taskpane.html")
    html   = read(task_path)
    styles = read(find("styles.css"))
    appjs  = read(find("app.js"))
    chart  = read(find("chart.umd.min.js"))
    with open(find("logo-marca.png"), "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode("ascii")

    # La foto se escribe junto a taskpane.html (raíz del repo, o HTML/publicar en tu PC).
    out = os.path.join(os.path.dirname(task_path), "foto-plantilla.html")

    n = {}  # contador de reemplazos, para validar la estructura

    # 1) quitar Office.js (la foto no vive dentro de Excel)
    html, n['office'] = re.subn(
        r'[ \t]*<script src="https://appsforoffice\.microsoft\.com[^"]*"></script>\s*\n',
        '', html)

    # 2) incrustar el CSS en lugar del <link>
    html, n['css'] = re.subn(
        r'<link rel="stylesheet" href="styles\.css"\s*/>',
        lambda m: '<style>\n' + styles + '\n</style>', html)

    # 3) logo -> base64
    html, n['logo'] = re.subn(
        r'src="logo-marca\.png"',
        lambda m: 'src="data:image/png;base64,' + logo_b64 + '"', html)

    # 4) subtítulo: deja claro que es una foto (no "en vivo")
    html = html.replace(
        'Validación al momento de la captura · en vivo desde Excel',
        'Foto del tablero de ingresos · solo lectura')

    # 5) quitar el botón "Descargar foto" (la foto es estática)
    html, n['btn'] = re.subn(
        r'[ \t]*<button id="btnDownload".*?</button>\s*\n',
        '', html, flags=re.S)

    # 6) Chart.js + marcadores de datos + app.js, todo incrustado
    def bloque(_m):
        return (
            '  <script>\n' + chart + '\n</script>\n'
            '  <script>window.__FACT__ = "__FACT_PLACEHOLDER__"; '
            'window.__TC__ = "__TC_PLACEHOLDER__";</script>\n'
            '  <script>\n' + appjs + '\n</script>'
        )
    html, n['scripts'] = re.subn(
        r'<script src="chart\.umd\.min\.js"></script>\s*\n\s*<script src="app\.js"></script>',
        bloque, html)

    # Validación: cada reemplazo clave debe ocurrir exactamente 1 vez
    for k in ('css', 'logo', 'btn', 'scripts'):
        if n[k] != 1:
            print("ERROR: el reemplazo '%s' ocurrió %d veces (se esperaba 1)." % (k, n[k]))
            print("Revisa que taskpane.html no haya cambiado de estructura.")
            sys.exit(1)
    if '"__FACT_PLACEHOLDER__"' not in html or '"__TC_PLACEHOLDER__"' not in html:
        print("ERROR: faltan los marcadores de datos en la salida."); sys.exit(1)

    if os.path.exists(out):
        shutil.copy2(out, out + ".bak")
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)

    print("Foto generada: %s (%.0f KB)" % (out, os.path.getsize(out) / 1024))


if __name__ == "__main__":
    main()
