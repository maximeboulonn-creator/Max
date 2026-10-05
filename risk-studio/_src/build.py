"""Assemble les consoles autonomes : recopie le socle (core.js), le kit des consoles
(console.js, console.css) dans chaque page et écrit le résultat dans le dossier parent.
Usage : python3 build.py"""
import pathlib

SRC = pathlib.Path(__file__).resolve().parent
OUT = SRC.parent
PAGES = {
    "C1_Matrice.html": "RiskStudio_C1_Matrice.html",
    "C2_Fonds.html": "RiskStudio_C2_Fonds.html",
    "C3_Decisions.html": "RiskStudio_C3_Decisions.html",
    "C4_AlertePrecoce.html": "RiskStudio_C4_AlertePrecoce.html",
}
PARTS = {
    "/*@CORE*/": (SRC / "core.js").read_text(encoding="utf-8"),
    "/*@KIT*/": (SRC / "console.js").read_text(encoding="utf-8"),
    "/*@KITCSS*/": (SRC / "console.css").read_text(encoding="utf-8"),
}

for src, dst in PAGES.items():
    html = (SRC / src).read_text(encoding="utf-8")
    for marker, content in PARTS.items():
        assert marker in html, (src, marker)
        html = html.replace(marker, content)
    (OUT / dst).write_text(html, encoding="utf-8")
    print("écrit", dst)
