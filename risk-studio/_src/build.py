"""Recopie le socle commun (core.js) dans chaque page et écrit les pages autonomes
dans le dossier parent. Usage : python3 build.py"""
import pathlib

SRC = pathlib.Path(__file__).resolve().parent
OUT = SRC.parent
PAGES = {
    "A_Note.html": "RiskStudio_A_Note.html",
    "B_Classeur.html": "RiskStudio_B_Classeur.html",
    "C_Console.html": "RiskStudio_C_Console.html",
}

core = (SRC / "core.js").read_text(encoding="utf-8")
for src, dst in PAGES.items():
    path = SRC / src
    if not path.exists():
        continue
    html = path.read_text(encoding="utf-8")
    assert "/*@CORE*/" in html, src
    (OUT / dst).write_text(html.replace("/*@CORE*/", core), encoding="utf-8")
    print("écrit", dst)
