# -*- coding: utf-8 -*-
"""Ajoute le type de chaque véhicule au registre KPI_Registre_Incidents_Fundcraft_France.

- Données : colonne « Type de véhicule » à côté de la liste des fonds (plage nommée Ref_Type).
- Registre : colonne calculée « Type de véhicule » après « Catégorie », lue dans le référentiel
  (aucune double saisie ; placée hors des plages du contrôle de saisie).
Usage : python type_vehicule.py source.xlsx sortie.xlsx sortie_redline.xlsx
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from complete_france import (Calc, conv, find_col, FIRST, LAST, matrices_origine,  # noqa: E402
                             restaurer_matrices)

TYPES = {
    "Fundcraft France SAS": "SGP (transverse)",
    "AirFund Conviction Value Capital": "FoF PE Evergreen",
    "Openstone Infraworld": "FoF Infra Evergreen",
    "Aletheon Growth III S.L.P.": "PE Closed-End",
    "Partners Group PEO Eltif Feeder": "PE Evergreen Eltif Feeder",
    "Otentiq Private Equity X": "FoF - Umbrella Access X",
}


def referentiel(k):
    sh = k.sheet("Données")
    cf = find_col(sh, "Fonds et véhicules")
    c = cf + 1
    sh.Columns.insertByIndex(c, 1)
    k.copy(sh, cf, 4, cf, 15, c, 4)
    sh.getCellByPosition(c, 4).setString("")
    sh.getCellByPosition(c, 5).setString("Type de véhicule")
    sh.Columns.getByIndex(c).Width = int(sh.Columns.getByIndex(cf).Width * 0.65)
    for r in range(6, 16):
        fonds = sh.getCellByPosition(cf, r).getString()
        x = sh.getCellByPosition(c, r)
        x.setString(TYPES.get(fonds, ""))
    k.mark("Données", c, 5, c, 15)
    lettre = chr(65 + c)
    k.name("Ref_Type", f"$Données.${lettre}$7:${lettre}$16")


def registre(k):
    sh = k.sheet("Registre")
    cat = find_col(sh, "Catégorie")
    c = cat + 1
    sh.Columns.insertByIndex(c, 1)
    k.copy(sh, cat, 4, cat, LAST - 1, c, 4)
    sh.getCellRangeByPosition(c, FIRST - 1, c, LAST - 1).clearContents(1 | 2 | 4 | 16)
    sh.getCellByPosition(c, 5).setString("Type de véhicule")
    sh.Columns.getByIndex(c).Width = sh.Columns.getByIndex(cat).Width
    cfonds = chr(65 + find_col(sh, "Fonds ou véhicule"))
    rng = sh.getCellRangeByPosition(c, FIRST - 1, c, LAST - 1)
    f = '=IF(${C}{r}="","",IFERROR(INDEX(Ref_Type,MATCH(${C}{r},Ref_Fonds,0)),"À vérifier"))'
    rng.setFormulaArray(tuple((conv(f.format(C=cfonds, r=r)),) for r in range(FIRST, LAST + 1)))
    k.mark("Registre", c, 4, c, LAST - 1)


def main(src, out, out_red):
    matrices = matrices_origine(src)
    k = Calc(src)
    try:
        referentiel(k)
        registre(k)
        restaurer_matrices(k, matrices)
        k.save(out)
        for s, l, t, r, b in k.red:
            k.sheet(s).getCellRangeByPosition(l, t, r, b).CellBackColor = 0xFFFF00
        k.save(out_red)
    finally:
        k.close()


if __name__ == "__main__":
    main(*sys.argv[1:4])
