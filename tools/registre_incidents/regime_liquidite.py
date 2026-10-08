# -*- coding: utf-8 -*-
"""Ajoute le régime de liquidité (evergreen / closed-end) au registre France et une ventilation
des incidents par régime au Dashboard.

- Données : colonne « Régime de liquidité » à côté du type de véhicule (Ref_Regime, liste L_Regime).
- Registre : colonne calculée « Régime de liquidité » (R_Regime), hors des plages du contrôle de saisie.
- Dashboard : bloc « Ventilation par régime de liquidité - 12 mois glissants » sous la ventilation par fonds.
Usage : python regime_liquidite.py source.xlsx sortie.xlsx sortie_redline.xlsx
"""
import os
import sys

import uno

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from complete_france import (Calc, conv, find_col, find_row, FIRST, LAST, matrices_origine,  # noqa: E402
                             restaurer_matrices)

REGIMES = ["Evergreen", "Closed-end", "SGP (transverse)"]
NON_RENSEIGNE = "Non renseigné"
PAR_FONDS = {
    "Fundcraft France SAS": "SGP (transverse)",
    "AirFund Conviction Value Capital": "Evergreen",
    "Openstone Infraworld": "Evergreen",
    "Aletheon Growth III S.L.P.": "Closed-end",
    "Partners Group PEO Eltif Feeder": "Evergreen",
    "Otentiq Private Equity X": "",  # à confirmer
}


def lettre(c):
    return chr(65 + c) if c < 26 else chr(64 + c // 26) + chr(65 + c % 26)


def trouver(sh, texte, lignes=60, colonnes=40):
    for r in range(lignes):
        for c in range(colonnes):
            if sh.getCellByPosition(c, r).getString() == texte:
                return c, r
    raise KeyError(texte)


def liste_validation(rng, nom):
    v = rng.Validation
    v.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")
    v.Formula1 = nom
    v.ShowList = 1
    v.IgnoreBlankCells = True
    v.ShowErrorMessage = True
    v.ErrorTitle = "Valeur hors liste"
    v.ErrorMessage = "Choisir un régime de la liste."
    rng.Validation = v


def donnees(k):
    sh = k.sheet("Données")
    ct = find_col(sh, "Type de véhicule")
    cf = find_col(sh, "Fonds et véhicules")
    c = ct + 1
    sh.Columns.insertByIndex(c, 1)
    k.copy(sh, ct, 4, ct, 15, c, 4)
    sh.getCellByPosition(c, 5).setString("Régime de liquidité")
    sh.Columns.getByIndex(c).Width = int(sh.Columns.getByIndex(ct).Width * 0.7)
    for r in range(6, 16):
        sh.getCellByPosition(c, r).setString(PAR_FONDS.get(sh.getCellByPosition(cf, r).getString(), ""))
    k.mark("Données", c, 5, c, 15)
    k.name("Ref_Regime", f"$Données.${lettre(c)}$7:${lettre(c)}$16")
    # liste des régimes, sous la liste des types de dépassement
    cl, rl = trouver(sh, "Type de dépassement")
    r0 = rl + 4
    k.copy(sh, cl, rl, cl, rl + 1, cl, r0)
    k.copy(sh, cl, rl + 1, cl, rl + 1, cl, r0 + 2)
    k.copy(sh, cl, rl + 1, cl, rl + 1, cl, r0 + 3)
    sh.getCellByPosition(cl, r0).setString("Régime de liquidité")
    for i, reg in enumerate(REGIMES):
        sh.getCellByPosition(cl, r0 + 1 + i).setString(reg)
    k.mark("Données", cl, r0, cl, r0 + len(REGIMES))
    k.name("L_Regime", f"$Données.${lettre(cl)}${r0 + 2}:${lettre(cl)}${r0 + 1 + len(REGIMES)}")
    liste_validation(sh.getCellRangeByPosition(c, 6, c, 15), "L_Regime")


def registre(k):
    sh = k.sheet("Registre")
    ct = find_col(sh, "Type de véhicule")
    c = ct + 1
    sh.Columns.insertByIndex(c, 1)
    k.copy(sh, ct, 4, ct, LAST - 1, c, 4)
    sh.getCellRangeByPosition(c, FIRST - 1, c, LAST - 1).clearContents(1 | 2 | 4 | 16)
    sh.getCellByPosition(c, 5).setString("Régime de liquidité")
    sh.Columns.getByIndex(c).Width = sh.Columns.getByIndex(ct).Width
    cf = lettre(find_col(sh, "Fonds ou véhicule"))
    f = (f'=IF(${cf}{{r}}="","",IFERROR(IF(INDEX(Ref_Regime,MATCH(${cf}{{r}},Ref_Fonds,0))="","{NON_RENSEIGNE}",'
         f'INDEX(Ref_Regime,MATCH(${cf}{{r}},Ref_Fonds,0))),"{NON_RENSEIGNE}"))')
    sh.getCellRangeByPosition(c, FIRST - 1, c, LAST - 1).setFormulaArray(
        tuple((conv(f.format(r=r)),) for r in range(FIRST, LAST + 1)))
    k.name("R_Regime", f"$Registre.${lettre(c)}${FIRST}:${lettre(c)}${LAST}")
    k.mark("Registre", c, 4, c, LAST - 1)


def dashboard(k):
    sh = k.sheet("Dashboard")
    band = find_row(sh, "Ventilation par fonds ou véhicule - 12 mois glissants")
    hdr = band + 1
    total = next(r for r in range(hdr, hdr + 40) if sh.getCellByPosition(1, r).getString() == "Total")
    n = len(REGIMES) + 1
    sh.Rows.insertByIndex(total + 1, n + 4)
    rb, rh, r1 = total + 2, total + 3, total + 4
    rt = r1 + n
    k.copy(sh, 1, band, 9, band, 1, rb)
    sh.getCellByPosition(1, rb).setString("Ventilation par régime de liquidité - 12 mois glissants")
    k.copy(sh, 1, hdr, 9, hdr, 1, rh)
    sh.getCellByPosition(1, rh).setString("Régime de liquidité")
    sh.Rows.getByIndex(rh).Height = sh.Rows.getByIndex(hdr).Height
    H = rh + 1  # ligne d'en-tête en 1-based
    per = 'R_Det,">="&Debut12M,R_Det,"<"&FinTx'
    for i in range(n):
        r = r1 + i
        R = r + 1
        k.copy(sh, 1, hdr + 1, 9, hdr + 1, 1, r)
        b = sh.getCellByPosition(1, r)
        if i < len(REGIMES):
            b.setFormula(conv(f"=INDEX(L_Regime,{i + 1})"))
        else:
            b.setString(NON_RENSEIGNE)
        for c, col in enumerate("CDEF"):
            sh.getCellByPosition(2 + c, r).setFormula(conv(
                f'=IF($B{R}="","",COUNTIFS(R_Regime,$B{R},R_Cat,{col}${H},{per}))'))
        sh.getCellByPosition(6, r).setFormula(conv(f'=IF($B{R}="","",SUM(C{R}:F{R}))'))
        sh.getCellByPosition(7, r).setFormula(conv(f'=IF($B{R}="","",COUNTIFS(R_Regime,$B{R},R_Grav,">=3",{per}))'))
        sh.getCellByPosition(8, r).setFormula(conv(f'=IF($B{R}="","",SUMIFS(R_PerteNette,R_Regime,$B{R},{per}))'))
        sh.getCellByPosition(9, r).setFormula(conv(f'=IF($B{R}="","",COUNTIFS(R_Regime,$B{R},R_Ouvert,1))'))
        sh.Rows.getByIndex(r).Height = sh.Rows.getByIndex(hdr + 1).Height
    k.copy(sh, 1, total, 9, total, 1, rt)
    for c, col in enumerate("CDEFGHIJ"):
        sh.getCellByPosition(2 + c, rt).setFormula(conv(f"=SUM({col}{r1 + 1}:{col}{rt})"))
    k.mark("Dashboard", 1, rb, 9, rt)


def main(src, out, out_red):
    matrices = matrices_origine(src)
    k = Calc(src)
    try:
        donnees(k)
        registre(k)
        dashboard(k)
        restaurer_matrices(k, matrices)
        k.save(out)
        for s, l, t, r, b in k.red:
            k.sheet(s).getCellRangeByPosition(l, t, r, b).CellBackColor = 0xFFFF00
        k.save(out_red)
    finally:
        k.close()


if __name__ == "__main__":
    main(*sys.argv[1:4])
