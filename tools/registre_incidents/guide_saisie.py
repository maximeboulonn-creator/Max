# -*- coding: utf-8 -*-
"""Ajoute le guide de saisie (commentaires d'en-tête) au registre France et produit une version vierge.

- Version de travail : chaque en-tête de colonne à saisir porte un commentaire (moment, caractère
  obligatoire, format, exemple), visible au survol dans Excel.
- Version redline : en-têtes commentés surlignés en jaune.
- Version vierge : même classeur, saisies du registre effacées (formules, référentiels et paramètres conservés).
Usage : python guide_saisie.py source.xlsx sortie.xlsx sortie_redline.xlsx sortie_vierge.xlsx
"""
import os
import sys
import zipfile

import uno

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from complete_france import Calc, find_col, fr, FIRST, LAST, matrices_origine, restaurer_matrices  # noqa: E402

D, T, C = "À la déclaration", "Pendant le traitement", "À la clôture"
GUIDE = {
    "Réf.": f"{D} - obligatoire. Identifiant unique INC-AAAA-NNN, numéroté à la suite par année (INC-2026-001, INC-2026-002...).",
    "Fonds ou véhicule": f"{D} - obligatoire. Liste. Véhicule concerné, ou Fundcraft France SAS pour un incident propre à la SGP. "
                         "Le type et le régime de liquidité se remplissent seuls.",
    "Sous-catégorie": f"{D} - obligatoire. Liste. La catégorie se calcule seule.",
    "Description": f"{D} - obligatoire. Une phrase factuelle, sans nom d'investisseur. "
                   "Exemple : Avis d'appel de fonds envoyé avec un montant erroné.",
    "Date de survenance": f"{D} - recommandé. Date de l'événement générateur, estimée si besoin. Format jj/mm/aaaa.",
    "Détection (date et heure)": f"{D} - obligatoire. Prise de connaissance par la SGP ou son délégataire. "
                                 "Format jj/mm/aaaa hh:mm. Cette date déclenche tous les calculs de la ligne.",
    "Déclaration au registre": f"{D} - obligatoire. Date d'inscription au registre. Format jj/mm/aaaa. Délai cible dans l'onglet Données.",
    "Prise en compte (date et heure)": f"{T}. Affectation d'un responsable et lancement des mesures conservatoires. Format jj/mm/aaaa hh:mm.",
    "Origine de la détection": f"{D} - obligatoire. Liste. Détermine le taux d'auto-détection.",
    "Cause racine": f"{T}, dès qu'elle est connue. Liste. Attendue avant la clôture des incidents significatifs et majeurs.",
    "Impact financier brut (EUR)": f"{T}. Coût total estimé : correction, indemnisation, frais. Entre dans le calcul de la gravité. Exemple : 12500.",
    "Recouvrement (EUR)": f"{T}. Montants récupérés auprès d'un délégataire, d'un assureur ou d'une contrepartie.",
    "Écart de VL (bps)": f"{T}, si erreur de VL. Écart en points de base. Entre dans le calcul de la gravité.",
    "Investisseurs lésés": f"{T}. Nombre d'investisseurs affectés. Entre dans le calcul de la gravité.",
    "Indemnisation investisseurs (EUR)": f"{T}. Montant versé ou à verser aux investisseurs.",
    "Impact réglementaire": f"{T}. Liste. Entre dans le calcul de la gravité.",
    "Impact réputationnel": f"{T}. Liste. Entre dans le calcul de la gravité.",
    "Gravité forcée": "Facultatif. Uniquement pour imposer une gravité différente du barème ; en donner la raison dans la description.",
    "Dépassement actif ou passif": f"{T}, si dépassement de limite. Liste Actif ou Passif ; fixe le délai de régularisation.",
    "Notification à une autorité": f"{T}, si une notification est requise. Liste (DORA, RGPD - CNIL, AMF, Tracfin...). L'échéance se calcule seule.",
    "Notification réalisée (date et heure)": f"{T}. Date et heure d'envoi. Pour DORA, vaut notification initiale. Format jj/mm/aaaa hh:mm.",
    "Classification DORA (date et heure)": f"{T}, si incident TIC majeur. Décision de classement en incident majeur. Format jj/mm/aaaa hh:mm.",
    "Rapport intermédiaire DORA (date et heure)": f"{T}, si incident TIC majeur. Au plus 72 h après la notification initiale. Format jj/mm/aaaa hh:mm.",
    "Rapport final DORA (date)": f"{C}, si incident TIC majeur. Au plus 1 mois après le rapport intermédiaire. Format jj/mm/aaaa.",
    "Accusé de réception (réclamation)": f"{T}, si réclamation investisseur. Date de l'accusé de réception (cible 10 j.o.). Format jj/mm/aaaa.",
    "Niveau d'escalade atteint": f"{T}. Liste. Le niveau requis par la gravité se calcule seul ; l'écart apparaît dans « Escalade conforme ».",
    "Date d'escalade": f"{T}. Date de saisine du niveau atteint. Exige un niveau atteint. Format jj/mm/aaaa.",
    "Nombre de réaffectations": f"{T}. Nombre de changements d'équipe responsable ; saisir 0 si aucun.",
    "Statut": "Obligatoire. Ouvert à la déclaration, En remédiation pendant le traitement, Clôturé à la clôture (exige une date de clôture).",
    "Résolution (date et heure)": f"{T}. Impact contenu : VL corrigée, opération régularisée, service rétabli. "
                                  "Distincte de la clôture. Format jj/mm/aaaa hh:mm.",
    "Date de clôture": f"{C}. Clôture validée par le Risk Management. Exige le statut Clôturé. Format jj/mm/aaaa.",
    "Réouvert": f"{C}. Oui si l'incident a été rouvert après une première résolution.",
    "Action corrective": f"{T}. Description de l'action décidée.",
    "Responsable": f"{T}. Personne ou équipe en charge de l'action.",
    "Échéance de l'action": f"{T}. Date cible. Une action non réalisée après l'échéance compte en retard. Format jj/mm/aaaa.",
    "Action réalisée le": f"{C}. Date de réalisation constatée. Format jj/mm/aaaa.",
}


def commenter(k):
    sh = k.sheet("Registre")
    ann = sh.getAnnotations()
    cols = []
    for lab, texte in GUIDE.items():
        c = find_col(sh, lab)
        ann.insertNew(k.addr(sh, c, 5), fr(texte))
        forme = sh.getCellByPosition(c, 5).getAnnotation().getAnnotationShape()
        forme.CharFontName, forme.CharHeight = "Calibri", 9
        taille = uno.createUnoStruct("com.sun.star.awt.Size")
        taille.Width, taille.Height = 7000, 450 * (len(texte) // 48 + 2)  # 1/100 mm
        forme.setSize(taille)
        k.mark("Registre", c, 5, c, 5)
        cols.append(c)
    return cols


def renommer_auteur(path, auteur="Risk Management"):
    """LibreOffice exporte les commentaires sous l'auteur « Unknown Author » : on le remplace dans le fichier."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename.startswith("xl/comments"):
                data = data.replace(b"<author>Unknown Author</author>", f"<author>{auteur}</author>".encode())
            zo.writestr(it, data)
    os.replace(tmp, path)


def vider(k, cols):
    sh = k.sheet("Registre")
    for c in cols:
        sh.getCellRangeByPosition(c, FIRST - 1, c, LAST - 1).clearContents(1 | 2 | 4)


def main(src, out, out_red, out_vierge):
    matrices = matrices_origine(src)
    k = Calc(src)
    try:
        cols = commenter(k)
        restaurer_matrices(k, matrices)
        k.save(out)
        for s, l, t, r, b in k.red:
            k.sheet(s).getCellRangeByPosition(l, t, r, b).CellBackColor = 0xFFFF00
        k.save(out_red)
        for s, l, t, r, b in k.red:
            k.sheet(s).getCellRangeByPosition(l, t, r, b).CellBackColor = 0x218E8E
        vider(k, cols)
        restaurer_matrices(k, matrices)
        k.save(out_vierge)
    finally:
        k.close()
    for p in (out, out_red, out_vierge):
        renommer_auteur(p)


if __name__ == "__main__":
    main(*sys.argv[1:5])
