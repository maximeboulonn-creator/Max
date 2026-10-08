# -*- coding: utf-8 -*-
"""Version HTML du registre des incidents, produite à partir de l'export XML validé.

Chaîne : classeur -> export_xml.py -> XML (validé contre registre_incidents.xsd) -> export_html.py -> HTML autonome.
La page est en consultation seule ; les données sont embarquées (aucun appel réseau hormis les polices).
Usage : python export_html.py registre.xml sortie.html
"""
import json
import os
import sys

from lxml import etree

ICI = os.path.dirname(os.path.abspath(__file__))
XSD = os.path.join(ICI, "registre_incidents.xsd")
GABARIT = os.path.join(ICI, "registre_template.html")

# Éléments répétables, par (parent, enfant) : toujours rendus sous forme de liste.
LISTES = {("Vehicules", "Vehicule"), ("SousCategories", "SousCategorie"), ("OriginesDetection", "Origine"),
          ("CausesRacines", "Cause"), ("NiveauxEscalade", "Niveau"), ("Notifications", "Notification"),
          ("BaremeGravite", "Niveau"), ("Parametres", "Parametre"), ("SeuilsIndicateurs", "Seuil"),
          ("Incidents", "Incident"), ("Indicateurs", "Section"), ("Section", "Indicateur"),
          ("ParVehicule", "Ligne"), ("ParRegimeLiquidite", "Ligne"), ("EvolutionTrimestrielle", "Ligne"),
          ("Ligne", "Categorie"), ("Total", "Categorie"), ("IncidentsSignales", "IncidentSignale")}


def nom(tag):
    return etree.QName(tag).localname


def convertir(e):
    tag, enfants = nom(e.tag), [k for k in e if isinstance(k.tag, str)]
    if not enfants:
        if not e.attrib:
            return e.text
        return {**e.attrib, **({"valeur": e.text} if e.text is not None else {})}
    d = dict(e.attrib)
    for k in enfants:
        t = nom(k.tag)
        v = convertir(k)
        if (tag, t) in LISTES:
            d.setdefault(t, []).append(v)
        else:
            d[t] = v
    for p, c in LISTES:
        if p == tag:
            d.setdefault(c, [])
    return d


def exporter(xml, out):
    doc = etree.parse(xml)
    schema = etree.XMLSchema(etree.parse(XSD))
    if not schema.validate(doc):
        raise SystemExit("XML non conforme au schéma : " + str(schema.error_log.last_error))
    data = convertir(doc.getroot())
    js = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    html = open(GABARIT, encoding="utf-8").read().replace("__DATA__", js)
    open(out, "w", encoding="utf-8").write(html)
    return len(data["Incidents"].get("Incident", []))


if __name__ == "__main__":
    print(f"{exporter(sys.argv[1], sys.argv[2])} incidents mis en page")
