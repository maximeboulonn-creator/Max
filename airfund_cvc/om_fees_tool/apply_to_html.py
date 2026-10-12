#!/usr/bin/env python3
"""Intègre le modèle de frais (JSON + module JS + CSS) dans l'Operating Memorandum HTML.

Usage : python3 -I apply_to_html.py <OM_source.html> <fees_model.json> <fees_module.js> <fees_module.css> <OM_sortie.html>
"""
import json
import re
import sys


def main(src, fees_json, mod_js, mod_css, dst):
    html = open(src, encoding='utf-8').read()
    fees = json.load(open(fees_json, encoding='utf-8'))
    js = open(mod_js, encoding='utf-8').read()
    css = open(mod_css, encoding='utf-8').read()
    assert '</script' not in js, 'le module JS ne doit pas contenir </script'

    # 1. bloc de données : DATA.fees et onglet fees-tool
    m = re.search(r'(<script id="data" type="application/json">)(.*?)(</script>)', html, re.S)
    data = json.loads(m.group(2))
    data['fees'] = fees
    for t in data['tabs']:
        if t['id'] == 'fees-tool':
            t['sheet'] = 'Modèle de frais'
            t['desc'] = ("Modèle de frais Fundcraft intégré (classeur v3) : saisie des paramètres et du barème à cinq tranches, "
                         "récap par palier d'Actif Net, détail Dépositaire / Valorisation, provision à la VL et prévisionnel de trésorerie.")
    data['fund']['search']['placeholder'] = data['fund']['search'].get('placeholder', '')
    blob = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    html = html[:m.start(2)] + blob + html[m.end(2):]

    # 2. module JS : du commentaire « Fees tool » jusqu'au commentaire « PRIIPs KIDs »
    start = html.index('  /* =====================================================================\n     Fees tool')
    end = html.index('  /* =====================================================================\n     PRIIPs KIDs')
    html = html[:start] + js + html[end:]

    # 3. CSS : avant la fermeture de la feuille de style principale
    i = html.index('</style>')
    html = html[:i] + css + html[i:]

    # 4. onglet Sources : le classeur de référence
    old_row = "<tr><td>Classeur de frais — stack à trois couches (FoF, CVC-PE, CVC-PESEC)<\\/td><td>AirFund CVC fees.xlsx — Dépositaire 15 000 €, Délégataire Comptable 15 000 €, Commissaire aux Comptes 5 500 €<\\/td><td>Lignes par défaut de l\\'onglet <a href=\"#tab=fees-tool\">frais<\\/a> (frais directs du Fonds uniquement) ; les diligences KYC et le plafond de commissions indirectes sont fournis en lignes désactivées.<\\/td><\\/tr>"
    new_row = ("<tr><td>Modèle de frais Fundcraft - classeur v3 (01_Input, 02_Classes_de_Parts, 03_Detail, 04_Recap, 05_Tresorerie, 06_Documentation)<\\/td>"
               "<td>Fundcraft_AirFund_Modele_Frais_3.xlsx - Convention Dépositaire BFCM (Annexe 11, 01/06/2026), annexe tarifaire CIC du 24/06/2026, lettre de mission Aplitec du 01/07/2026, décisions des 11 et 14/09/2026<\\/td>"
               "<td>Intégré tel quel à l\\'onglet <a href=\"#tab=fees-tool\">Modèle de frais<\\/a> : mêmes paramètres, même barème et mêmes formules (récap par palier, détail Dépositaire / Valorisation, prévisionnel de trésorerie) ; toutes les saisies sont modifiables et conservées dans le fichier.<\\/td><\\/tr>")
    assert old_row in html, 'ligne Sources introuvable'
    html = html.replace(old_row, new_row)

    open(dst, 'w', encoding='utf-8').write(html)
    print('écrit', dst, len(html), 'octets')


if __name__ == '__main__':
    main(*sys.argv[1:6])
