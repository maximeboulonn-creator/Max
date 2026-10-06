# -*- coding: utf-8 -*-
from lib import *
from params import name
import datetime as dt

FUNDS=['MAPIF II','MSIG 3 (USD)','PG NGI','Ares AGI ELTIF']
TERMS=[
 ('Dénomination et forme','Macquarie Alliance Partners Infrastructure Fund II SCSp, Luxembourg, non réglementée, FIA fermé','Macquarie Specialized Infrastructure Global 3 (USD) : LP Delaware et SCSp luxembourgeoises, non réglementées','Partners Group Evergreen SICAV, compartiment Next Generation Infrastructure, SICAV Partie II','Ares Global Infrastructure ELTIF, compartiment d\'AEFS SICAV, ELTIF agréé CSSF'),
 ('Stratégie','Secondaires d\'infrastructure : LP-led 55-75 %, GP-led 25-45 %, directs et primaires ≤ 5 %','Dette d\'infrastructure senior et subordonnée, originée ou acquise, OCDE à 80 % au moins','Infrastructure « next generation » : directs 50-90 %, fonds 10-50 %, liquidités 0-20 %','Infrastructure core, fonds propres et dette, ≥ 70 % d\'actifs éligibles ELTIF, ≤ 30 % liquides'),
 ('Statut à la date d\'arrêté','Non opérationnel, blind pool (PPM avril 2026)','Aucun actif, LPA non signés (PPM février 2026)','En activité depuis 2024, seul fonds cible avec portefeuille','Période de souscription initiale, blind pool, fin attendue le 01/11/2026'),
 ('Devise de référence','USD, catégorie unique','USD ; MSIG 3 (EUR) est un fonds distinct','USD ; classes « H » couvertes sans obligation','EUR'),
 ('Allocation cible','34 %','17 %','17 %','17 %'),
 ('Souscriptions','Closings sur 12 mois (+ 6), appels à 10 j.o. de préavis','Closings sur 2 ans (+ 6 mois), hausses d\'engagement semestrielles','Mensuelles, préavis 10 j.o.','Mensuelles, cut-off 15 jours calendaires'),
 ('Rachats','Aucun ; durée 10 ans + 2 + 2 ; cession agréée par le GP','Annuel au 31/03 après lock-up de 3 ans, préavis 45 jours, payé au run-off (Withdrawing Account), minimum 10 M$','Mensuels, préavis 45 j.o., règlement 4 j.o. après VL','Trimestriels, cut-off le 1er j.o. du dernier mois, règlement 2 j.o. après publication'),
 ('Plafonnement','Sans objet','Sans objet ; conversion forcée possible sous 300 M$ de VL','5 % de la VL par trimestre, distributions et conversions comprises ; Special Dealing à 180 jours, prix discrétionnaire','5 % de la VL par trimestre (net) et limite ELTIF (27,3 % de l\'assiette liquide)'),
 ('Pénalité de sortie','Sans objet','Aucune ; non-appelé survivant au retrait','Redemption Fee discrétionnaire ≤ 5 % (classe I) et Exit Fee ≤ 5 %, sans butoir de durée','Early Redemption Deduction de 5 % avant 24 mois'),
 ('Valeur liquidative','Au moins annuelle ; états trimestriels non audités','Trimestrielle, US GAAP, exercice au 31/03, reporting à 30 j.o.','Mensuelle, publiée 21 j.o. après le Dealing Day','Mensuelle, publiée 15 j.o. après la Subscription Date'),
 ('Frais de gestion','1,25 % des engagements puis de la VL et du non-appelé','1,00 % du capital investi, levier compris','1,25 % de la VL augmentée des engagements non appelés','Jusqu\'à 1,25 % de la VL (classe C)'),
 ('Commission de performance','12,5 %, preferred return 8 %, catch-up 100 %, clawback','12,5 %, catch-up 80 %, hurdle 8 % (levered) ou 6 % (unlevered), look-back 3 ans','15 %, hurdle 5 %, high water mark, catch-up 100 %','12,5 %, hurdle 5 %, high water mark, catch-up 100 %'),
 ('Levier','Emprunts ≤ 30 % des engagements, facilités ≤ 30 % de la VL, filiales non plafonnées ; plafonds AIFM non communiqués','Manche levered : financement ≤ 2,5 fois les apports (≈ 350 % de la VL) ; dette non-recourse exclue','Emprunt ≤ 40 % de la VL ; AIFM 200 % engagement, 410 % brut','Emprunt ELTIF ≤ 50 % à partir de 2029 ; AIFM 400 % engagement, 500 % brut'),
 ('Ticket minimum','10 M$, dérogation à obtenir par écrit','10 M$, dérogation à obtenir par écrit ; élection AWI au Subscription Agreement','1 M USD (classe I)','1 M€ (classe C)'),
 ('TRI net cible','12 à 15 % (synthèse : 12 à 14 %)','10 à 12 % (levered) ou 8 à 10 % (unlevered)','Non chiffré au prospectus (synthèse : 10 à 12 %)','Non chiffré au prospectus (synthèse : 8 à 10 %)'),
 ('SFDR','Article 6','Fonds USD non classé article 8 ; MSIG 3 (EUR) attendu article 8','Article 8','Article 6'),
 ('Prestataires','AIFM Macquarie Asset Management Europe ; dépositaire J.P. Morgan SE ; PwC','AIFM Macquarie Asset Management Credit Advisers US ; administrateur SEI ; aucun dépositaire AIFMD identifié','AIFM Partners Group (Luxembourg) ; dépositaire Northern Trust ; PwC','AIFM Ares Management Luxembourg ; dépositaire BNY Mellon ; EY'),
]
IRP={ # (marché, liquidité, crédit, contrepartie, opérationnel, levier) scores sur max (13,4,2,6,10,3)
 'Fonds':(10,4,2,1,7,0),'MAPIF II':(10,1,0,2,8,3),'MSIG 3 (USD)':(11,3,2,4,8,3),'PG NGI':(10,4,2,5,10,3),'Ares AGI ELTIF':(10,4,2,4,7,3)}
MAX=(13,4,2,6,10,3); W=(0.25,0.30,0.10,0.10,0.15,0.10)
TYPO=['Risque de marché','Risque de liquidité','Risque de crédit','Risque de contrepartie','Risque opérationnel','Risque de levier']

def build_fonds(wb):
    ws=wb.create_sheet('4 Fonds cibles')
    S=Sheet(ws,'FONDS CIBLES - conditions contractuelles et profil de risque initial',
            'Lecture des prospectus et PPM (classeur « Openstone underlying funds risk and liquidity terms », 3 septembre 2026) et Initial Risk Profile du 06/10/2026.',
            {'A':2,'B':26,'C':40,'D':40,'E':40,'F':40})
    S.hdr(['Terme']+FUNDS)
    for t in TERMS: S.row(list(t),h=42)
    S.blank()
    S.block('INITIAL RISK PROFILE - scores par typologie (score / maximum atteignable), pondération 25 / 30 / 10 / 10 / 15 / 10')
    S.hdr(['Typologie','Fonds (propre)']+FUNDS)
    ws.column_dimensions['G'].width=20
    S.ncols=7
    cols=['Fonds']+FUNDS
    f=S.r
    for i,t in enumerate(TYPO):
        S.row([t]+[f'{IRP[c][i]} / {MAX[i]}' for c in cols],h=16)
    # numeric score rows
    r=S.row(['Score pondéré (0 à 100)']+['' for c in cols],h=16)
    for j,c in enumerate(cols):
        expr='+'.join(f'{W[i]}*{IRP[c][i]}/{MAX[i]}' for i in range(6))
        cell=ws.cell(r,3+j,f'=ROUND(100*({expr}),1)'); cell.number_format='0.0'; cell.fill=fill('D9F785'); cell.font=F(True,INK)
    for j,nm in enumerate(['IRPc_Fonds','IRPc_MAPIF','IRPc_MSIG','IRPc_PG','IRPc_ARES']): name(wb,ws,nm,f'{get_column_letter(3+j)}{r}')
    r2=S.row(['Notation']+[f'=IF({get_column_letter(3+j)}{r}>56,"Élevé",IF({get_column_letter(3+j)}{r}>34,"Moyen","Faible"))' for j in range(5)],h=16)
    r3=S.row(['Allocation cible','-','=W_MAPIF','=W_MSIG','=W_PG','=W_ARES'],fmts=[None,None,'0 %','0 %','0 %','0 %'],h=16)
    r4=S.row(['Score en transparence (pondéré par l\'allocation, poche liquide à 0)',f'=ROUND(D{r3}*D{r}+E{r3}*E{r}+F{r3}*F{r}+G{r3}*G{r},1)','','','',''],fmts=[None,'0.0'],key_cols=(1,))
    name(wb,ws,'IRPc_Transp',f'C{r4}')
    r5=S.row(['Score en transparence rebasé hors poche liquide',f'=ROUND(C{r4}/(1-W_LIQ),1)','','','',''],fmts=[None,'0.0'])
    name(wb,ws,'IRPc_Transp_hp',f'C{r5}')
    S.blank()
    S.note('Échelle IRP : Faible 0 à 34, Moyen 34 à 56, Élevé 56 à 100. Les scores sont ceux du classeur Initial Risk Profile du 06/10/2026, recalculés ici à partir des scores par typologie ; les notations détaillées et sources figurent dans ce classeur.')

RECS=[
 ('R1','Recaler la note de liquidité et le classeur de stress tests sur le projet de prospectus du 30/09/2026','Préavis de 91 jours (centralisation 31/03 et 30/09) et non 180 ; Demandes Prioritaires maintenues à l\'art. 7.2.6 ; issue du différé après 4 dates à écrire.','Fonction Risques','Avant le dépôt du prospectus'),
 ('R2','Trancher les onze écarts du registre des limites du prospectus','Emprunt interdit mais supposé aux art. 3.1.1, 14.2 et 25.2.6 ; plancher de 15 % permanent sans clause de dépassement passif ; double plafond de 130 % ; assiette du plafonnement ; souscription au nominal ; durée de placement 10 ans contre 8 ans en synthèse ; barème de conseil.','Société de Gestion, Conseiller, Paul Hastings','Avant le dépôt du prospectus'),
 ('R3','Réintroduire au prospectus le niveau maximal de levier (brut et engagement)','Art. 23 § 1 a) de la directive 2011/61/UE ; le projet du 30/09 l\'a supprimé. Proposition : 100 % en engagement et en brut, trésorerie exclue.','Société de Gestion','Avant le dépôt du prospectus'),
 ('R4','Obtenir par écrit les dérogations de ticket minimum de MAPIF II et de MSIG 3 (10 M$ chacun)','À 50 M€ d\'encours cible, 34 % et 17 % représentent 17 et 8,5 M€ : l\'allocation cible n\'est pas constructible sans dérogation.','Conseiller, puis Comité d\'investissement','Avant tout engagement'),
 ('R5','Arrêter le véhicule et la manche de MSIG 3, et la classe de PG NGI','USD ou EUR ; levered (≈ 350 % de levier, hurdle 8 %) ou unlevered ; classe I ou classe H couverte. Confirmer l\'admission d\'Openstone Infraworld comme Automatic Withdrawing Investor avant signature.','Comité d\'investissement, avis de la Fonction Risques','Avant tout engagement'),
 ('R6','Fixer des limites internes de concentration et de change','Gérant unique 40 % de l\'actif net (alerte 35 %) ; fonds cible 35 % (alerte 30 %) ; dollar non couvert 40 %. L\'allocation cible (Macquarie 51 %, MAPIF II 34 %) devra être justifiée ou revue.','Comité des risques','Prochain comité'),
 ('R7','Séquencer les engagements fermés pour que les appels à douze mois restent couverts','Point de rupture du test inversé : appels de 12,5 % de l\'actif net sur douze mois (non-appelé de 25 %). Rythme d\'engagement et non-appelé plafonné à 15 % de l\'actif net à la fin du blocage.','Comité d\'investissement','À chaque engagement'),
 ('R8','Caler la première fenêtre de rachat sur la fin de la déduction Ares (24 mois)','Un rachat chez Ares avant 24 mois coûte 5 % de la ligne ; la première Date de Centralisation (31/03/2029 pour une constitution fin 2026) doit suivre les 24 mois de détention des parts Ares.','Fonction Risques, gestion de trésorerie','Au lancement'),
 ('R9','Écrire la procédure de revue des valeurs liquidatives reçues','Seuil de péremption des VL des fonds cibles (trimestrielles et décalées pour MAPIF II et MSIG 3), contrôle de cohérence, escalade en cas d\'écart. Calendrier Annexe IV tenant compte de l\'exercice MSIG 3 au 31/03.','Fonction Risques','T4 2026'),
 ('R10','Documenter la gestion du conflit d\'intérêts du Conseiller','Le Conseiller cumule conseil en investissement, distribution principale et faculté de révoquer la Société de Gestion (art. 3.1.2 à 3.1.5) ; accord préalable requis pour toute modification du prospectus.','Dirigeants effectifs, Conformité','T4 2026'),
 ('R11','Compléter le DIC et la synthèse commerciale','SRI 4 à reporter au prospectus et à la synthèse ; durée de détention 10 ans ; commission de rachat « néant » alors que le prospectus mentionne des droits de sortie de 5 % avant 2 ans dans des versions antérieures.','Société de Gestion, Conseiller','Avant la commercialisation'),
 ('R12','Rejouer les stress tests sur l\'encours réel à chaque arrêté trimestriel et avant chaque décision de plafonnement','Hypothèses du classeur à arrêter en Comité : collecte, non-appelé de départ, décote de cession, part des parts de distribution.','Fonction Risques','Dès la première VL'),
]
DEC=[
 ('1','Figer le texte de liquidité du prospectus avant dépôt','Préavis 91 jours, Demandes Prioritaires, issue après 4 dates, assiette du plafonnement, clause de dépassement passif du plancher de 15 %.','R1, R2','Comité des risques puis Société de Gestion et Conseiller','Avant le dépôt'),
 ('2','Réintroduire le levier maximal AIFM (100 % engagement et brut)','Obligation de l\'art. 23 § 1 a) de la directive ; sans emprunt ni dérivé, le niveau est sans coût.','R3','Société de Gestion','Avant le dépôt'),
 ('3','Conditionner tout engagement fermé aux dérogations écrites et aux choix de véhicule','Dérogations de ticket MAPIF II et MSIG 3, confirmation AWI, manche et devise de MSIG 3, classe PG NGI.','R4, R5','Comité d\'investissement','Avant tout engagement'),
 ('4','Arrêter les limites internes KRI-08 à KRI-11','Gérant 40 %, fonds cible 35 %, dollar non couvert 40 %, péremption des VL à 120 jours pour les fonds fermés.','R6, R9','Comité des risques','Prochain comité'),
 ('5','Adopter la règle de séquencement des engagements et de la première fenêtre','Appels à douze mois couverts par la poche ; non-appelé ≤ 15 % de l\'actif net à la fin du blocage ; première centralisation après 24 mois de détention Ares.','R7, R8','Comité d\'investissement, avis de la Fonction Risques','Au lancement'),
 ('6','Formaliser la gouvernance du Conseiller et la revue des VL reçues','Cartographie du conflit conseil / distribution / révocation ; procédure de revue des VL et calendrier Annexe IV.','R9, R10','Dirigeants effectifs','T4 2026'),
]
def build_recos(wb):
    ws=wb.create_sheet('5 Recommandations')
    S=Sheet(ws,'RECOMMANDATIONS DE LA FONCTION RISQUES ET DÉCISIONS SOUMISES AU COMITÉ',
            'Classées par urgence ; chaque recommandation indique l\'instance compétente et le délai proposé.',
            {'A':2,'B':8,'C':48,'D':70,'E':30,'F':24})
    S.hdr(['N°','Recommandation','Motif','Instance','Délai'])
    f=S.r
    for t in RECS: S.row(list(t),h=48)
    for i in range(len(RECS)):
        name(wb,ws,f'Rec{i+1}',f'C{f+i}')
    S.blank()
    S.block('SIX DÉCISIONS SOUMISES AU COMITÉ DES RISQUES')
    S.hdr(['N°','Décision','Proposition de la Fonction Risques','Recommandations','Instance','Échéance'])
    ws.column_dimensions['G'].width=20
    f=S.r
    for t in DEC: S.row(list(t),h=40)
    for i in range(len(DEC)):
        name(wb,ws,f'Dec{i+1}',f'C{f+i}'); name(wb,ws,f'DecP{i+1}',f'D{f+i}'); name(wb,ws,f'DecE{i+1}',f'G{f+i}')
    S.blank(); S.ncols=6
    S.note('Les recommandations R11 et R12 sont suivies par la Fonction Risques sans décision du Comité. Les points entérinés sans débat : échelle de verdict de la charte, convention TTL (positions sans fenêtre garantie au-delà de 365 jours), scénarios de la note de liquidité § 6.')

SOURCES=[
 ('Prospectus et Règlement Openstone Infraworld, projet',dt.date(2026,9,30),30,'Référentiel des limites, conditions de rachat, frais, gouvernance.'),
 ('Synthèse commerciale Openstone Infraworld v5.4',dt.date(2026,9,30),60,'Allocation cible, fonds sélectionnés, frais des fonds cibles, objectifs non garantis. Document de pré-commercialisation.'),
 ('Note de liquidité Openstone Infraworld',dt.date(2026,9,28),90,'Outils de gestion de la liquidité, seuils d\'alerte, programme de simulations. Préavis de 180 jours à recaler.'),
 ('Classeur « Openstone Stress Test liquidité »',dt.date(2026,9,24),90,'Résultats des neuf scénarios et des tests inversés (onglet 3).'),
 ('Registre des limites du prospectus (25/09/2026) et liste des douze écarts',dt.date(2026,10,5),30,'Onglet 1 et recommandation R2.'),
 ('Initial Risk Profile Openstone Infraworld et quatre fonds cibles',dt.date(2026,10,6),90,'Scores IRP (onglet 4).'),
 ('Classeur « Openstone underlying funds risk and liquidity terms »',dt.date(2026,9,14),90,'Conditions des fonds cibles et vingt écarts documentaires (onglet 4).'),
 ('Documents d\'informations clés PRIIPs, parts A1 à B3',dt.date(2026,9,29),60,'SRI 4, période de détention 10 ans, scénarios de performance, coûts.'),
 ('Prospectus AEFS SICAV, annexe 7 Ares Global Infrastructure ELTIF',dt.date(2026,8,11),180,'Termes Ares (visa CSSF du 11/08/2026).'),
 ('Prospectus Partners Group Evergreen SICAV, supplément 2',dt.date(2026,6,9),180,'Termes PG NGI (visa CSSF du 09/06/2026).'),
 ('PPM Macquarie Alliance Partners Infrastructure Fund II',dt.date(2026,4,30),180,'Termes MAPIF II ; LPA du 20/03/2026.'),
 ('PPM Macquarie Specialized Infrastructure Global 3 (USD)',dt.date(2026,2,28),180,'Termes MSIG 3 ; first closing non connu.'),
]
def build_data(wb):
    ws=wb.create_sheet('DATA')
    S=Sheet(ws,'SOURCES ET FRAÎCHEUR DES DONNÉES',
            'Chaque source, sa date, son ancienneté à la date d\'arrêté et le seuil de péremption retenu.',
            {'A':2,'B':62,'C':12,'D':12,'E':10,'F':12,'G':70})
    S.hdr(['Source','Date','Ancienneté (j)','Seuil (j)','Statut','Usage dans le rapport'])
    f=S.r
    for t in SOURCES:
        r=S.r
        S.row([t[0],t[1],f'=DateArrete-C{r}',t[2],f'=IF(D{r}<=E{r},"Conforme","Périmée")',t[3]],fmts=[None,'dd/mm/yyyy','0','0'],h=26,input_cols=(1,3))
    l=S.r-1
    r=S.row(['Sources périmées',f'=COUNTIF(F{f}:F{l},"Périmée")','','','',''],fmts=[None,'0'],key_cols=(1,)); name(wb,ws,'Src_perimees',f'C{r}')
    name(wb,ws,'Src_total',f'C{r+1}')
    S.row(['Sources recensées',f'=COUNTA(B{f}:B{l})','','','',''],fmts=[None,'0'])
    S.blank(); S.ncols=6
    S.note('Cadre : directive 2011/61/UE art. 15 et 16, modifiée par la directive (UE) 2024/927 (annexe V) ; règlement délégué (UE) 231/2013 art. 38 à 49 ; orientations ESMA 34-39-897 (simulations de crise de liquidité) ; instruction AMF DOC-2017-05 (plafonnement des rachats) ; règlement (UE) 1286/2014 PRIIPs ; art. L. 214-154 et suivants du CMF et 423-27 et suivants du RG AMF.')
