# -*- coding: utf-8 -*-
from lib import *
from openpyxl.workbook.defined_name import DefinedName
import datetime as dt

def name(wb, ws, nm, coord):
    wb.defined_names[nm]=DefinedName(nm, attr_text=f"'{ws.title}'!${coord[0]}${coord[1:]}")

def build_params(wb):
    ws=wb.create_sheet('0 Paramètres')
    S=Sheet(ws,'PARAMÈTRES - OPENSTONE INFRAWORLD, COMPARTIMENT I · rapport de risque initial, octobre 2026',
            'Toutes les valeurs du rapport découlent de ce bloc ; aucune autre feuille ne redéfinit un paramètre. Saisies sur fond crème, valeurs dérivées en formules.',
            {'A':2,'B':44,'C':14,'D':10,'E':52,'F':11,'G':60})
    S.hdr(['Paramètre','Valeur','Unité','Source','Fiabilité','Réserve']); hdrcols=6
    S.ncols=6
    P={}
    def row(key,label,val,unit,src,fiab,res='',fmt=None,derived=False):
        r=S.row([label,val,unit,src,fiab,res],fmts=[None,fmt],input_cols=() if derived else (1,))
        P[key]=f"'0 Paramètres'!$C${r}"; name(wb,ws,key,f'C{r}'); return r
    S.block('FONDS ET DATES')
    row('DateArrete','Date d\'arrêté du rapport',dt.date(2026,10,6),'date','Date de l\'analyse ; prospectus projet du 30 septembre 2026','Élevée','Le Fonds n\'est pas constitué : rapport initial sur documentation, sans valeur liquidative ni passif.','dd/mm/yyyy')
    row('DateProspectus','Date du projet de prospectus',dt.date(2026,9,30),'date','Prospectus Openstone Infraworld, projet du 30 septembre 2026','Élevée','Plusieurs mentions restent à compléter (SRI, ISIN A3c, mentions de liquidité).','dd/mm/yyyy')
    row('AN_cible','Actif net cible',50000000,'euros','Encours cible retenu par la note de liquidité et l\'IRP','Moyenne','Prévision commerciale ; les simulations sont rejouées sur l\'encours réel une fois le Fonds lancé.','# ##0')
    row('AN_min','Actif net minimal réglementaire',300000,'euros','Règlement art. 2 ; dissolution si l\'actif reste en deçà 30 jours','Élevée','','# ##0')
    row('Seuil_Declenchement','Seuil de la Date de Déclenchement',10000000,'euros','Prospectus, définitions du Compartiment I : 10 M€ d\'engagements cumulés ou 2e anniversaire','Élevée','Bascule des parts B (conseil 0,57 %) vers le barème des parts A (1,07 %).','# ##0')
    row('Nominal','Valeur nominale de la part',100,'euros','Prospectus art. 26.3 - souscription au nominal tant que la VL lui est inférieure','Élevée','Souscription au-dessus de la VL si celle-ci passe sous 100 € : point de traitement équitable (registre, n° 9).','# ##0.00')
    S.block('ALLOCATION CIBLE  ·  synthèse Openstone de septembre 2026, non contractuelle')
    row('W_MAPIF','Macquarie Alliance Partners Infrastructure Fund II (MAPIF II)',0.34,'% AN','Synthèse Openstone septembre 2026, p. 2','Moyenne','Fonds fermé de secondaires, USD ; conditionné à l\'acceptation de la souscription et à la dérogation de ticket (10 M$).','0 %')
    row('W_MSIG','Macquarie Specialized Infrastructure Global 3 (MSIG 3)',0.17,'% AN','Synthèse Openstone septembre 2026, p. 2','Moyenne','Fonds de dette fermé hybride ; véhicule (USD / EUR) et manche (levered / unlevered) non arrêtés.','0 %')
    row('W_PG','Partners Group Next Generation Infrastructure (PG NGI)',0.17,'% AN','Synthèse Openstone septembre 2026, p. 2','Moyenne','Evergreen, devise de référence USD ; classe I ou classe H couverte à arrêter.','0 %')
    row('W_ARES','Ares Global Infrastructure ELTIF (Ares AGI)',0.17,'% AN','Synthèse Openstone septembre 2026, p. 2','Moyenne','Evergreen ELTIF en EUR ; souscription au nominal jusqu\'au 01/11/2026 attendu.','0 %')
    row('W_LIQ','Actifs liquides',0.15,'% AN','Synthèse septembre 2026 ; prospectus art. 21.8 (minimum 15 % de l\'actif)','Élevée','Plancher écrit comme obligation permanente, sans clause de dépassement passif (registre, n° 2).','0 %')
    r=S.row(['Total de l\'allocation','=W_MAPIF+W_MSIG+W_PG+W_ARES+W_LIQ','% AN','dérivé','-','Doit boucler à 100 %.'],fmts=[None,'0 %']); P['W_TOT']=f"'0 Paramètres'!$C${r}"; name(wb,ws,'W_TOT',f'C{r}')
    r=S.row(['Part des fonds fermés (MAPIF II + MSIG 3)','=W_MAPIF+W_MSIG','% AN','dérivé','-','Aucune liquidité avant échéance ; sortie par cession secondaire décotée.'],fmts=[None,'0 %']); P['W_FERMES']=f"'0 Paramètres'!$C${r}"; name(wb,ws,'W_FERMES',f'C{r}')
    r=S.row(['Part des fonds evergreen (PG NGI + Ares AGI)','=W_PG+W_ARES','% AN','dérivé','-','Rachat trimestriel plafonné à 5 % du fonds cible.'],fmts=[None,'0 %']); P['W_EVER']=f"'0 Paramètres'!$C${r}"; name(wb,ws,'W_EVER',f'C{r}')
    r=S.row(['Part gérée par Macquarie','=W_MAPIF+W_MSIG','% AN','dérivé','-','Concentration gérant : 51 % de l\'allocation cible chez un seul groupe.'],fmts=[None,'0 %']); P['W_MACQ']=f"'0 Paramètres'!$C${r}"; name(wb,ws,'W_MACQ',f'C{r}')
    r=S.row(['Part libellée en dollars - hypothèse basse (MAPIF II)','=W_MAPIF','% AN','dérivé','-','Si MSIG 3 est souscrit via le fonds EUR et PG NGI en classe couverte.'],fmts=[None,'0 %']); name(wb,ws,'W_USD_BAS',f'C{r}')
    r=S.row(['Part libellée en dollars - hypothèse haute (MAPIF II + MSIG 3 USD + PG NGI)','=W_MAPIF+W_MSIG+W_PG','% AN','dérivé','-','Devise de référence de PG NGI : USD (supplément 2) ; MSIG 3 (USD) sans classe EUR.'],fmts=[None,'0 %']); name(wb,ws,'W_USD_HAUT',f'C{r}')
    S.block('CONDITIONS DE LIQUIDITÉ DU FONDS  ·  prospectus du 30 septembre 2026')
    row('Blocage_mois','Période de Blocage par souscription',24,'mois','Prospectus art. 7.2.2 - 2 ans à compter de chaque souscription, imputation FIFO','Élevée','')
    row('Preavis_j','Préavis de rachat',91,'jours','Prospectus art. 7.2.3 - centralisation le 31 mars pour la VL du 30 juin, le 30 septembre pour celle du 31 décembre','Élevée','La note de liquidité du 24/09 et le classeur de stress tests retiennent 180 jours : à recaler.')
    row('Preavis_max_j','Préavis prolongé au plus',180,'jours','Prospectus art. 7.2.7 - 180 jours calendaires au plus, 12 mois au plus','Élevée','')
    row('Gate','Plafonnement des Rachats par date de centralisation',0.05,'% AN','Prospectus art. 7.2.6 - 5 % de l\'Actif Net, demandes nettes des souscriptions','Élevée','Assiette : actif net de la VL d\'exécution, inconnue à l\'activation (registre, n° 7).','0,0 %')
    row('Gate_dates','Reports successifs au plus',4,'dates','Prospectus art. 7.2.6 - 4 Dates de Centralisation ; Demandes Prioritaires après 2 reports','Élevée','L\'issue au-delà de 4 dates n\'est pas écrite ; la note de liquidité prévoit l\'examen de la suspension.')
    row('Reglement_JO','Délai de règlement des rachats',35,'j.o.','Prospectus art. 7.2.5 - 5 Jours Ouvrés après publication, publication sous 30 Jours Ouvrés','Élevée','')
    row('Centralisation_sous_JO','Centralisation des souscriptions avant la VL',25,'j.o.','Prospectus art. 7.1.2','Élevée','Paiement intégral à la centralisation ; sommes bloquées hors actif net.')
    row('Publication_JO','Publication de la valeur liquidative',30,'j.o.','Prospectus art. 6.2','Élevée','VL mensuelle au dernier Jour Ouvré du mois.')
    row('Liq_min','Plancher d\'actifs liquides',0.15,'% actif','Prospectus art. 21.8','Élevée','Assiette en actif, les autres limites en actif net (registre, n° 8).','0 %')
    row('Liq_cible','Poche liquide visée par la gestion',0.16,'% AN','Note de liquidité du 24/09/2026, § 1','Moyenne','','0 %')
    row('Liq_monetaire_min','Dont monétaire et dépôts au minimum (règle interne)',0.05,'% actif','Note de liquidité, § 2','Moyenne','Règle interne, non opposable.','0 %')
    row('Surengagement_max','Plafond de Sur-Engagement',1.30,'% AN','Prospectus art. 23 ; art. 3.1.1 (engagements hors bilan, 130 %)','Élevée','Deux plafonds à 130 % sans précision de cumul (registre, n° 6).','0 %')
    row('Emprunt_max','Emprunts d\'espèces',0,'% AN','Prospectus art. 22 - néant ; ni ligne de crédit ni instrument dérivé','Élevée','Les art. 3.1.1, 14.2 et 25.2.6 supposent encore un emprunt (registre, n° 1).','0 %')
    row('Levier_max','Levier AIFM, méthode de l\'engagement',1.00,'% AN','Note de liquidité § 4 ; à réintroduire au prospectus (art. 23 § 1 a) de la directive 2011/61/UE)','Moyenne','Le projet du 30/09 ne mentionne plus de niveau maximal de levier.','0 %')
    row('Entreprises_max','Entreprises Cibles (co-investissements directs)',0.30,'% AN','Prospectus art. 22','Élevée','Traitement d\'un dépassement passif non précisé.','0 %')
    row('Pers_phys_max','Détention maximale d\'une personne physique',0.10,'% parts','Prospectus art. 5.1 et 8.1','Élevée','Contrôle à la souscription et au transfert.','0 %')
    S.block('FRAIS  ·  prospectus art. 27')
    row('Fee_SGP','Commission de Gestion (tranche 0 - 100 M€)',0.0018,'% AN','Prospectus art. 27.2.1 - 0,18 % ; 0,15 % de 100 à 500 M€ ; 0,10 % au-delà ; minimum 60 000 € par an','Élevée','Le minimum de 60 000 € pèse 0,12 % à 50 M€ et 1,2 % à 5 M€.','0,00 %')
    row('Fee_CIF_A','Commission de Conseil, parts A (tranche 0 - 100 M€)',0.0107,'% AN','Prospectus art. 27.2.2 - 1,07 % ; 1,10 % ; 1,15 %','Élevée','Barème qualifié de dégressif alors que les taux croissent (registre, n° 11).','0,00 %')
    row('Fee_CIF_B','Commission de Conseil, parts B jusqu\'à la Date de Déclenchement',0.0057,'% AN','Prospectus art. 27.2.2','Élevée','','0,00 %')
    row('Fee_Distrib','Commission de Distribution, parts A1 et B1',0.0075,'% AN','Prospectus art. 27.2.3','Élevée','Rétrocédée au Conseiller pour les Distributeurs.','0,00 %')
    row('Fee_Entree','Droits d\'entrée au plus',0.05,'% prix','Prospectus art. 27.2.4','Élevée','Non acquis au Fonds.','0 %')
    row('Fee_Const','Frais de constitution au plus',0.0125,'% AN','Prospectus art. 27.1 - le plus élevé de 60 000 € ou 1,25 % de l\'actif net au 1er anniversaire, amortis sur 5 exercices','Élevée','','0,00 %')
    row('Fee_Total_A1','Frais récurrents directs, part A1 (hors frais fixes)','=Fee_SGP+Fee_CIF_A+Fee_Distrib','% AN','dérivé','-','Hors dépositaire, comptable, CAC et frais des fonds cibles (1,00 à 1,25 % plus carried 12,5 à 15 %).','0,00 %',derived=True)
    S.block('SEUILS D\'ALERTE DE LA NOTE DE LIQUIDITÉ  ·  § 5, vert / orange / rouge')
    row('S_file_rouge','File d\'attente (% AN) - rouge au-delà de',0.25,'% AN','Note de liquidité § 5','Moyenne','Orange entre 0 et 25 %.','0 %')
    row('S_resorption_rouge','Délai de résorption de la file - rouge au-delà de',25,'mois','Note de liquidité § 5','Moyenne','Orange de 12 à 25 mois.')
    row('S_poche_rouge','Poche d\'actifs liquides - rouge en deçà de',0.05,'% actif','Note de liquidité § 5','Moyenne','Orange de 5 à 15 %.','0 %')
    row('S_couv_vert','Couverture de liquidité à 12 mois - vert au-delà de',1.5,'ratio','Note de liquidité § 5','Moyenne','Orange de 1,0 à 1,5 ; rouge en deçà de 1,0.','0.0')
    row('S_couv_rouge','Couverture de liquidité à 12 mois - rouge en deçà de',1.0,'ratio','Note de liquidité § 5','Moyenne','','0.0')
    row('S_sureng_orange','Sur-engagement - orange au-delà de',1.20,'% AN','Note de liquidité § 5','Moyenne','Rouge au-delà de 130 %.','0 %')
    row('S_top1_orange','Premier porteur - orange au-delà de',0.10,'% AN','Note de liquidité § 5','Moyenne','Rouge au-delà de 15 %.','0 %')
    row('S_top5_orange','Cinq premiers porteurs - orange au-delà de',0.25,'% AN','Note de liquidité § 5','Moyenne','Rouge au-delà de 40 %.','0 %')
    S.block('INITIAL RISK PROFILE  ·  classeur du 06/10/2026, scores 0 à 100')
    row('IRP_Fonds','Score IRP du Fonds, caractéristiques propres',71.4,'/100','Initial Risk Profile Openstone Infraworld, onglet Openstone_Infraworld','Élevée','ÉLEVÉ ; liquidité, crédit saturés ; contrepartie et levier faibles.','0.0')
    row('IRP_Transp','Score IRP en transparence, pondéré par l\'allocation cible',62.1,'/100','Idem, bloc Transparence','Élevée','73,1 hors poche liquide.','0.0')
    row('IRP_MAPIF','Score IRP MAPIF II',52.1,'/100','Idem, onglet MAPIF_II','Élevée','MOYEN - fonds fermé sans risque de liquidité propre ; levier et opérationnel élevés.','0.0')
    row('IRP_MSIG','Score IRP MSIG 3 (USD, manche levered)',82.3,'/100','Idem, onglet MSIG_3','Élevée','ÉLEVÉ - crédit et levier saturés ; fonds non lancé à la date du PPM.','0.0')
    row('IRP_PG','Score IRP PG Next Generation Infrastructure',92.6,'/100','Idem, onglet PG_NGI','Élevée','ÉLEVÉ - liquidité, opérationnel, levier saturés ; seul fonds cible en activité.','0.0')
    row('IRP_ARES','Score IRP Ares Global Infrastructure ELTIF',86.4,'/100','Idem, onglet Ares_AGI','Élevée','ÉLEVÉ - blind pool en montée en charge, levier AIFM 400 % / 500 %.','0.0')
    row('SRI','Indicateur synthétique de risque (DIC PRIIPs)',4,'/7','DIC du 29/09/2026, parts A1 à B3 - classe 4, relèvement prudentiel d\'une classe','Élevée','Le prospectus du 30/09 et la synthèse laissent le SRI à compléter (X/7).')
    row('PDR','Période de détention recommandée',10,'ans','Prospectus art. 4.2 (supérieure à 10 ans) ; DIC du 29/09 (10 ans)','Élevée','La synthèse Openstone affiche 8 ans : à aligner (registre, n° 10).')
    S.blank()
    S.note('Fiabilité : Élevée = texte contractuel ou classeur validé ; Moyenne = prévision, hypothèse interne ou document commercial. Les montants en euros sont des cibles : le Fonds n\'a ni valeur liquidative ni porteur à la date d\'arrêté.',6)
    return P
