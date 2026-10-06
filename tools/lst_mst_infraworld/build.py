# -*- coding: utf-8 -*-
"""LST + MST Openstone Infraworld, Compartiment I. Tout paramètre vit dans l'onglet Paramètres (cellules cyan) ; le reste est en formules."""
import datetime as dt,math,statistics
from style import *
from series import monthly_returns
bk=Book(); wb=bk.wb
T=8                      # dates de rachat semestrielles projetées
PC=[L(3+t) for t in range(T)]   # colonnes C..J des périodes
SCEN=[('S1','Base','Rachats courants, collecte normale, marchés stables'),
      ('S2','Plausible','Correction de marché à 1,5 écart-type, rachats de 7,5 % sur deux dates, collecte divisée par deux'),
      ('S3','Sévère','Réplique 2008 : pire semestre historique des proxies, rachats de 15 % sur deux dates, collecte nulle, sortie du premier porteur'),
      ('S4','Extrême','Queue de distribution à 3,5 écarts-types, rachats de 25 % puis 20 %, gel des distributions, pairs saturant les fonds evergreen')]
LINES=[('MAPIF II','Fermé','C1'),('MSIG 3','Fermé','C3'),('PG NGI','Ouvert','C1'),('Ares AGI','Ouvert','ARES'),('Fonds HY UCITS','Liquide','C4'),('Trésorerie','Liquide',None)]
NL=len(LINES)
# ======================================================================= PARAMÈTRES
P=Sheet(bk,'Paramètres','PARAMÈTRES - ANCRES DU PROSPECTUS, ALLOCATION, HYPOTHÈSES ET CALIBRATION',
        'Source unique de tous les calculs. Cellules cyan : saisies modifiables. Les autres onglets ne contiennent que des formules. Les ancres citent l\'article du prospectus du 30/09/2026 (colonne Source).',
        [2,46,14,12,12,12,12,12,12,12,12,46],ncols=11)
ws=P.ws
def prow(label,val,unit,src,nm=None,fmt=None,inp=True,note=''):
    r=P.row([label,val,unit,src,None,None,None,None,None,None,note],fmts=[None,fmt],inputs=(1,) if inp else (),merge=[(5,11)] if False else None)
    if nm: P.nm(nm,r)
    return r
P.sec('A. IDENTIFICATION ET DATES')
P.hdr(['Paramètre','Valeur','Unité','Source','','','','','','','Commentaire'])
prow('Fonds','Openstone Infraworld, Compartiment I','-','Prospectus, p. 1',inp=False)
prow('Société de Gestion','Fundcraft France SAS, GP-20260008','-','Prospectus, p. 1',inp=False)
prow('Structure','Fonds de fonds evergreen investissant en direct dans deux fonds fermés, deux fonds evergreen, un fonds obligataire HY quotidien et de la trésorerie','-','Prospectus art. 3',inp=False)
prow('Date d\'arrêté',dt.date(2026,10,6),'date','-','DateArrete',DATE)
prow('Date de première VL (lancement)',dt.date(2026,12,31),'date','Hypothèse : constitution au T4 2026','Date_lancement',DATE,note='Pilote la fin du blocage et la phase du Fonds.')
prow('Fin de la Période de Blocage = première date de rachat','=EDATE(Date_lancement,Blocage_mois)','date','Prospectus art. 7.2.2 : deux ans à compter de la souscription','Date_depart',DATE,inp=False)
prow('Actif net au départ de la projection',50000000,'EUR','Objectif de collecte de la synthèse (50 M€)','AN_depart',EUR)
prow('Phase du Fonds','=IF(YEARFRAC(Date_lancement,Date_depart)<3,"Montée en charge","Régime de croisière")','-','Dérivé','Phase',inp=False,note='En montée en charge, la concentration des porteurs domine le risque de plafonnement.')
prow('Années depuis le lancement à la date de départ','=YEARFRAC(Date_lancement,Date_depart)','ans','Dérivé','Annees_lancement','0.0',inp=False)
P.blank()
P.sec('B. ANCRES DU PROSPECTUS   (aucune valeur de passif codée ailleurs)')
P.hdr(['Ancre','Valeur','Unité','Source','','','','','','','Commentaire'])
prow('Fréquence de la VL','Mensuelle','-','Prospectus art. 6 : dernier jour calendaire de chaque mois','NAV_Frequence')
prow('Fréquence des souscriptions','Mensuelle, centralisation 25 JO avant la VL','-','Prospectus art. 7.1.2','Sous_Frequence')
prow('Fréquence des rachats','Semestrielle : VL du 30/06 et du 31/12','-','Prospectus art. 7.2.1','Rachat_Frequence')
prow('Dates de rachat par an',2,'dates','Prospectus art. 7.2.1','Rachats_par_an',D)
prow('Préavis : centralisation le 31/03 ou le 30/09',91,'jours cal.','Prospectus art. 7.2.3','Preavis_j',D,note='Jours entre la Date de Centralisation et la VL de rachat.')
prow('Publication de la VL après la date de VL',30,'JO','Prospectus art. 6 et 7.1.5','Publication_JO',D)
prow('Règlement après publication de la VL',5,'JO','Prospectus art. 7.2.4 : cinq Jours Ouvrés au plus','Reglement_JO',D)
prow('Délai VL à cash (publication + règlement)','=ROUND((Publication_JO+Reglement_JO)*7/5,0)','jours cal.','Dérivé','Reglement_cal',D,inp=False)
prow('Plafond de Rachats (gate) par date',0.05,'% AN','Prospectus art. 7.2.6 : 5 % de l\'Actif Net à la VL de rachat','Gate',PCT,note='Assiette : demandes centralisées nettes des souscriptions versées.')
prow('Base du plafond','AN à la VL du 30/06 ou du 31/12, demandes nettes des souscriptions','-','Prospectus art. 7.2.6','Gate_base')
prow('Reports successifs avant examen d\'une suspension',4,'dates','Prospectus art. 7.2.6 ; note de liquidité § 5','Reports_max',D,note='Les demandes reportées perdent la priorité ; Demandes Prioritaires après deux reports.')
prow('Période de Blocage',24,'mois','Prospectus art. 7.2.2','Blocage_mois',D)
prow('Déduction de sortie anticipée',0,'%','DIC du 29/09/2026 : aucune commission de rachat','Deduction_sortie',PCT)
prow('Prolongation du Délai de Préavis : préavis maximal',180,'jours cal.','Prospectus art. 7.2.7','Preavis_max_j',D)
prow('Prolongation du Délai de Préavis : durée maximale',12,'mois','Prospectus art. 7.2.7','Prolong_max_mois',D)
prow('Suspension des Rachats','Oui, circonstances exceptionnelles, sans plafonnement préalable','-','Prospectus art. 7.2.8','Suspension_pouvoir')
prow('Actifs liquides minimum',0.15,'% AN','Prospectus art. 3 et DIC : 15 % au moins en actifs liquides','Liq_min',PCT)
prow('Sur-engagement maximal',1.30,'% AN','Prospectus art. 23','Surengagement_max',PCT0)
prow('Emprunts et instruments dérivés','Interdits','-','Prospectus art. 22 ; aucune couverture de change possible','Emprunt')
prow('Entreprises Cibles (co-investissements) au plus',0.30,'% AN','Prospectus art. 3','Entreprises_max',PCT0)
P.blank()
P.sec('C. ALLOCATION À LA FIN DU BLOCAGE ET TERMES DE LIQUIDITÉ DES LIGNES','Poids en % de l\'actif net au départ. Non-appelé : part de la ligne encore à appeler (fonds fermés). Gate par trimestre et pression des pairs : fonds evergreen. TTL : délai de conversion en cash en jours, par scénario. Décote : haircut de cession sur le marché secondaire.')
hdr_r=P.hdr(['Ligne','Type','Poids','Non-appelé (% ligne)','Gate trim. (fonds)','Préavis fonds (jours)','TTL normal','TTL S2','TTL S3','TTL S4','Source'],h=30)
LP={}  # line param cells
line_defs=[
 ('MAPIF II','Fermé',0.34,0.26,0,0,365,540,730,1095,'PPM avril 2026 : aucun rachat, cession agréée par le GP ; appels à 10 JO'),
 ('MSIG 3','Fermé',0.17,0.35,0,45,365,540,730,1095,'PPM février 2026 : retrait annuel après lock-up 3 ans payé au run-off ; traité comme fermé sur l\'horizon'),
 ('PG NGI','Ouvert',0.17,0,0.05,63,180,270,365,540,'Supplément 2 : rachats mensuels, préavis 45 JO, gate 5 % de la VL par trimestre, Special Dealing 180 j'),
 ('Ares AGI','Ouvert',0.17,0,0.05,30,120,180,365,540,'Annexe 7 : rachats trimestriels, gate 5 % net par trimestre, limite ELTIF 27,3 %, déduction 5 % avant 24 mois'),
 ('Fonds HY UCITS','Liquide',0.10,0,1,2,7,14,30,30,'OPCVM obligataire à VL quotidienne (prospectus art. 3)'),
 ('Trésorerie','Liquide',0.05,0,1,0,1,1,1,1,'Dépôts et monétaire ; cible 5 % fixée par la Fonction Risques'),
]
for i,(nm_,typ,w,na,g,pre,t0,t2,t3,t4,src) in enumerate(line_defs):
    r=P.row([nm_,typ,w,na,g,pre,t0,t2,t3,t4,src],fmts=[None,None,PCT0,PCT0,PCT0,D,D,D,D,D],inputs=(2,3,4,5,6,7,8,9),h=26)
    LP[i]=dict(row=r,w=f'Paramètres!$D${r}',na=f'Paramètres!$E${r}',gate=f'Paramètres!$F${r}',pre=f'Paramètres!$G${r}',ttl=[f'Paramètres!${c}${r}' for c in 'HIJK'],typ=typ)
r=P.row(['Total','', f'=SUM(D{LP[0]["row"]}:D{LP[5]["row"]})',None,None,None,None,None,None,None,'Doit être égal à 100 %'],fmts=[None,None,PCT0],key=(2,)); P.nm('W_total',r,4)
P.nm('W_cash',LP[5]['row'],4); P.nm('W_HY',LP[4]['row'],4)
P.blank()
P.sec('C bis. HYPOTHÈSES PAR LIGNE : RENDEMENT, DISTRIBUTIONS, PROXY ET TRANSMISSION DES CHOCS','Rendement : TRI net cible du DIC. Distributions : % de la valeur par an. Proxy coté : composite du classeur des cours. Transmission : part du choc du proxy coté répercutée dans la VL de la ligne (1 pour un actif coté ; environ 0,5 pour une VL d\'expertise, rapport VEV privée / VEV cotée du DIC). Retard : part du choc reportée au semestre suivant. Dollar : part de la ligne en dollars non couverts.')
P.hdr(['Ligne','Rendement net annuel','Distributions annuelles','Proxy coté','Transmission','Retard (part au semestre suivant)','Dollar non couvert','Courbe en J (S3, année 1-3)','Haircut vente S2','Haircut vente S3','Haircut vente S4'],h=36)
line_h=[('MAPIF II',0.135,0.06,'C1 Infrastructure cotée',0.5,0.5,1.0,-0.05,0.15,0.25,0.35),
        ('MSIG 3',0.11,0.07,'C3 Direct lending coté',0.5,0.5,1.0,0.0,0.15,0.25,0.35),
        ('PG NGI',0.115,0.02,'C1 Infrastructure cotée',0.5,0.5,0.0,0.0,0.05,0.15,0.25),
        ('Ares AGI',0.095,0.05,'Mix 60 % C1, 20 % C3, 20 % C4',0.5,0.5,0.0,-0.03,0.05,0.15,0.25),
        ('Fonds HY UCITS',0.06,0.0,'C4 Haut rendement coté',1.0,0.0,0.0,0.0,0.01,0.03,0.05),
        ('Trésorerie',0.025,0.0,'Taux monétaire',0.0,0.0,0.0,0.0,0.0,0.0,0.0)]
for i,(nm_,rd,ds,px,tr,lag,usd,jc,h2,h3,h4) in enumerate(line_h):
    r=P.row([nm_,rd,ds,px,tr,lag,usd,jc,h2,h3,h4],fmts=[None,PCT,PCT,None,DEC,DEC,PCT0,PCT,PCT0,PCT0,PCT0],inputs=(1,2,4,5,6,7,8,9,10))
    LP[i].update(rend=f'Paramètres!$C${r}',dist=f'Paramètres!$D${r}',trans=f'Paramètres!$F${r}',lag=f'Paramètres!$G${r}',usd=f'Paramètres!$H${r}',jc=f'Paramètres!$I${r}',hc={'S1':'0','S2':f'Paramètres!$J${r}','S3':f'Paramètres!$K${r}','S4':f'Paramètres!$L${r}'},row2=r)
P.blank()
P.sec('D. HYPOTHÈSES DE FLUX EN RÉGIME NORMAL')
P.hdr(['Paramètre','Valeur','Unité','Source','','','','','','','Commentaire'])
prow('Rachats demandés par date de rachat',0.02,'% AN','Hypothèse : aucun historique','Rachats_norm',PCT)
prow('Collecte par semestre',0.10,'% AN','Prévision du Conseiller à confirmer ; nulle dans les scénarios S3 et S4','Collecte_norm',PCT)
prow('Appels des fonds fermés par semestre',  '=1/6','part du non-appelé initial','Tirage sur trois ans','Appels_base',PCT,inp=False)
prow('Frais totaux du Fonds (part A1)',0.0211,'% AN / an','Classeur DIC v2.3','Frais_an',PCT2)
prow('Trésorerie opérationnelle minimale',0.01,'% AN','Fonction Risques : plancher avant toute vente','Cash_min',PCT)
prow('Trésorerie cible',0.05,'% AN','Fonction Risques (5 % de cash)','Cash_cible',PCT)
prow('Poche liquide cible (trésorerie + fonds HY)',0.16,'% AN','Fonction Risques : un point au-dessus du minimum de 15 % de l\'art. 3','Liq_cible',PCT)
prow('Poche liquide d\'exécution (trésorerie + HY) sous laquelle on désinvestit les fonds cibles',0.10,'% AN','Note de liquidité § 5','Poche_min_exec',PCT)
prow('Part du premier porteur',0.15,'% AN','Hypothèse : registre non constitué','Top1',PCT)
prow('Part des cinq premiers porteurs',0.35,'% AN','Hypothèse','Top5',PCT)
prow('Part de l\'actif net encore sous blocage à la date de départ',0.50,'% AN','Hypothèse : souscriptions étalées sur les deux premières années','Part_bloquee',PCT,note='Réduit les demandes possibles aux premières dates ; sert au profil Annex IV passif.')
P.blank()
P.sec('E. INTERRUPTEURS DE GOUVERNANCE   1 = actif, 0 = inactif')
P.hdr(['Interrupteur','Valeur','Unité','Source','','','','','','','Commentaire'])
prow('Plafonnement appliqué (0 = la Société de Gestion renonce au gate)',1,'1/0','Prospectus art. 7.2.6 : faculté, non obligation','Gate_actif',D)
prow('Suspension automatique après le nombre maximal de reports',1,'1/0','Note de liquidité § 5','Suspension_auto',D)
prow('Prolongation du Délai de Préavis activée',0,'1/0','Prospectus art. 7.2.7','Prolong_actif',D,note='Agit sur le profil de passif (LTTL) ; n\'ajoute pas de cash.')
prow('Cession secondaire autorisée pour honorer les rachats',1,'1/0','Décision de gestion','Cession_autorisee',D)
prow('Profil de flux des tests inversés (1 = S1 normal, 2 = S2, 3 = S3, 4 = S4)',1,'1-4','Fonction Risques : par défaut, flux normaux hors collecte pour isoler l\'effet des rachats','Reverse_flux',D,note='La collecte est nulle dans tous les tests inversés (ESMA 34-39-882).')
prow('Réinvestissement de la trésorerie excédentaire (au-dessus de la cible) dans le fonds HY puis les fonds evergreen',1,'1/0','Décision de gestion','Reinvest_actif',D,note='Les fonds fermés ne reçoivent que leurs appels ; la trésorerie cible est conservée.')
P.blank()
P.sec('F. SEUILS DE VERDICT   (note de liquidité § 5)')
P.hdr(['Seuil','Valeur','Unité','Source','','','','','','','Commentaire'])
prow('File d\'attente : rouge au-delà de',0.25,'% AN','Note de liquidité § 5','S_file_rouge',PCT)
prow('Poche liquide : rouge en dessous de',0.05,'% AN','Note de liquidité § 5','S_poche_rouge',PCT)
prow('Poche liquide : orange en dessous de',0.15,'% AN','Prospectus art. 3 (15 % d\'actifs liquides)','S_poche_orange',PCT)
prow('Couverture à douze mois : rouge en dessous de',1.0,'x','Note de liquidité § 5','S_couv_rouge',DEC)
prow('Couverture à douze mois : orange en dessous de',1.5,'x','Note de liquidité § 5','S_couv_orange',DEC)
prow('Choc de VL (MST) : orange au-delà de',0.10,'%','Pratique de place','S_mst_orange',PCT0)
prow('Choc de VL (MST) : rouge au-delà de',0.25,'%','Pratique de place (Solvabilité II : 49 % pour le non coté)','S_mst_rouge',PCT0)
P.blank()
# ---- G. CALIBRATION LST par scénario et par date
P.sec('G. CALIBRATION DES SCÉNARIOS DE LIQUIDITÉ   par date de rachat (semestres 1 à 8)','Taux de rachat en % de l\'actif net de début de période. Multiplicateurs appliqués aux hypothèses de régime normal. Pression des pairs : part de la capacité de rachat des fonds evergreen consommée par les autres investisseurs (1 = plus rien n\'est récupérable). Porteur 1 : 1 si le premier porteur demande la sortie intégrale à la date 1.')
CAL={}
def cal_block(key,label,unit,fmt,rows):
    P.hdr([label]+[f'Date {t+1}' for t in range(T)]+['Unité','Commentaire'],h=18)
    for (code,_,_),vals,com in zip(SCEN,rows,['','','','']):
        r=P.row([code]+list(vals)+[unit,com],fmts=[None]+[fmt]*T,inputs=tuple(range(1,1+T)))
        CAL[(key,code)]=[f'Paramètres!${PC[t]}${r}' for t in range(T)]
cal_block('rachat','Taux de rachat demandé','% AN',PCT,[
 [0.02]*8,[0.075,0.075,0.05,0.03,0.02,0.02,0.02,0.02],[0.15,0.15,0.10,0.05,0.03,0.02,0.02,0.02],[0.25,0.20,0.15,0.10,0.05,0.03,0.02,0.02]])
cal_block('collecte','Multiplicateur de collecte','x','0.00',[
 [1]*8,[0.5,0.5,0.5,0.75,1,1,1,1],[0,0,0,0,0.25,0.5,0.5,0.5],[0,0,0,0,0,0,0.25,0.25]])
cal_block('dist','Multiplicateur des distributions','x','0.00',[
 [1]*8,[0.7,0.6,0.7,0.85,1,1,1,1],[0.4,0.2,0.25,0.35,0.5,0.6,0.7,0.8],[0.2,0.1,0.1,0.15,0.25,0.4,0.5,0.6]])
cal_block('appels','Multiplicateur des appels de fonds','x','0.00',[
 [1]*8,[1,1,1,1,1,1,1,1],[1.5,1.5,1,1,1,1,1,1],[2,2,1,1,1,1,1,1]])
cal_block('pairs','Pression des pairs sur les gates evergreen','0-1','0.00',[
 [0.3]*8,[0.5,0.5,0.5,0.4,0.3,0.3,0.3,0.3],[1,1,0.8,0.6,0.5,0.4,0.3,0.3],[1,1,1,0.8,0.6,0.5,0.4,0.3]])
cal_block('porteur1','Sortie intégrale du premier porteur','1/0','0',[
 [0]*8,[0]*8,[1,0,0,0,0,0,0,0],[1,0,0,0,0,0,0,0]])
P.blank()
P.sec('H. CALIBRATION DES CHOCS DE MARCHÉ (MST)   en écarts-types semestriels du proxy coté, par semestre','Le même profil s\'applique à chaque facteur coté (infrastructure, direct lending, crédit HY) avec son propre écart-type (onglet MST Données). Négatif = choc, positif = reprise. Le choc dollar est saisi directement en % (EUR/USD : négatif = dollar en baisse contre l\'euro).')
cal_block('sigma','Choc en écarts-types','σ','0.00',[
 [-0.5,0,0,0,0,0,0,0],[-1.5,0,0.5,0.5,0,0,0,0],[-2.5,-1,0,0.75,0.75,0.5,0,0],[-3.5,-1.5,-0.5,0,0.5,0.5,0.5,0.5]])
cal_block('usd','Choc dollar (EUR/USD, % du cours)','%',PCT,[
 [0]*8,[-0.05,0,0,0,0,0,0,0],[-0.15,0,0,0.05,0,0,0,0],[-0.25,-0.05,0,0,0.05,0,0,0]])
P.note('Références de calibration : S2 ≈ percentile 7 % d\'un semestre ; S3 ≈ pire semestre historique des composites (2008 non couvert par l\'historique de 15 ans, d\'où le recours à 2,5 σ) ; S4 ≈ queue à 3,5 σ, cohérente avec le choc Solvabilité II de 49 % appliqué à une VL transmise à moitié. Les chocs sont appliqués aux proxies cotés puis transmis aux VL privées selon la colonne Transmission de l\'onglet (retard d\'un semestre pour la part non transmise immédiatement).',h=44)
PAR_LAST=P.r
# ======================================================================= MST DONNÉES
months,comp,fx=monthly_returns()
MD=Sheet(bk,'MST Données','MST - SÉRIES MENSUELLES DES PROXIES COTÉS ET STATISTIQUES','Rendements mensuels simples des composites équipondérés du classeur des cours (devise locale, dividendes réinvestis ; Brookfield exclu) et variation du cours EUR/USD. Valeurs importées ; statistiques en formules. Source : Yahoo Finance, extraction du 06/10/2026.',
        [2,10,13,13,13,13,13,3,34,13,13,13,13],ncols=12)
MD.hdr(['Mois','C1 Infrastructure cotée','C2 Dette d\'infra cotée','C3 Direct lending coté','C4 Crédit HY coté','EUR/USD (variation)'],h=30)
R0=MD.r
for (y,m) in months:
    MD.row([dt.date(y,m,1),comp['C1'].get((y,m)),comp['C2'].get((y,m)),comp['C3'].get((y,m)),comp['C4'].get((y,m)),fx['EURUSD'].get((y,m))],fmts=[DATE.replace('dd/','') if False else 'mm/yyyy',PCT2,PCT2,PCT2,PCT2,PCT2],bold_first=False)
RL=MD.r-1
FAC={'C1':'C','C2':'D','C3':'E','C4':'F','USD':'G'}
# statistics block to the right (columns I..)
ws=MD.ws
def mstat(r,label,formula_by_col,fmt):
    ws.cell(r,9,label).font=F(True); ws.cell(r,9).border=B_BOT
    for j,k in enumerate(['C1','C2','C3','C4','USD']):
        c=ws.cell(r,10+j,formula_by_col(FAC[k])); c.font=F(); c.number_format=fmt; c.alignment=al('right'); c.border=B_BOT
r=R0-1
for j,k in enumerate(['C1 Infra','C2 Dette infra','C3 Direct lending','C4 Crédit HY','EUR/USD']):
    c=ws.cell(r,10+j,k); c.font=F(True,WHITE); c.fill=fill(DARK); c.alignment=al('center',wrap=True)
c=ws.cell(r,9,'Statistique'); c.font=F(True,WHITE); c.fill=fill(DARK)
rng=lambda col: f'{col}{R0}:{col}{RL}'
STAT={}
def add(label,fn,fmt,key):
    global r
    r+=1; mstat(r,label,fn,fmt); STAT[key]={k:f"'MST Données'!${L(10+j)}${r}" for j,k in enumerate(['C1','C2','C3','C4','USD'])}
add('Nombre de mois',lambda c:f'=COUNT({rng(c)})',D,'n')
add('Rendement mensuel moyen',lambda c:f'=AVERAGE({rng(c)})',PCT2,'mu')
add('Écart-type mensuel',lambda c:f'=STDEV({rng(c)})',PCT2,'sd_m')
add('Volatilité annualisée',lambda c:f'=STDEV({rng(c)})*SQRT(12)',PCT,'vol')
add('Écart-type semestriel (σ × √6)',lambda c:f'=STDEV({rng(c)})*SQRT(6)',PCT,'sd_s')
add('Rendement semestriel moyen (× 6)',lambda c:f'=AVERAGE({rng(c)})*6',PCT,'mu_s')
add('Pire mois',lambda c:f'=MIN({rng(c)})',PCT,'worst_m')
add('Autocorrélation d\'ordre 1',lambda c:f'=CORREL({c}{R0+1}:{c}{RL},{c}{R0}:{c}{RL-1})',D3,'ar1')
add('Asymétrie',lambda c:f'=SKEW({rng(c)})',DEC,'skew')
add('Aplatissement (excès)',lambda c:f'=KURT({rng(c)})',DEC,'kurt')
# rolling 6m window columns: compute in hidden-ish columns O..S : cumulative 6m return
for j,k in enumerate(['C1','C2','C3','C4','USD']):
    col=L(15+j); src=FAC[k]
    c=ws.cell(R0-1,15+j,f'{k} : 6 mois glissants'); c.font=F(True,WHITE); c.fill=fill(DARK); c.alignment=al('center',wrap=True); ws.column_dimensions[col].width=13
    for rr in range(R0+5,RL+1):
        c=ws.cell(rr,15+j,f'=PRODUCT(1+{src}{rr-5}:{src}{rr})-1'); c.number_format=PCT2; c.font=F()
        from openpyxl.worksheet.formula import ArrayFormula
        c.value=ArrayFormula(f'{col}{rr}',f'=PRODUCT(1+{src}{rr-5}:{src}{rr})-1')
ws.column_dimensions['N'].width=3
add('Pire semestre glissant observé',lambda c:f'=MIN({L(15+list(FAC.values()).index(c))}{R0+5}:{L(15+list(FAC.values()).index(c))}{RL})',PCT,'worst_6')
add('Percentile 5 % des semestres glissants',lambda c:f'=PERCENTILE({L(15+list(FAC.values()).index(c))}{R0+5}:{L(15+list(FAC.values()).index(c))}{RL},0.05)',PCT,'p5_6')
add('Percentile 1 % des semestres glissants',lambda c:f'=PERCENTILE({L(15+list(FAC.values()).index(c))}{R0+5}:{L(15+list(FAC.values()).index(c))}{RL},0.01)',PCT,'p1_6')
r+=2
c=ws.cell(r,9,'MATRICE DE CORRÉLATION DES RENDEMENTS MENSUELS'); c.font=F(True,TEAL); r+=1
for j,k in enumerate(['C1','C2','C3','C4','USD']):
    c=ws.cell(r,10+j,k); c.font=F(True,WHITE); c.fill=fill(DARK); c.alignment=al('center')
for i,ki in enumerate(['C1','C2','C3','C4','USD']):
    r+=1; ws.cell(r,9,ki).font=F(True)
    for j,kj in enumerate(['C1','C2','C3','C4','USD']):
        c=ws.cell(r,10+j,f'=CORREL({rng(FAC[ki])},{rng(FAC[kj])})'); c.number_format=DEC; c.font=F(); c.alignment=al('right')
r+=2
c=ws.cell(r,9,'LISSAGE DES VL PRIVÉES (Geltner)  -  entrées du classeur DIC v2.3'); c.font=F(True,TEAL); r+=1
geltner=[('Autocorrélation d\'ordre 1 des rendements trimestriels Preqin Infrastructure',0.1169,D3,'G_ar1'),
         ('Volatilité annualisée brute (Preqin Infrastructure)',0.0297,PCT2,'G_vol_brute'),
         ('Volatilité annualisée désmoothée (Preqin Infrastructure)',0.0334,PCT2,'G_vol_desm'),
         ('Autocorrélation mensuelle après interpolation (variante ordre 2, infra)',0.7126,D3,'G_ar1_m'),
         ('VEV proxy privé désmoothé (DIC)',0.0502,PCT2,'VEV_priv'),
         ('VEV contrôle coté retenue (DIC)',0.1088,PCT2,'VEV_cote')]
for lab,v,fmt,nm_ in geltner:
    ws.cell(r,9,lab).font=F(); c=ws.cell(r,10,v); c.number_format=fmt; c.font=F(False,INPUT_TXT); c.fill=fill(CYAN); c.alignment=al('right'); bk.name(nm_,ws,f'$J${r}'); r+=1
ws.cell(r,9,'Rapport de lissage 1 / (1 - AR1)').font=F(True); c=ws.cell(r,10,'=1/(1-G_ar1)'); c.number_format=DEC; c.font=F(); c.alignment=al('right'); r+=1
ws.cell(r,9,'Transmission implicite des chocs cotés = VEV privée / VEV cotée').font=F(True); c=ws.cell(r,10,'=VEV_priv/VEV_cote'); c.number_format=DEC; c.font=F(); c.alignment=al('right'); bk.name('Transmission_impl',ws,f'$J${r}'); r+=1
ws.cell(r,9,'Volatilité du composite coté C1 / volatilité désmoothée Preqin').font=F(True); c=ws.cell(r,10,f"={STAT['vol']['C1']}/G_vol_desm"); c.number_format=DEC; c.font=F(); c.alignment=al('right'); r+=1
c=ws.cell(r,9,'Lecture : les VL d\'expertise lissent les chocs ; la transmission retenue dans l\'onglet Paramètres (0,5 par défaut) est cohérente avec le rapport VEV privée / VEV cotée. Une transmission de 1 reviendrait à valoriser les fonds cibles comme des actifs cotés.'); c.font=F(False,GREY); c.alignment=al(wrap=True,v='top'); ws.merge_cells(start_row=r,start_column=9,end_row=r,end_column=14); ws.row_dimensions[r].height=40
# ======================================================================= MST SCÉNARIOS
MS=Sheet(bk,'MST Scénarios','MST - CHOCS DE MARCHÉ PAR SCÉNARIO ET PAR SEMESTRE, RENDEMENT DES LIGNES ET TRAJECTOIRE DE VL','Chocs des facteurs cotés = profil en σ (Paramètres H) × écart-type semestriel du facteur (MST Données). Rendement d\'une ligne = dérive (TRI cible / 2) + transmission × choc du proxy (part immédiate et part retardée d\'un semestre) + effet dollar + courbe en J. Trajectoire de VL hors flux de passif (effet marché pur), reprise par le moteur LST.',
        [2,44,12,12,12,12,12,12,12,12,40],ncols=10)
MSR={}  # (code,'line',i) -> row ; (code,'nav') -> row ; (code,'fac',k)
for code,lab,desc in SCEN:
    MS.sec(f'{code} - {lab.upper()}   {desc}')
    MS.hdr(['Semestre']+[f'{t+1}' for t in range(T)]+['Lecture'],h=16)
    rdate=MS.row(['Date de fin de semestre']+[f'=EDATE(Date_depart,{6*t})' for t in range(T)]+[''],fmts=[None]+[DATE]*T)
    MS.row(['Profil de choc (σ)']+[f'={CAL[("sigma",code)][t]}' for t in range(T)]+['Paramètres H'],fmts=[None]+['0.00']*T)
    fr={}
    for k,lab2 in [('C1','Choc C1 Infrastructure cotée'),('C3','Choc C3 Direct lending coté'),('C4','Choc C4 Crédit HY coté')]:
        fr[k]=MS.row([lab2]+[f'={CAL[("sigma",code)][t]}*{STAT["sd_s"][k]}' for t in range(T)]+[f'σ semestriel du facteur × profil'],fmts=[None]+[PCT]*T)
    fr['USD']=MS.row(['Choc dollar (EUR/USD)']+[f'={CAL[("usd",code)][t]}' for t in range(T)]+['Paramètres H, en %'],fmts=[None]+[PCT]*T)
    fr['ARES']=MS.row(['Choc proxy Ares (60 % C1, 20 % C3, 20 % C4)']+[f'=0.6*{PC[t]}{fr["C1"]}+0.2*{PC[t]}{fr["C3"]}+0.2*{PC[t]}{fr["C4"]}' for t in range(T)]+['Mix de la documentation Ares'],fmts=[None]+[PCT]*T)
    MSR[(code,'fac')]=fr
    MS.blank()
    MS.hdr(['Rendement semestriel de la ligne']+[f'{t+1}' for t in range(T)]+['Composition'],h=16)
    for i,(nm_,typ,px) in enumerate(LINES):
        p=LP[i]
        if px is None:
            f=lambda t: f'={p["rend"]}/2'
            comp_txt='Taux monétaire / 2'
        else:
            frow=fr[px]
            def f(t,frow=frow,p=p):
                shock=f'{p["trans"]}*((1-{p["lag"]})*{PC[t]}{frow}'+(f'+{p["lag"]}*{PC[t-1]}{frow}' if t>0 else '')+')'
                usd=f'+{p["usd"]}*{PC[t]}{fr["USD"]}'
                jc=f'+IF(AND({PC[t]}{frow}<0,"{code}"<>"S1",{t+1}<=2),{p["jc"]},0)' if True else ''
                return f'={p["rend"]}/2+{shock}{usd}{jc}'
            comp_txt=f'Dérive + transmission × choc {px} (retard d\'un semestre) + dollar + courbe en J'
        rr=MS.row([nm_]+[f(t) for t in range(T)]+[comp_txt],fmts=[None]+[PCT]*T)
        MSR[(code,'line',i)]=rr
    # portfolio market-only NAV path, weights fixed at start
    MS.blank()
    rp=MS.row(['Rendement du portefeuille (poids de départ)']+[ '='+'+'.join(f'{LP[i]["w"]}*{PC[t]}{MSR[(code,"line",i)]}' for i in range(NL)) for t in range(T)]+['Σ poids × rendement de ligne'],fmts=[None]+[PCT]*T)
    rn=MS.row(['Indice de VL hors flux (base 100)']+[(f'=100*(1+{PC[0]}{rp})' if t==0 else f'={PC[t-1]}{MS.r}*(1+{PC[t]}{rp})') for t in range(T)]+['Effet marché seul'],fmts=[None]+['0.0']*T)
    rdd=MS.row(['Écart au départ']+[f'={PC[t]}{rn}/100-1' for t in range(T)]+[''],fmts=[None]+[PCT]*T)
    rmin=MS.row(['Pire écart cumulé (creux)',f'=MIN(C{rdd}:J{rdd})',None,None,None,None,None,None,None,'Mesure retenue pour le verdict MST'],fmts=[None,PCT],key=(1,))
    rrec=MS.row(['Semestre de retour au-dessus du niveau de départ',f'=IFERROR(MATCH(TRUE,INDEX(C{rdd}:J{rdd}>=0,0),0),"au-delà de 8")',None,None,None,None,None,None,None,'Premier semestre où l\'indice repasse 100'],fmts=[None,D])
    rv=MS.row(['Verdict MST',f'=IF(-C{rmin}>S_mst_rouge,"Rouge",IF(-C{rmin}>S_mst_orange,"Orange","Vert"))',None,None,None,None,None,None,None,'Vert ≤ 10 % ; Orange 10 à 25 % ; Rouge > 25 %'],key=(1,))
    MS.ws.cell(rv,3).alignment=al('center')
    MSR[(code,'nav')]=dict(rp=rp,rn=rn,rdd=rdd,rmin=rmin,rrec=rrec,rv=rv,rdate=rdate)
    bk.name(f'MST_{code}_creux',MS.ws,f'$C${rmin}'); bk.name(f'MST_{code}_verdict',MS.ws,f'$C${rv}'); bk.name(f'MST_{code}_reprise',MS.ws,f'$C${rrec}')
    MS.blank()
verdict_cf(MS.ws,f'C6:C{MS.r}')
# ======================================================================= ALP
A=Sheet(bk,'ALP','ALP - PROFIL DE LIQUIDITÉ DE L\'ACTIF (look-through, date de départ)','Valeur de chaque ligne à la fin du blocage, délai de conversion en cash (TTL) par scénario et classement par tranche Annex IV. Les fonds evergreen ne sont pas liquides à 90 jours : préavis, gate trimestriel de 5 % et pression des pairs portent le délai de récupération complète au-delà de 180 jours. Le non-appelé des fonds fermés est une sortie, pas une ressource.',
        [2,34,14,10,10,10,10,10,10,10,10,10,10,10],ncols=13)
A.sec('A. VALEURS ET DÉLAIS DE CONVERSION')
A.hdr(['Ligne','Valeur (EUR)','% AN','TTL normal','TTL S2','TTL S3','TTL S4','Non-appelé (EUR)','Capacité de rachat par semestre, normal (EUR)','Haircut S2','Haircut S3','Haircut S4',''],h=40)
AR={}
for i,(nm_,typ,px) in enumerate(LINES):
    p=LP[i]
    cap=(f'=C{A.r}*{p["gate"]}*2*(1-{CAL[("pairs","S1")][0]})' if typ=='Ouvert' else (f'=C{A.r}' if typ=='Liquide' else '=0'))
    r=A.row([nm_,f'={p["w"]}*AN_depart',f'=C{A.r}/AN_depart',f'={p["ttl"][0]}',f'={p["ttl"][1]}',f'={p["ttl"][2]}',f'={p["ttl"][3]}',f'=C{A.r}*{p["na"]}',cap,f'={p["hc"]["S2"]}',f'={p["hc"]["S3"]}',f'={p["hc"]["S4"]}'],fmts=[None,EUR,PCT,D,D,D,D,EUR,EUR,PCT0,PCT0,PCT0])
    AR[i]=r
f0,l0=AR[0],AR[NL-1]
rt=A.row(['Total',f'=SUM(C{f0}:C{l0})',f'=SUM(D{f0}:D{l0})',None,None,None,None,f'=SUM(I{f0}:I{l0})',f'=SUM(J{f0}:J{l0})'],fmts=[None,EUR,PCT,None,None,None,None,EUR,EUR],key=(1,2,7,8))
bk.name('NA_total',A.ws,f'$I${rt}'); bk.name('Cap_normal',A.ws,f'$J${rt}')
A.row(['Sur-engagement au départ (AN + non-appelé) / AN',f'=(C{rt}+I{rt})/C{rt}',None,None,None,None,None,None,'=IF((C'+str(rt)+'+I'+str(rt)+')/C'+str(rt)+'>Surengagement_max,"Rouge : plafond de l\'art. 23 dépassé","Vert : sous le plafond de l\'art. 23")'],fmts=[None,PCT0],merge=[(10,14)])
A.blank()
A.sec('B. DÉLAI MOYEN PONDÉRÉ DE CONVERSION (WATTL) ET T90','WATTL = Σ valeur × TTL / actif net. T90 = TTL de la ligne où le cumul des valeurs triées par TTL atteint 90 % de l\'actif net.')
A.hdr(['Mesure','Normal','S2','S3','S4','','','','','','','',''],h=16)
cols=['E','F','G','H']  # TTL columns in table A are E..H? check: B ligne, C valeur, D %AN, E TTL normal, F S2, G S3, H S4
rw=A.row(['WATTL (jours)']+[f'=SUMPRODUCT(C{f0}:C{l0},{c}{f0}:{c}{l0})/C{rt}' for c in cols],fmts=[None]+[D]*4,key=(1,2,3,4))
for j,s in enumerate(['N','S2','S3','S4']): bk.name(f'WATTL_{s}',A.ws,f'${L(3+j)}${rw}')
def t90(c):
    # smallest TTL such that sum of value with TTL<=x >= 90% of total : iterate over candidate TTLs
    return f'=MIN(IF(SUMIF({c}{f0}:{c}{l0},"<="&{c}{f0}:{c}{l0},C{f0}:C{l0})>=0.9*C{rt},{c}{f0}:{c}{l0},100000))'
from openpyxl.worksheet.formula import ArrayFormula
r90=A.row(['T90 (jours pour liquider 90 % de l\'actif)']+[None]*4,fmts=[None]+[D]*4)
for j,c in enumerate(cols):
    cell=A.ws.cell(r90,3+j); cell.value=ArrayFormula(f'{L(3+j)}{r90}',t90(c)); cell.number_format=D; cell.font=F(); cell.alignment=al('right'); cell.border=B_BOT
for j,s in enumerate(['N','S2','S3','S4']): bk.name(f'T90_{s}',A.ws,f'${L(3+j)}${r90}')
A.row(['Actifs liquides à 30 jours (HQLA), % AN']+[f'=SUMIF({c}{f0}:{c}{l0},"<=30",C{f0}:C{l0})/C{rt}' for c in cols],fmts=[None]+[PCT]*4)
rh=A.row(['HQLA après haircut, % AN',f'=SUMIF(E{f0}:E{l0},"<=30",C{f0}:C{l0})/C{rt}']+[f'=SUMPRODUCT(({c}{f0}:{c}{l0}<=30)*C{f0}:C{l0}*(1-{h}{f0}:{h}{l0}))/C{rt}' for c,h in zip(cols[1:],['K','L','M'])],fmts=[None]+[PCT]*4,key=(1,2,3,4))
for j,s in enumerate(['N','S2','S3','S4']): bk.name(f'HQLA_{s}',A.ws,f'${L(3+j)}${rh}')
A.row(['Capacité de rachat des fonds evergreen sur douze mois, % AN','=Cap_normal*2/AN_depart',f'=Cap_normal*2*(1-{CAL[("pairs","S2")][0]})/(1-{CAL[("pairs","S1")][0]})/AN_depart',f'=Cap_normal*2*(1-{CAL[("pairs","S3")][0]})/(1-{CAL[("pairs","S1")][0]})/AN_depart',f'=Cap_normal*2*(1-{CAL[("pairs","S4")][0]})/(1-{CAL[("pairs","S1")][0]})/AN_depart'],fmts=[None]+[PCT]*4)
A.blank()
A.sec('C. PROFIL ANNEX IV, CÔTÉ ACTIF   % de l\'actif net liquidable par tranche (AIFMD Annex IV, section 2, question 178)')
A.hdr(['Tranche','Normal','S2','S3','S4','Champ Annex IV','','','','','','',''],h=16)
buckets=[('1 jour',0,1,'1 day'),('2 à 7 jours',1,7,'2-7 days'),('8 à 30 jours',7,30,'8-30 days'),('31 à 90 jours',30,90,'31-90 days'),('91 à 180 jours',90,180,'91-180 days'),('181 à 365 jours',180,365,'181-365 days'),('Plus de 365 jours',365,100000,'more than 365 days')]
rb0=A.r
for lab,lo,hi,field in buckets:
    A.row([lab]+[f'=SUMPRODUCT(({c}{f0}:{c}{l0}>{lo})*({c}{f0}:{c}{l0}<={hi})*C{f0}:C{l0})/C{rt}' for c in cols]+[field],fmts=[None]+[PCT]*4)
A.row(['Total']+[f'=SUM({L(3+j)}{rb0}:{L(3+j)}{A.r-1})' for j in range(4)]+['Doit être égal à 100 %'],fmts=[None]+[PCT]*4,key=(1,2,3,4))
A.blank()
A.note('Convention TTL : trésorerie 1 jour ; OPCVM HY 7 jours (30 en S3) ; fonds evergreen 180 jours en régime normal (préavis 45 JO ou 30 jours, gate trimestriel, règlement), 365 jours et plus en S3 lorsque les pairs saturent le gate ; fonds fermés 365 jours (cession secondaire agréée) à 1 095 jours en S4 (marché secondaire fermé). Les haircuts de cession s\'appliquent aux ventes forcées dans le moteur LST.',h=44)
# ======================================================================= LLP
Lq=Sheet(bk,'LLP','LLP - PROFIL DE LIQUIDITÉ DU PASSIF ET ADÉQUATION ACTIF-PASSIF','Délai de liquidité du passif (LTTL) du point de vue du porteur : préavis, attente de la prochaine date de rachat, publication et règlement. Comparaison au WATTL de l\'actif par scénario. Profil Annex IV côté passif.',
        [2,52,14,14,14,14,40],ncols=6)
Lq.sec('A. DÉLAI DE LIQUIDITÉ DU PASSIF (LTTL), EN JOURS CALENDAIRES')
Lq.hdr(['Composante','Minimum','Moyen','Maximum','Préavis prolongé','Source'],h=16)
r1=Lq.row(['Préavis (centralisation à VL)','=Preavis_j','=Preavis_j','=Preavis_j','=Preavis_max_j','Prospectus art. 7.2.3 et 7.2.7'],fmts=[None,D,D,D,D])
r2=Lq.row(['Attente de la prochaine date de centralisation',0,'=ROUND(365/Rachats_par_an/2,0)','=ROUND(365/Rachats_par_an,0)','=ROUND(365/Rachats_par_an/2,0)','Deux dates par an : 0 à 182 jours, 91 en moyenne'],fmts=[None,D,D,D,D])
r3=Lq.row(['Publication de la VL et règlement','=Reglement_cal','=Reglement_cal','=Reglement_cal','=Reglement_cal','Prospectus art. 6 et 7.2.4'],fmts=[None,D,D,D,D])
rl=Lq.row(['LTTL hors blocage',f'=SUM(C{r1}:C{r3})',f'=SUM(D{r1}:D{r3})',f'=SUM(E{r1}:E{r3})',f'=SUM(F{r1}:F{r3})','Délai entre la décision de sortir et le cash'],fmts=[None,D,D,D,D],key=(1,2,3,4))
for j,s in enumerate(['min','moy','max','prol']): bk.name(f'LTTL_{s}',Lq.ws,f'${L(3+j)}${rl}')
Lq.row(['Porteurs encore sous blocage à la date de départ','=Part_bloquee',None,None,None,'Paramètres D : part de l\'actif net souscrite depuis moins de 24 mois'],fmts=[None,PCT])
Lq.row(['LTTL moyen pondéré du passif (blocage résiduel moyen de 12 mois pour la part bloquée)','=(1-Part_bloquee)*LTTL_moy+Part_bloquee*(365+LTTL_moy)',None,None,None,'Dérivé'],fmts=[None,D],key=(1,)); bk.name('LTTL_pond',Lq.ws,f'$C${Lq.r-1}')
Lq.blank()
Lq.sec('B. ADÉQUATION ACTIF-PASSIF   écart = WATTL actif - LTTL moyen (jours). Positif : l\'actif se convertit plus lentement que le passif ne sort ; l\'écart est porté par la poche liquide, les flux entrants et le plafonnement.')
Lq.hdr(['Mesure','Normal','S2','S3','S4','Lecture'],h=16)
Lq.row(['WATTL actif (jours)','=WATTL_N','=WATTL_S2','=WATTL_S3','=WATTL_S4','Onglet ALP'],fmts=[None,D,D,D,D])
Lq.row(['LTTL moyen du passif (jours)','=LTTL_moy','=LTTL_moy','=LTTL_moy','=LTTL_prol','S4 : préavis prolongé à 180 jours (art. 7.2.7)'],fmts=[None,D,D,D,D])
rm=Lq.row(['Écart actif - passif (jours)','=C'+str(Lq.r-2)+'-C'+str(Lq.r-1),'=D'+str(Lq.r-2)+'-D'+str(Lq.r-1),'=E'+str(Lq.r-2)+'-E'+str(Lq.r-1),'=F'+str(Lq.r-2)+'-F'+str(Lq.r-1),'Vert < 100 ; Orange 100 à 250 ; Rouge > 250'],fmts=[None,D,D,D,D],key=(1,2,3,4))
for j,s in enumerate(['N','S2','S3','S4']): bk.name(f'Ecart_TTL_{s}',Lq.ws,f'${L(3+j)}${rm}')
Lq.row(['Verdict']+[f'=IF({L(3+j)}{rm}>250,"Rouge",IF({L(3+j)}{rm}>100,"Orange","Vert"))' for j in range(4)]+['Le plafonnement de 5 % est la réponse structurelle à cet écart'],key=(1,2,3,4))
verdict_cf(Lq.ws,f'C{Lq.r-1}:F{Lq.r-1}')
Lq.row(['Couverture de la sortie maximale sur douze mois (2 × gate) par l\'actif convertible dans le LTTL maximal, S3',f'=SUMPRODUCT((ALP!G{f0}:G{l0}<=LTTL_max)*ALP!C{f0}:C{l0}*(1-ALP!L{f0}:L{l0}))/(2*Gate*AN_depart)',None,None,None,'Doit dépasser 1 : la poche liquide couvre deux dates plafonnées'],fmts=[None,DEC],key=(1,)); bk.name('Couv_LTTL_S3',Lq.ws,f'$C${Lq.r-1}')
Lq.blank()
Lq.sec('C. PROFIL ANNEX IV, CÔTÉ PASSIF   % de l\'actif net rachetable par tranche (question 186), hypothèse de porteurs répartis uniformément')
Lq.hdr(['Tranche','Normal','S2 (gate actif)','S3 (gate et suspension)','Préavis prolongé','Champ Annex IV'],h=16)
rp0=Lq.r
Lq.row(['1 jour',0,0,0,0,'1 day'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['2 à 7 jours',0,0,0,0,'2-7 days'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['8 à 30 jours',0,0,0,0,'8-30 days'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['31 à 90 jours',0,0,0,0,'31-90 days'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['91 à 180 jours','=(1-Part_bloquee)*0.5','=MIN(Gate,(1-Part_bloquee)*0.5)',0,0,'91-180 days ; une date de rachat au plus dans la fenêtre'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['181 à 365 jours','=(1-Part_bloquee)*0.5','=MIN(Gate,(1-Part_bloquee)*0.5)','=MIN(Gate,(1-Part_bloquee))','=MIN(2*Gate,(1-Part_bloquee))','181-365 days ; deuxième date, LTTL max 322 jours'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['Plus de 365 jours',f'=1-SUM(C{rp0}:C{Lq.r-1})',f'=1-SUM(D{rp0}:D{Lq.r-1})',f'=1-SUM(E{rp0}:E{Lq.r-1})',f'=1-SUM(F{rp0}:F{Lq.r-1})','more than 365 days ; porteurs bloqués et demandes reportées'],fmts=[None,PCT,PCT,PCT,PCT])
Lq.row(['Total']+[f'=SUM({L(3+j)}{rp0}:{L(3+j)}{Lq.r-1})' for j in range(4)]+[''],fmts=[None]+[PCT]*4,key=(1,2,3,4))
Lq.note('Hypothèses : les porteurs hors blocage demandent au plus le plafond par date ; en S3 la première date est plafonnée puis suspendue ; en préavis prolongé la première date glisse de 89 jours. Ce profil alimente la question 186 du reporting Annex IV ; le profil d\'actif (ALP) la question 178.',h=30)
# ======================================================================= LST MOTEUR
E=Sheet(bk,'LST Moteur','LST - MOTEUR DE PROJECTION PAR SCÉNARIO : HUIT DATES DE RACHAT SEMESTRIELLES','Pour chaque scénario : passif (demandes, report, plafond, suspension), flux des lignes (appels, distributions, chocs de marché du MST), cascade de ressources (trésorerie au-dessus du plancher, vente du fonds HY, rachats evergreen sous gate et pression des pairs, cession secondaire décotée), puis indicateurs et verdict par date. Tous les paramètres viennent de l\'onglet Paramètres.',
        [2,52,13,13,13,13,13,13,13,13,40],ncols=10)
ER={}   # (code,key)->row
def put(code,key,label,fn,fmt,note='',key_row=False,bold=False):
    vals=[label]+[fn(t) for t in range(T)]+[note]
    r=E.row(vals,fmts=[None]+[fmt]*T,key=(1,2,3,4,5,6,7,8) if key_row else ())
    if bold: E.ws.cell(r,2).font=F(True)
    ER[(code,key)]=r; return r
def MR(code,i,t): return "'MST Scénarios'!"+PC[t]+str(MSR[(code,'line',i)])
def R(code,key,t,prev=False):
    return f'{PC[t-1] if prev else PC[t]}{{{key}}}'
for code,lab,desc in SCEN:
    E.sec(f'{code} - {lab.upper()}   {desc}')
    E.hdr(['Date de rachat']+[f'{t+1}' for t in range(T)]+['Formule et lecture'],h=16)
    put(code,'date','Date de VL de rachat',lambda t:f'=EDATE(Date_depart,{6*t})',DATE)
    # NAV début
    put(code,'an0','Actif net de début de période',lambda t:('=AN_depart' if t==0 else f'={PC[t-1]}{{anfin}}'),EUR,'t = 1 : AN de départ ; ensuite AN de fin précédent')
    # passif
    put(code,'subs','Souscriptions versées sur la période',lambda t:f'=Collecte_norm*{CAL[("collecte",code)][t]}*{R(code,"an0",t)}',EUR,'Collecte normale × multiplicateur du scénario')
    put(code,'dem','Demandes de rachat nouvelles',lambda t:f'=({CAL[("rachat",code)][t]}+{CAL[("porteur1",code)][t]}*Top1)*{R(code,"an0",t)}',EUR,'Taux du scénario × AN, plus sortie du premier porteur le cas échéant')
    put(code,'rep_in','Demandes reportées des dates précédentes',lambda t:('=0' if t==0 else f'={R(code,"rep_out",t,True)}'),EUR,'Report automatique (art. 7.2.6)')
    put(code,'dem_tot','Demandes totales à la date',lambda t:f'={R(code,"dem",t)}+{R(code,"rep_in",t)}',EUR)
    put(code,'plafond','Montant exécutable (plafond net des souscriptions)',lambda t:f'=IF(Gate_actif=1,Gate*{R(code,"an0",t)}+{R(code,"subs",t)},{R(code,"dem_tot",t)})',EUR,'Gate × AN + souscriptions versées ; sans limite si le gate est levé')
    put(code,'susp','Suspension des rachats active (1/0)',lambda t:('=0' if t==0 else f'=IF(OR({R(code,"susp",t,True)}=1,AND(Suspension_auto=1,{R(code,"consec",t,True)}>=Reports_max)),1,0)'),D,'Déclenchée après le nombre maximal de reports consécutifs (Paramètres E)')
    put(code,'exec','Rachats exécutés',lambda t:f'=IF({R(code,"susp",t)}=1,0,MIN({R(code,"dem_tot",t)},{R(code,"plafond",t)}))',EUR,'Zéro en suspension')
    put(code,'rep_out','Demandes reportées à la date suivante',lambda t:f'={R(code,"dem_tot",t)}-{R(code,"exec",t)}',EUR)
    put(code,'gated','Date plafonnée ou suspendue (1/0)',lambda t:f'=IF({R(code,"rep_out",t)}>1,1,0)',D)
    put(code,'consec','Dates plafonnées consécutives',lambda t:('=C'+str(ER[(code,'gated')]) if t==0 else f'=IF({R(code,"gated",t)}=1,{R(code,"consec",t,True)}+1,0)'),D)
    put(code,'file','File d\'attente en % de l\'AN de début',lambda t:f'={R(code,"rep_out",t)}/{R(code,"an0",t)}',PCT)
    E.blank()
    # lignes : valeurs début
    for i,(nm_,typ,px) in enumerate(LINES[:5]):
        p=LP[i]
        put(code,f'v0_{i}',f'{nm_} - valeur de début',lambda t,i=i,p=p:(f'={p["w"]}*AN_depart' if t==0 else f'={PC[t-1]}{{vfin_{i}}}'),EUR)
    put(code,'cash0','Trésorerie de début',lambda t:('=W_cash*AN_depart' if t==0 else f'={PC[t-1]}{{cashfin}}'),EUR)
    put(code,'na0','Non-appelé de début (fonds fermés)',lambda t:('=NA_total' if t==0 else f'={PC[t-1]}{{nafin}}'),EUR,'Engagements restant à appeler')
    E.blank()
    put(code,'appels','Appels de fonds des fonds fermés',lambda t:f'=MIN({R(code,"na0",t)},NA_total*Appels_base*{CAL[("appels",code)][t]})',EUR,'Rythme normal × multiplicateur ; MAPIF II 60 %, MSIG 3 40 % du non-appelé')
    put(code,'dist_f','Distributions des fonds fermés',lambda t:f'=({R(code,"v0_0",t)}*{LP[0]["dist"]}+{R(code,"v0_1",t)}*{LP[1]["dist"]})/2*{CAL[("dist",code)][t]}',EUR,'Valeur × taux annuel / 2 × multiplicateur')
    put(code,'dist_o','Distributions des fonds evergreen',lambda t:f'=({R(code,"v0_2",t)}*{LP[2]["dist"]}+{R(code,"v0_3",t)}*{LP[3]["dist"]})/2*{CAL[("dist",code)][t]}',EUR,'Imputées sur la capacité de rachat des fonds (PG NGI)')
    put(code,'frais','Frais du Fonds',lambda t:f'=Frais_an/2*{R(code,"an0",t)}',EUR)
    put(code,'int','Produits de trésorerie',lambda t:f'={R(code,"cash0",t)}*{LP[5]["rend"]}/2',EUR)
    put(code,'besoin','Besoin de trésorerie brut',lambda t:f'={R(code,"exec",t)}+{R(code,"appels",t)}+{R(code,"frais",t)}-{R(code,"subs",t)}-{R(code,"dist_f",t)}-{R(code,"dist_o",t)}-{R(code,"int",t)}',EUR,'Sorties moins entrées de la période')
    put(code,'dispo','Trésorerie mobilisable au-dessus du plancher',lambda t:f'=MAX(0,{R(code,"cash0",t)}-Cash_min*{R(code,"an0",t)})',EUR,'Plancher opérationnel (Paramètres D)')
    put(code,'besoin2','Besoin après trésorerie',lambda t:f'=MAX(0,{R(code,"besoin",t)}-{R(code,"dispo",t)})',EUR)
    put(code,'vente_hy','Vente du fonds HY',lambda t:f'=MIN({R(code,"besoin2",t)},{R(code,"v0_4",t)}*(1+{MR(code,4,t)}))',EUR,'Première ressource après la trésorerie')
    put(code,'perte_hy','Décote réalisée sur la vente HY',lambda t:f'={R(code,"vente_hy",t)}*{LP[4]["hc"][code]}',EUR,'Haircut S2 à S4 (Paramètres C bis)')
    put(code,'besoin3','Besoin après vente HY',lambda t:f'=MAX(0,{R(code,"besoin2",t)}-{R(code,"vente_hy",t)}*(1-{LP[4]["hc"][code]}))',EUR)
    put(code,'cap_evg','Capacité de rachat des fonds evergreen (gates, pairs, distributions)',lambda t:f'=MAX(0,({R(code,"v0_2",t)}*{LP[2]["gate"]}+{R(code,"v0_3",t)}*{LP[3]["gate"]})*2*(1-{CAL[("pairs",code)][t]})-{R(code,"dist_o",t)})',EUR,'Deux trimestres de gate à 5 %, part des pairs déduite')
    put(code,'cible_evg','Rachats evergreen demandés : besoin résiduel et reconstitution de la poche d\'exécution',lambda t:f'=MAX({R(code,"besoin3",t)},Poche_min_exec*{R(code,"an0",t)}-({R(code,"cash0",t)}-{R(code,"besoin",t)}+{R(code,"vente_hy",t)}*(1-{LP[4]["hc"][code]})+{R(code,"v0_4",t)}*(1+{MR(code,4,t)})-{R(code,"vente_hy",t)}))',EUR,'La Fonction Risques demande le rachat dès que trésorerie + HY passent sous la poche d\'exécution (Paramètres D)')
    put(code,'rach_evg','Rachats obtenus des fonds evergreen',lambda t:f'=MAX(0,MIN({R(code,"cible_evg",t)},{R(code,"cap_evg",t)}))',EUR,'Au prorata PG NGI / Ares AGI, dans la limite de la capacité sous gate')
    put(code,'besoin4','Besoin résiduel',lambda t:f'=MAX(0,{R(code,"besoin3",t)}-{R(code,"rach_evg",t)})',EUR)
    put(code,'cession','Cession secondaire des fonds fermés (valeur cédée)',lambda t:f'=IF(Cession_autorisee=1,MIN({R(code,"besoin4",t)}/(1-{LP[0]["hc"][code]}),{R(code,"v0_0",t)}+{R(code,"v0_1",t)}),0)',EUR,'Montant à céder pour encaisser le besoin après décote')
    put(code,'perte_c','Décote réalisée sur la cession',lambda t:f'={R(code,"cession",t)}*{LP[0]["hc"][code]}',EUR)
    put(code,'defaut','Défaut de liquidité (besoin non couvert)',lambda t:f'=MAX(0,{R(code,"besoin4",t)}-{R(code,"cession",t)}*(1-{LP[0]["hc"][code]}))',EUR,'Doit rester nul ; sinon la suspension s\'impose')
    put(code,'cash_pre','Trésorerie avant réinvestissement',lambda t:f'={R(code,"cash0",t)}-{R(code,"besoin",t)}+{R(code,"vente_hy",t)}*(1-{LP[4]["hc"][code]})+{R(code,"rach_evg",t)}+{R(code,"cession",t)}*(1-{LP[0]["hc"][code]})+{R(code,"defaut",t)}',EUR,'Le défaut est réintégré pour garder un solde cohérent (rachat non payé)')
    put(code,'exces','Excédent de trésorerie au-dessus de la cible',lambda t:f'=MAX(0,{R(code,"cash_pre",t)}-Cash_cible*({R(code,"an0",t)}+{R(code,"subs",t)}-{R(code,"exec",t)}))*Reinvest_actif',EUR,'Réinvesti si l\'interrupteur est actif (Paramètres E)')
    put(code,'reinv_hy','Réinvestissement dans le fonds HY (retour au poids cible)',lambda t:f'=MIN({R(code,"exces",t)},MAX(0,(Liq_cible-Cash_cible)*({R(code,"an0",t)}+{R(code,"subs",t)}-{R(code,"exec",t)})-({R(code,"v0_4",t)}*(1+{MR(code,4,t)})-{R(code,"vente_hy",t)})))',EUR)
    put(code,'reinv_evg','Réinvestissement dans les fonds evergreen (moitié PG NGI, moitié Ares AGI)',lambda t:f'={R(code,"exces",t)}-{R(code,"reinv_hy",t)}',EUR,'Les fonds fermés ne reçoivent que les appels')
    put(code,'cashfin','Trésorerie de fin',lambda t:f'={R(code,"cash_pre",t)}-{R(code,"reinv_hy",t)}-{R(code,"reinv_evg",t)}',EUR)
    E.blank()
    # valeurs fin par ligne
    share=lambda t,i:(f'{R(code,"v0_"+str(i),t)}/({R(code,"v0_0",t)}+{R(code,"v0_1",t)})')
    put(code,'vfin_0','MAPIF II - valeur de fin',lambda t:f'={R(code,"v0_0",t)}*(1+{MR(code,0,t)})+{R(code,"appels",t)}*{LP[0]["na"]}*{LP[0]["w"]}/({LP[0]["na"]}*{LP[0]["w"]}+{LP[1]["na"]}*{LP[1]["w"]})-{R(code,"v0_0",t)}*{LP[0]["dist"]}/2*{CAL[("dist",code)][t]}-{R(code,"cession",t)}*{share(t,0)}',EUR,'Valeur × (1 + rendement MST) + appels - distributions - cession')
    put(code,'vfin_1','MSIG 3 - valeur de fin',lambda t:f'={R(code,"v0_1",t)}*(1+{MR(code,1,t)})+{R(code,"appels",t)}*{LP[1]["na"]}*{LP[1]["w"]}/({LP[0]["na"]}*{LP[0]["w"]}+{LP[1]["na"]}*{LP[1]["w"]})-{R(code,"v0_1",t)}*{LP[1]["dist"]}/2*{CAL[("dist",code)][t]}-{R(code,"cession",t)}*{share(t,1)}',EUR)
    evs=lambda t,i:(f'{R(code,"v0_"+str(i),t)}/({R(code,"v0_2",t)}+{R(code,"v0_3",t)})')
    put(code,'vfin_2','PG NGI - valeur de fin',lambda t:f'={R(code,"v0_2",t)}*(1+{MR(code,2,t)})-{R(code,"v0_2",t)}*{LP[2]["dist"]}/2*{CAL[("dist",code)][t]}-{R(code,"rach_evg",t)}*{evs(t,2)}+{R(code,"reinv_evg",t)}/2',EUR,'Rendement MST, distributions, rachats obtenus au prorata, réinvestissement')
    put(code,'vfin_3','Ares AGI - valeur de fin',lambda t:f'={R(code,"v0_3",t)}*(1+{MR(code,3,t)})-{R(code,"v0_3",t)}*{LP[3]["dist"]}/2*{CAL[("dist",code)][t]}-{R(code,"rach_evg",t)}*{evs(t,3)}+{R(code,"reinv_evg",t)}/2',EUR)
    put(code,'vfin_4','Fonds HY UCITS - valeur de fin',lambda t:f'={R(code,"v0_4",t)}*(1+{MR(code,4,t)})-{R(code,"vente_hy",t)}+{R(code,"reinv_hy",t)}',EUR)
    put(code,'nafin','Non-appelé de fin',lambda t:f'={R(code,"na0",t)}-{R(code,"appels",t)}-{R(code,"cession",t)}*{R(code,"na0",t)}/({R(code,"v0_0",t)}+{R(code,"v0_1",t)}+{R(code,"na0",t)})',EUR,'La cession emporte sa part de non-appelé')
    put(code,'anfin','Actif net de fin',lambda t:f'={R(code,"vfin_0",t)}+{R(code,"vfin_1",t)}+{R(code,"vfin_2",t)}+{R(code,"vfin_3",t)}+{R(code,"vfin_4",t)}+{R(code,"cashfin",t)}',EUR,'Somme des lignes et de la trésorerie',bold=True)
    put(code,'ctrl','Contrôle : AN fin - (AN début + souscriptions - rachats - frais + P&L lignes - décotes)',lambda t:f'=ROUND({R(code,"anfin",t)}-({R(code,"an0",t)}+{R(code,"subs",t)}-{R(code,"exec",t)}+{R(code,"defaut",t)}-{R(code,"frais",t)}+{R(code,"int",t)}+{R(code,"v0_0",t)}*{MR(code,0,t)}+{R(code,"v0_1",t)}*{MR(code,1,t)}+{R(code,"v0_2",t)}*{MR(code,2,t)}+{R(code,"v0_3",t)}*{MR(code,3,t)}+{R(code,"v0_4",t)}*{MR(code,4,t)}-{R(code,"perte_hy",t)}-{R(code,"perte_c",t)}),0)',EUR,'Doit être nul')
    E.blank()
    # indicateurs
    put(code,'poche','Poche liquide de fin (trésorerie + HY), % AN',lambda t:f'=({R(code,"cashfin",t)}+{R(code,"vfin_4",t)})/{R(code,"anfin",t)}',PCT,'Rouge < 5 % ; Orange < 15 %')
    put(code,'couv','Couverture à douze mois : ressources mobilisables / deux dates au plafond',lambda t:f'=({R(code,"cashfin",t)}+{R(code,"vfin_4",t)}+2*{R(code,"cap_evg",t)}+2*({R(code,"dist_f",t)}+{R(code,"dist_o",t)})-2*MIN({R(code,"nafin",t)},NA_total*Appels_base*{CAL[("appels",code)][min(t+1,T-1)]}))/(2*Gate*{R(code,"anfin",t)})',DEC,'Rouge < 1,0 ; Orange < 1,5')
    put(code,'sureng','Sur-engagement (AN + non-appelé) / AN',lambda t:f'=({R(code,"anfin",t)}+{R(code,"nafin",t)})/{R(code,"anfin",t)}',PCT0,'Plafond 130 % (art. 23)')
    put(code,'usd','Exposition dollar non couverte, % AN',lambda t:f'=({R(code,"vfin_0",t)}*{LP[0]["usd"]}+{R(code,"vfin_1",t)}*{LP[1]["usd"]})/{R(code,"anfin",t)}',PCT)
    put(code,'gel','Triple gel (collecte, distributions et capacité evergreen simultanément réduites de moitié)',lambda t:f'=IF(AND({CAL[("collecte",code)][t]}<0.5,{CAL[("dist",code)][t]}<0.5,(1-{CAL[("pairs",code)][t]})<0.5*(1-{CAL[("pairs","S1")][t]})),1,0)',D,'Signature de risque evergreen (ESMA 34-39-882)')
    put(code,'verdict','Verdict de la date',lambda t:f'=IF(OR({R(code,"susp",t)}=1,{R(code,"defaut",t)}>1,{R(code,"file",t)}>S_file_rouge,{R(code,"poche",t)}<S_poche_rouge,{R(code,"couv",t)}<S_couv_rouge),"Rouge",IF(OR({R(code,"gated",t)}=1,{R(code,"poche",t)}<S_poche_orange,{R(code,"couv",t)}<S_couv_orange,{R(code,"gel",t)}=1),"Orange","Vert"))',None,'Seuils de l\'onglet Paramètres F',key_row=True)
    verdict_cf(E.ws,f'C{ER[(code,"verdict")]}:J{ER[(code,"verdict")]}')
    for cc in range(3,3+T): E.ws.cell(ER[(code,'verdict')],cc).alignment=al('center')
    E.blank()
    # résumé scénario
    E.hdr(['Synthèse du scénario','Valeur','','','','','','','','Lecture'],h=16)
    def srow(label,formula,fmt,nm_,note=''):
        r=E.row([label,formula,None,None,None,None,None,None,None,note],fmts=[None,fmt],key=(1,)); bk.name(f'LST_{code}_{nm_}',E.ws,f'$C${r}'); return r
    srow('Dates plafonnées ou suspendues',f'=SUM(C{ER[(code,"gated")]}:J{ER[(code,"gated")]})',D,'gated')
    srow('File d\'attente maximale, % AN',f'=MAX(C{ER[(code,"file")]}:J{ER[(code,"file")]})',PCT,'file')
    srow('File résiduelle à la date 8, % AN',f'=J{ER[(code,"file")]}',PCT,'file8','Indicateur unique du Conducting Officer (< 5 % vert, 5 à 15 % orange, 15 à 30 % élevé, > 30 % sévère)')
    srow('Première date de suspension',f'=IFERROR(INDEX(C{ER[(code,"date")]}:J{ER[(code,"date")]},MATCH(1,C{ER[(code,"susp")]}:J{ER[(code,"susp")]},0)),"-")',DATE,'susp')
    srow('Poche liquide minimale, % AN',f'=MIN(C{ER[(code,"poche")]}:J{ER[(code,"poche")]})',PCT,'poche')
    srow('Couverture à douze mois minimale',f'=MIN(C{ER[(code,"couv")]}:J{ER[(code,"couv")]})',DEC,'couv')
    srow('Rachats obtenus des fonds evergreen, cumul (EUR)',f'=SUM(C{ER[(code,"rach_evg")]}:J{ER[(code,"rach_evg")]})',EUR,'evg')
    srow('Cessions secondaires, cumul (EUR)',f'=SUM(C{ER[(code,"cession")]}:J{ER[(code,"cession")]})',EUR,'cession')
    srow('Décotes réalisées, cumul (EUR)',f'=SUM(C{ER[(code,"perte_hy")]}:J{ER[(code,"perte_hy")]})+SUM(C{ER[(code,"perte_c")]}:J{ER[(code,"perte_c")]})',EUR,'pertes')
    srow('Défaut de liquidité, cumul (EUR)',f'=SUM(C{ER[(code,"defaut")]}:J{ER[(code,"defaut")]})',EUR,'defaut')
    srow('Actif net à la date 8 / actif net de départ',f'=J{ER[(code,"anfin")]}/AN_depart-1',PCT,'an8')
    srow('Première date de triple gel',f'=IFERROR(INDEX(C{ER[(code,"date")]}:J{ER[(code,"date")]},MATCH(1,C{ER[(code,"gel")]}:J{ER[(code,"gel")]},0)),"-")',DATE,'gel')
    srow('Sur-engagement maximal',f'=MAX(C{ER[(code,"sureng")]}:J{ER[(code,"sureng")]})',PCT0,'sureng')
    srow('Verdict LST du scénario',f'=IF(COUNTIF(C{ER[(code,"verdict")]}:J{ER[(code,"verdict")]},"Rouge")>0,"Rouge",IF(COUNTIF(C{ER[(code,"verdict")]}:J{ER[(code,"verdict")]},"Orange")>0,"Orange","Vert"))',None,'verdict','Pire verdict des huit dates')
    verdict_cf(E.ws,f'C{E.r-1}')
    E.blank(2)
# resolve deferred row references {key} inside each scenario block
import re
for code,_,_ in SCEN:
    keys={k:v for (c,k),v in ER.items() if c==code}
    rows=[v for (c,k),v in ER.items() if c==code]
    for rr in range(min(rows),max(rows)+1):
        for cc in range(3,3+T):
            c=E.ws.cell(rr,cc); v=c.value
            if isinstance(v,str) and '{' in v:
                c.value=re.sub(r'\{([a-z0-9_]+)\}',lambda m:str(keys[m.group(1)]),v)
# ======================================================================= LST REVERSE
RV=Sheet(bk,'LST Reverse','LST - TESTS INVERSÉS : TAUX DE RACHAT PERSISTANT, AVEC ET SANS PLAFONNEMENT','Moteur compact en % de l\'actif net de départ (actif net supposé constant, effet marché neutralisé) : un même taux de rachat demandé à chaque date, collecte nulle, flux de ressources du profil choisi dans l\'onglet Paramètres (par défaut le régime normal, pour isoler l\'effet des rachats ; 3 pour cumuler avec les flux S3). Cascade identique au moteur principal : poche liquide, rachats evergreen pour reconstituer la poche d\'exécution, cession secondaire décotée au-delà. Avec plafonnement : suspension après le nombre maximal de reports. L\'écart entre les deux points de rupture mesure la valeur du plafonnement.',
        [2,40,11,11,11,11,11,11,11,11,3,12,12,12,12,34],ncols=15)
rates=[0.025,0.05,0.075,0.10,0.125,0.15,0.20,0.25,0.30]
RV.sec('A. RESSOURCES PAR DATE EN % DE L\'ACTIF NET (profil de flux choisi, hors rachats et hors collecte)')
RV.hdr(['Flux']+[f'Date {t+1}' for t in range(T)]+['','','','','Source'],h=16)
rv_cap=RV.row(['Capacité de rachat des fonds evergreen']+[f'=({LP[2]["w"]}*{LP[2]["gate"]}+{LP[3]["w"]}*{LP[3]["gate"]})*2*(1-CHOOSE(Reverse_flux,{CAL[("pairs","S1")][t]},{CAL[("pairs","S2")][t]},{CAL[("pairs","S3")][t]},{CAL[("pairs","S4")][t]}))' for t in range(T)]+[None,None,None,None,'Gates trimestriels × 2, part des pairs déduite'],fmts=[None]+[PCT]*T)
rv_dist=RV.row(['Distributions des fonds cibles']+[f'=(({LP[0]["w"]}*{LP[0]["dist"]}+{LP[1]["w"]}*{LP[1]["dist"]})+({LP[2]["w"]}*{LP[2]["dist"]}+{LP[3]["w"]}*{LP[3]["dist"]}))/2*CHOOSE(Reverse_flux,{CAL[("dist","S1")][t]},{CAL[("dist","S2")][t]},{CAL[("dist","S3")][t]},{CAL[("dist","S4")][t]})' for t in range(T)]+[None,None,None,None,'Multiplicateur S3'],fmts=[None]+[PCT]*T)
rv_app=RV.row(['Appels des fonds fermés']+[f'=MIN(MAX(0,NA_total/AN_depart-SUM($C${RV.r}:{PC[t-1]}${RV.r})),NA_total/AN_depart*Appels_base*CHOOSE(Reverse_flux,{CAL[("appels","S1")][t]},{CAL[("appels","S2")][t]},{CAL[("appels","S3")][t]},{CAL[("appels","S4")][t]}))' if t>0 else f'=MIN(NA_total/AN_depart,NA_total/AN_depart*Appels_base*CHOOSE(Reverse_flux,{CAL[("appels","S1")][0]},{CAL[("appels","S2")][0]},{CAL[("appels","S3")][0]},{CAL[("appels","S4")][0]}))' for t in range(T)]+[None,None,None,None,'Jusqu\'à épuisement du non-appelé'],fmts=[None]+[PCT]*T)
rv_fr=RV.row(['Frais du Fonds']+['=Frais_an/2' for t in range(T)]+[None,None,None,None,''],fmts=[None]+[PCT]*T)
rv_net=RV.row(['Ressources nettes hors rachats']+[f'={PC[t]}{rv_dist}-{PC[t]}{rv_app}-{PC[t]}{rv_fr}' for t in range(T)]+[None,None,None,None,'Négatif : les appels dépassent les distributions'],fmts=[None]+[PCT]*T,key=tuple(range(1,T+1)))
RV.blank()
hc3="CHOOSE(Reverse_flux,"+LP[0]['hc']['S2']+","+LP[0]['hc']['S2']+","+LP[0]['hc']['S3']+","+LP[0]['hc']['S4']+")"
def reverse_block(gated):
    RV.sec(('B. AVEC PLAFONNEMENT' if gated else 'C. SANS PLAFONNEMENT')+('   (gate 5 % par date, reports automatiques, suspension après le nombre maximal de reports)' if gated else '   (gate levé par la Société de Gestion : toutes les demandes sont honorées)'))
    RV.hdr(['Taux par date et variable']+[f'Date {t+1}' for t in range(T)]+['','File à la date 8','Poche min','Cessions à la date 4','Cessions à la date 8','Verdict'],h=28)
    first=RV.r; rows=[]
    for rt_ in rates:
        r0=RV.r
        RV.row([rt_]+[None]*T,fmts=[PCT],inputs=(0,))
        rate=f'$B${r0}'
        rdem=RV.r; RV.row(['   Demande totale (nouvelle + reportée)']+[(f'={rate}' if t==0 else f'={rate}+{PC[t-1]}{rdem+2}') for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        rexe=RV.r
        if gated: RV.row(['   Exécuté']+[f'=IF(AND({rate}>Gate+0.0001,{t+1}>Reports_max,Suspension_auto=1),0,MIN({PC[t]}{rdem},Gate))' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        else: RV.row(['   Exécuté']+[f'={PC[t]}{rdem}' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        rfile=RV.r; RV.row(['   File reportée']+[f'={PC[t]}{rdem}-{PC[t]}{rexe}' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        rpav=RV.r; RV.row(['   Poche liquide avant appel aux fonds cibles']+[((f'=W_cash+W_HY' if t==0 else f'={PC[t-1]}{rpav+3}')+f'-{PC[t]}{rexe}+{PC[t]}{rv_net}') for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        revg=RV.r; RV.row(['   Rachats obtenus des fonds evergreen']+[f'=MIN({PC[t]}{rv_cap},MAX(0,Poche_min_exec-{PC[t]}{rpav}))' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        rces=RV.r; RV.row(['   Cession secondaire (valeur cédée)']+[f'=IF(Cession_autorisee=1,MAX(0,Cash_min-({PC[t]}{rpav}+{PC[t]}{revg}))/(1-{hc3}),0)' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        rpf=RV.r; RV.row(['   Poche liquide de fin']+[f'={PC[t]}{rpav}+{PC[t]}{revg}+{PC[t]}{rces}*(1-{hc3})' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        rcum=RV.r; RV.row(['   Cessions cumulées']+[f'=SUM($C${rces}:{PC[t]}{rces})' for t in range(T)],fmts=[None]+[PCT]*T,bold_first=False)
        # summary on the rate row
        ws=RV.ws
        for col,f,fmt in [(12,f'=J{rfile}',PCT),(13,f'=MIN(C{rpf}:J{rpf})',PCT),(14,f'=F{rcum}',PCT),(15,f'=J{rcum}',PCT)]:
            c=ws.cell(r0,col,f); c.number_format=fmt; c.font=F(True); c.alignment=al('right'); c.border=B_BOT
        if gated: v=f'=IF(AND({rate}>Gate+0.0001,Reports_max<8,Suspension_auto=1),"Rouge : suspension à la date "&(Reports_max+1),IF(N{r0}>0.0001,"Rouge : cessions décotées",IF(M{r0}<S_poche_rouge,"Rouge : poche sous 5 %",IF(OR(M{r0}<S_poche_orange,{rate}>Gate+0.0001),"Orange","Vert"))))'
        else: v=f'=IF(O{r0}>({LP[0]["w"]}+{LP[1]["w"]}),"Rouge : cession impossible",IF(N{r0}>0.0001,"Rouge : cessions décotées",IF(M{r0}<S_poche_rouge,"Rouge : poche sous 5 %",IF(M{r0}<S_poche_orange,"Orange","Vert"))))'
        c=ws.cell(r0,16,v); c.font=F(True); c.border=B_BOT; c.alignment=al('left')
        rows.append(r0)
    verdict_cf(RV.ws,f'P{rows[0]}:P{rows[-1]}')
    return rows
rows_g=reverse_block(True)
rg=RV.row(['Point de rupture : premier taux imposant des cessions décotées dans les quatre premières dates',None]+[None]*(T-1)+[None,None,None,None,None,'Avec le gate, la file absorbe l\'excès de demande ; la cession n\'intervient qu\'à cause des appels et des frais'],key=(1,))
RV.ws.cell(rg,3).value=ArrayFormula(f'C{rg}',f'=MIN(IF(N{rows_g[0]}:N{rows_g[-1]}>0.0001,B{rows_g[0]}:B{rows_g[-1]},9))'); RV.ws.cell(rg,3).number_format=PCT; bk.name('Rupture_gated',RV.ws,f'$C${rg}')
rgs=RV.row(['Premier taux conduisant à la suspension',None]+[None]*(T-1)+[None,None,None,None,None,'Par construction : premier taux supérieur au plafond'],key=(1,))
RV.ws.cell(rgs,3).value=ArrayFormula(f'C{rgs}',f'=MIN(IF(B{rows_g[0]}:B{rows_g[-1]}>Gate+0.0001,B{rows_g[0]}:B{rows_g[-1]},9))'); RV.ws.cell(rgs,3).number_format=PCT; bk.name('Rupture_suspension',RV.ws,f'$C${rgs}')
RV.blank()
rows_u=reverse_block(False)
ru=RV.row(['Point de rupture : premier taux imposant des cessions décotées dans les quatre premières dates',None]+[None]*(T-1)+[None,None,None,None,None,'Sans gate, la poche et les flux S3 s\'épuisent dès les premières dates'],key=(1,))
RV.ws.cell(ru,3).value=ArrayFormula(f'C{ru}',f'=MIN(IF(N{rows_u[0]}:N{rows_u[-1]}>0.0001,B{rows_u[0]}:B{rows_u[-1]},9))'); RV.ws.cell(ru,3).number_format=PCT; bk.name('Rupture_ungated',RV.ws,f'$C${ru}')
rui=RV.row(['Premier taux rendant la cession impossible à la date 8 (fonds fermés épuisés)',None]+[None]*(T-1)+[None,None,None,None,None,''],key=(1,))
RV.ws.cell(rui,3).value=ArrayFormula(f'C{rui}',f'=MIN(IF(O{rows_u[0]}:O{rows_u[-1]}>({LP[0]["w"]}+{LP[1]["w"]}),B{rows_u[0]}:B{rows_u[-1]},9))'); RV.ws.cell(rui,3).number_format=PCT; bk.name('Rupture_impossible',RV.ws,f'$C${rui}')
rvv=RV.row(['Valeur du plafonnement : écart des points de rupture (points de % AN par date)','=Rupture_gated-Rupture_ungated']+[None]*(T-1)+[None,None,None,None,None,'Nul si la poche s\'épuise avant que le gate ne joue : la protection du gate se lit alors sur les cessions'],fmts=[None,PCT],key=(1,)); bk.name('Valeur_gate',RV.ws,f'$C${rvv}')
rv10=RV.row(['Cessions cumulées à la date 8 au taux de 10 % par date : avec plafonnement, puis sans',f'=O{rows_g[3]}',f'=O{rows_u[3]}']+[None]*(T-2)+[None,None,None,None,None,'Le gate transforme des cessions forcées en file d\'attente puis en suspension'],fmts=[None,PCT,PCT],key=(1,2)); bk.name('Cession_gate_10',RV.ws,f'$C${rv10}'); bk.name('Cession_nogate_10',RV.ws,f'$D${rv10}')
RV.blank()
RV.sec('D. APPELS DE FONDS ACCÉLÉRÉS   non-appelé entièrement appelé sur douze mois, ressources S3, rachats au plafond')
RV.hdr(['Non-appelé de départ (% AN)','Appels sur 12 mois (% AN)','Ressources 12 mois hors appels (% AN)','Rachats au plafond 12 mois (% AN)','Cession nécessaire (% AN)','Verdict','','','','','','','','',''],h=40)
rc0=RV.r
for na in [0.05,0.075,0.10,0.125,0.15,0.20,0.25,0.30]:
    r=RV.r
    RV.row([na,f'=B{r}',f'=W_cash-Cash_min+W_HY*(1-{LP[4]["hc"]["S2"]})+C{rv_cap}+D{rv_cap}+C{rv_dist}+D{rv_dist}-C{rv_fr}-D{rv_fr}',f'=2*Gate',f'=MAX(0,C{r}+E{r}-D{r})/(1-{hc3})',f'=IF(F{r}>0.0001,"Rouge : cession forcée",IF(D{r}-C{r}-E{r}<S_poche_rouge,"Orange : poche épuisée","Vert"))'],fmts=[PCT,PCT,PCT,PCT,PCT],inputs=(0,))
rcl=RV.r-1
rca=RV.row(['Point de rupture : premier niveau de non-appelé imposant une cession',None]+[None]*(T-1)+[None,None,None,None,None,'Plafond de sur-engagement : 130 % (art. 23)'],key=(1,))
RV.ws.cell(rca,3).value=ArrayFormula(f'C{rca}',f'=MIN(IF(F{rc0}:F{rcl}>0.0001,B{rc0}:B{rcl},9))'); RV.ws.cell(rca,3).number_format=PCT; bk.name('Rupture_appels',RV.ws,f'$C${rca}')
verdict_cf(RV.ws,f'G{rc0}:G{rcl}')
# ======================================================================= CONCENTRATION
C=Sheet(bk,'Concentration','CONCENTRATION DES PORTEURS ET CHOC DU PREMIER PORTEUR','Registre non constitué : répartition hypothétique à remplacer par le registre réel (cellules cyan). Les ratios alimentent la question 118 du reporting Annex IV et le choc « porteur 1 » des scénarios S3 et S4.',
        [2,10,22,14,14,14,14,40],ncols=7)
C.hdr(['Rang','Porteur (anonymisé)','Part de l\'AN','Encours (EUR)','Cumul','Alerte','Commentaire'],h=16)
seed=[0.15,0.08,0.05,0.04,0.03,0.03,0.02,0.02,0.02,0.02]
rc0=C.r
for i,sh in enumerate(seed):
    r=C.r
    C.row([i+1,f'POR-{i+1:03d}',sh,f'=D{r}*AN_depart',f'=SUM(D{rc0}:D{r})',f'=IF(D{r}>0.2,"Rouge : > 20 %",IF(D{r}>0.1,"Orange : 10 à 20 %","Vert"))','Hypothèse' if i else 'Hypothèse : premier porteur = paramètre Top1'],fmts=[D,None,PCT,EUR,PCT],inputs=(2,),aligns=[None,None,None,None,None,'left'])
    if i==0: C.ws.cell(r,4).value='=Top1'
rcl=C.r-1
C.row([None,'Autres porteurs',f'=1-SUM(D{rc0}:D{rcl})',f'=D{C.r}*AN_depart',1,'',''],fmts=[None,None,PCT,EUR,PCT])
C.blank()
C.sec('RATIOS ET CHOC')
C.hdr(['Mesure','','Valeur','','','Lecture',''],h=16)
C.row(['Part du premier porteur',None,f'=D{rc0}',None,None,'Vert ≤ 10 % ; Orange 10 à 20 % ; Rouge > 20 %'],fmts=[None,None,PCT],key=(2,))
C.row(['Part des trois premiers',None,f'=SUM(D{rc0}:D{rc0+2})',None,None,''],fmts=[None,None,PCT])
C.row(['Part des cinq premiers',None,f'=SUM(D{rc0}:D{rc0+4})',None,None,'Alimente le scénario de concentration du rapport de risque'],fmts=[None,None,PCT])
C.row(['Part des dix premiers',None,f'=SUM(D{rc0}:D{rcl})',None,None,''],fmts=[None,None,PCT])
C.row(['Indice de Herfindahl (dix premiers et solde)',None,f'=SUMPRODUCT(D{rc0}:D{rcl+1},D{rc0}:D{rcl+1})*10000',None,None,'> 2 500 : concentré'],fmts=[None,None,'#,##0'])
rq=C.row(['Dates de rachat consommées par la sortie intégrale du premier porteur',None,f'=D{rc0}/Gate',None,None,'Part / plafond par date'],fmts=[None,None,'0.0'],key=(2,)); bk.name('Top1_dates',C.ws,f'$D${rq}')
C.row(['Verdict du choc porteur 1',None,f'=IF(D{rq}<=2,"Vert : absorbé en deux dates",IF(D{rq}<=Reports_max,"Orange : cascade sur plusieurs dates","Rouge : dépasse le nombre maximal de reports"))',None,None,'Au-delà du nombre maximal de reports, la suspension devient probable'],key=(2,)); bk.name('Top1_verdict',C.ws,f'$D${C.r-1}')
verdict_cf(C.ws,f'D{C.r-1}')
# ======================================================================= INTÉGRÉ
I=Sheet(bk,'Intégré','LST × MST - MATRICE DE STRESS CONJOINTE ET INDICATEURS DE PILOTAGE','Chaque scénario LST est déjà calculé avec le choc de marché du même scénario (le moteur lit les rendements de l\'onglet MST Scénarios) : le plafond en euros se contracte quand la VL baisse et que les demandes montent. Cette page rassemble les sorties des deux volets et les seuils d\'escalade.',
        [2,34,14,14,14,14,40],ncols=6)
I.sec('A. MATRICE CONJOINTE')
I.hdr(['Indicateur','S1 Base','S2 Plausible','S3 Sévère','S4 Extrême','Lecture'],h=16)
def irow(label,fn,fmt,note='',key=False,center=False):
    r=I.row([label]+[fn(c) for c,_,_ in SCEN]+[note],fmts=[None]+[fmt]*4,key=(1,2,3,4) if key else (),aligns=[None]+(['center']*4 if center else [None]*4)); return r
irow('Creux de VL hors flux (MST)',lambda c:f'=MST_{c}_creux',PCT,'Effet marché seul, transmission des chocs cotés incluse')
irow('Semestre de retour au niveau de départ (MST)',lambda c:f'=MST_{c}_reprise',D)
irow('Dates plafonnées ou suspendues (LST)',lambda c:f'=LST_{c}_gated',D)
irow('File d\'attente maximale, % AN (LST)',lambda c:f'=LST_{c}_file',PCT)
irow('File résiduelle à la date 8, % AN',lambda c:f'=LST_{c}_file8',PCT,'Indicateur unique du Conducting Officer',key=True)
irow('Première suspension',lambda c:f'=LST_{c}_susp',DATE)
irow('Poche liquide minimale, % AN',lambda c:f'=LST_{c}_poche',PCT)
irow('Couverture à douze mois minimale',lambda c:f'=LST_{c}_couv',DEC)
irow('Cessions secondaires cumulées (EUR)',lambda c:f'=LST_{c}_cession',EUR)
irow('Décotes réalisées cumulées (EUR)',lambda c:f'=LST_{c}_pertes',EUR)
irow('Défaut de liquidité cumulé (EUR)',lambda c:f'=LST_{c}_defaut',EUR,'Doit rester nul')
irow('Actif net à la date 8 / départ (flux et marché)',lambda c:f'=LST_{c}_an8',PCT)
irow('Première date de triple gel',lambda c:f'=LST_{c}_gel',DATE)
irow('Verdict MST',lambda c:f'=MST_{c}_verdict',None,center=True)
irow('Verdict LST',lambda c:f'=LST_{c}_verdict',None,center=True)
rj=irow('Verdict conjoint',lambda c:f'=IF(OR(MST_{c}_verdict="Rouge",LST_{c}_verdict="Rouge"),"Rouge",IF(OR(MST_{c}_verdict="Orange",LST_{c}_verdict="Orange"),"Orange","Vert"))',None,'Pire des deux volets',key=True,center=True)
for j,(c,_,_) in enumerate(SCEN): bk.name(f'Joint_{c}',I.ws,f'${L(3+j)}${rj}')
verdict_cf(I.ws,f'C{rj-2}:F{rj}')
I.blank()
I.sec('B. SEUILS D\'ESCALADE   (Conducting Officer, Comité des risques)')
I.hdr(['Déclencheur','Seuil','Valeur observée (S3)','État','','Action'],h=16)
I.row(['File résiduelle à la date 8','=0.3','=LST_S3_file8','=IF(D'+str(I.r)+'>C'+str(I.r)+',"Rouge","Vert")',None,'Au-dessus de 30 % : réunion du Comité avant la prochaine date de rachat ; préparer la suspension'],fmts=[None,PCT,PCT])
I.row(['Satisfaction des demandes à la date 1',"=0.8",f"=IFERROR(C{ER[('S3','exec')]}/C{ER[('S3','dem_tot')]},1)",'=IF(D'+str(I.r)+'<C'+str(I.r)+',"Orange","Vert")',None,'En dessous de 80 % : information du Comité et des distributeurs'],fmts=[None,PCT0,PCT0])
I.row(['Poche liquide minimale','=S_poche_orange','=LST_S3_poche','=IF(D'+str(I.r)+'<C'+str(I.r)+',"Orange","Vert")',None,'Sous 15 % : plan de reconstitution à 30 jours ; suspension des souscriptions nettes dans les fonds fermés'],fmts=[None,PCT,PCT])
I.row(['Creux de VL','=S_mst_orange','=-MST_S3_creux','=IF(D'+str(I.r)+'>C'+str(I.r)+',"Orange","Vert")',None,'Au-delà de 10 % : MST ad hoc et revue des VL des fonds cibles'],fmts=[None,PCT0,PCT])
I.row(['Premier porteur','=0.2','=Top1','=IF(D'+str(I.r)+'>C'+str(I.r)+',"Rouge",IF(D'+str(I.r)+'>0.1,"Orange","Vert"))',None,'Au-dessus de 10 % : protocole de préavis renforcé avec le porteur'],fmts=[None,PCT0,PCT0])
I.row(['Sur-engagement maximal','=Surengagement_max','=LST_S3_sureng','=IF(D'+str(I.r)+'>C'+str(I.r)+',"Rouge","Vert")',None,'Dépassement passif toléré par l\'art. 23, mais aucun nouvel engagement'],fmts=[None,PCT0,PCT0])
verdict_cf(I.ws,f'E{I.r-6}:E{I.r-1}')
# ======================================================================= SOURCES
S=Sheet(bk,'Sources','SOURCES ET TRAÇABILITÉ','Chaque paramètre de l\'onglet Paramètres cite sa source. Cette page récapitule les documents et la règle de calibration des séries.',[2,46,16,70],ncols=3)
S.hdr(['Document','Date','Usage dans le classeur'],h=16)
for d in [('Prospectus et Règlement Openstone Infraworld, projet',dt.date(2026,9,30),'Ancres de passif (art. 6, 7.1, 7.2.1 à 7.2.8, 22, 23), actifs liquides 15 %, sur-engagement 130 %'),
          ('Documents d\'informations clés PRIIPs A1c et A1d',dt.date(2026,9,29),'Absence de commission de rachat, SRI 4, TRI cibles et coûts'),
          ('Classeur KID PRIIPs Openstone Infraworld v2.3',dt.date(2026,9,22),'TRI nets cibles par ligne, frais du Fonds, VEV privée et cotée, autocorrélation et volatilités Preqin (bloc Geltner)'),
          ('Note de liquidité Openstone Infraworld',dt.date(2026,9,28),'Seuils de verdict, nombre maximal de reports, poche d\'exécution'),
          ('Synthèse commerciale Openstone Infraworld v5.4',dt.date(2026,9,30),'Allocation cible, objectif de collecte 50 M€'),
          ('PPM Macquarie Alliance Partners Infrastructure Fund II',dt.date(2026,4,30),'Fonds fermé : aucun rachat, appels à 10 JO, cession agréée'),
          ('PPM Macquarie Specialized Infrastructure Global 3 (USD)',dt.date(2026,2,28),'Retrait annuel après lock-up de 3 ans, traité comme fermé sur l\'horizon'),
          ('Prospectus Partners Group Evergreen SICAV, supplément 2',dt.date(2026,6,9),'Gate 5 % par trimestre, préavis 45 JO, Special Dealing'),
          ('Prospectus AEFS SICAV, annexe 7 Ares Global Infrastructure ELTIF',dt.date(2026,8,11),'Gate 5 % net par trimestre, limite ELTIF, déduction 5 % avant 24 mois'),
          ('Classeur des cours des proxies cotés (Yahoo Finance)',dt.date(2026,10,6),'Composites C1 à C4 et EUR/USD : écarts-types semestriels et pires semestres (MST Données)'),
          ('ESMA 34-39-882, orientations sur les simulations de crise de liquidité (2020)',dt.date(2020,7,16),'Approche bilatérale actif-passif, tests inversés, concentration des porteurs'),
          ('Règlement délégué (UE) 231/2013, art. 46 à 49 ; directive 2011/61/UE art. 16',dt.date(2013,7,22),'Obligation de simulations de crise et de cohérence profil de liquidité / politique de rachat')]:
    S.row(list(d),fmts=[None,DATE,None],h=26)
S.blank()
S.note('Règle de calibration des séries : rendements mensuels simples, composites équipondérés des titres disponibles (Brookfield exclu, INFR à partir de 2016), mois en cours exclu, rendements aberrants (|ln| > 0,6) ignorés. Écart-type semestriel = écart-type mensuel × √6 ; pires semestres sur fenêtres glissantes de six mois.',h=30)
# ======================================================================= TABLEAU DE BORD
Dsh=Sheet(bk,'Tableau de bord','TABLEAU DE BORD LST ET MST - OPENSTONE INFRAWORLD, COMPARTIMENT I','="Date d\'arrêté : "&TEXT(DateArrete,"jj/mm/aaaa")',[2,44,15,15,15,15,44],ncols=6)
Dsh.ws.cell(3,2).value=f'="Projection à partir de la fin du blocage ("&DAY(Date_depart)&"/"&RIGHT("0"&MONTH(Date_depart),2)&"/"&YEAR(Date_depart)&"), huit dates de rachat semestrielles, actif net de départ "&ROUND(AN_depart/1000000,0)&"{NB}M€. Phase : "&Phase&". Les cellules sont des formules vers les onglets de calcul."'
Dsh.sec('A. INDICATEURS DE TÊTE')
Dsh.hdr(['Indicateur','Valeur','','','','Lecture'],h=16)
def drow(label,f,fmt,note,key=True):
    r=Dsh.row([label,f,None,None,None,note],fmts=[None,fmt],key=(1,) if key else ()); return r
drow('Écart actif - passif en régime normal (jours)','=Ecart_TTL_N',D,'WATTL actif moins LTTL moyen ; porté par la poche liquide et le plafonnement')
drow('Écart actif - passif en S3 (jours)','=Ecart_TTL_S3',D,'Justification structurelle du gate de 5 %')
drow('Actifs liquides à 30 jours, normal / S3','=HQLA_N',PCT,'Trésorerie et fonds HY après haircut')
Dsh.ws.cell(Dsh.r-1,4).value='=HQLA_S3'; Dsh.ws.cell(Dsh.r-1,4).number_format=PCT; Dsh.ws.cell(Dsh.r-1,4).font=F(True); Dsh.ws.cell(Dsh.r-1,4).fill=fill(LIME)
drow('Capacité maximale de sortie sur douze mois (2 × gate)','=2*Gate',PCT0,'Prospectus art. 7.2.6')
drow('Part du premier porteur et dates de gate consommées','=Top1',PCT,'Choc porteur 1')
Dsh.ws.cell(Dsh.r-1,4).value='=Top1_dates'; Dsh.ws.cell(Dsh.r-1,4).number_format='0.0'; Dsh.ws.cell(Dsh.r-1,4).font=F(True); Dsh.ws.cell(Dsh.r-1,4).fill=fill(LIME)
drow('Point de rupture avec plafonnement (taux par date)','=Rupture_gated',PCT,'Premier taux imposant des cessions décotées malgré le gate ; suspension dès le premier taux supérieur au gate')
Dsh.ws.cell(Dsh.r-1,4).value='=Rupture_suspension'; Dsh.ws.cell(Dsh.r-1,4).number_format=PCT; Dsh.ws.cell(Dsh.r-1,4).font=F(True); Dsh.ws.cell(Dsh.r-1,4).fill=fill(LIME)
drow('Point de rupture sans plafonnement (taux par date)','=Rupture_ungated',PCT,'Premier taux imposant des cessions décotées ; cession impossible au-delà de la valeur en colonne D')
Dsh.ws.cell(Dsh.r-1,4).value='=Rupture_impossible'; Dsh.ws.cell(Dsh.r-1,4).number_format=PCT; Dsh.ws.cell(Dsh.r-1,4).font=F(True); Dsh.ws.cell(Dsh.r-1,4).fill=fill(LIME)
drow('Point de rupture des appels de fonds (non-appelé, % AN)','=Rupture_appels',PCT,'Non-appelé au-delà duquel une cession s\'impose en douze mois')
drow('Cessions à la date 8 pour 10 % de rachats par date : avec puis sans plafonnement','=Cession_gate_10',PCT,'Valeur du plafonnement : la file remplace la vente forcée')
Dsh.ws.cell(Dsh.r-1,4).value='=Cession_nogate_10'; Dsh.ws.cell(Dsh.r-1,4).number_format=PCT; Dsh.ws.cell(Dsh.r-1,4).font=F(True); Dsh.ws.cell(Dsh.r-1,4).fill=fill(LIME)
drow('Première date de triple gel en S3','=LST_S3_gel',DATE,'Collecte, distributions et capacité evergreen gelées ensemble')
Dsh.blank()
Dsh.sec('B. VERDICTS PAR SCÉNARIO')
Dsh.hdr(['Scénario','MST (creux de VL)','LST (file à la date 8)','Suspension','Verdict conjoint','Lecture'],h=16)
for c,lab,desc in SCEN:
    r=Dsh.row([f'{c} {lab}',f'=MST_{c}_creux',f'=LST_{c}_file8',f'=LST_{c}_susp',f'=Joint_{c}',desc],fmts=[None,PCT,PCT,DATE,None],h=26,key=(4,),aligns=[None,None,None,None,'center'])
verdict_cf(Dsh.ws,f'F{Dsh.r-4}:F{Dsh.r-1}')
Dsh.blank()
Dsh.sec('C. ACTIONS RECOMMANDÉES   déduites des seuils de l\'onglet Intégré')
acts=[
 '=IF(Phase="Montée en charge","Phase de montée en charge : le risque dominant est la concentration des porteurs, pas la cascade de plafonnement ; LST mensuel jusqu\'à la troisième année.","Régime de croisière : LST trimestriel avant chaque date de centralisation.")',
 '=IF(Top1>0.1,"Premier porteur au-dessus de 10 % : convenir d\'un protocole de préavis renforcé et simuler sa sortie avant chaque date.","Concentration des porteurs sous contrôle.")',
 '=IF(LST_S3_file8>0.3,"S3 : file résiduelle au-dessus de 30 % à la date 8 ; préparer la suspension et l\'information de l\'AMF.",IF(LST_S3_file8>0.15,"S3 : file résiduelle élevée ; pré-positionner la Prolongation du Délai de Préavis (art. 7.2.7) et la communication aux distributeurs.","S3 : file résorbable sous les outils du prospectus."))',
 '=IF(LST_S2_poche<S_poche_orange,"S2 : poche liquide sous 15 % ; relever la trésorerie cible ou ralentir les engagements fermés.","S2 : poche liquide conforme à l\'article 3.")',
 '=IF(-MST_S3_creux>S_mst_rouge,"MST S3 : creux de VL au-delà de 25 % ; revoir la transmission des chocs et l\'exposition dollar non couverte (51 %).","MST S3 : creux de VL contenu ; surveiller le dollar (aucune couverture possible, art. 22).")',
 '=IF(Rupture_appels<=NA_total/AN_depart+0.0001,"Appels : le non-appelé de départ dépasse le point de rupture ; séquencer les engagements fermés.","Appels : le non-appelé de départ reste sous le point de rupture ("&ROUND(Rupture_appels*100,0)&" % de l\'AN).")',
 '="Reporting : les profils ALP et LLP alimentent les questions 178 et 186 de l\'Annex IV ; la concentration la question 118."']
for a in acts: Dsh.row([a],merge=[(2,7)],h=30,bold_first=False)
Dsh.blank()
Dsh.sec('D. HYPOTHÈSES CLÉS À VALIDER EN COMITÉ')
for a in ['Taux de rachat par scénario et collecte (Paramètres G) : aucun historique, hypothèses de la Fonction Risques.',
          'Transmission des chocs cotés aux VL privées : 0,5 par défaut, cohérente avec le rapport VEV privée / VEV cotée du DIC (MST Données).',
          'Pression des pairs sur les gates de PG NGI et d\'Ares AGI : 1 en S3 et S4 aux deux premières dates (aucun rachat récupérable).',
          'Décotes de cession secondaire : 15 % (S2), 25 % (S3), 35 % (S4) sur les fonds fermés ; 1 à 5 % sur le fonds HY.',
          'Registre des porteurs hypothétique : à remplacer avant la première VL.']:
    Dsh.row([a],merge=[(2,7)],h=18,bold_first=False)
# ----------------------------------------------------------------- order & save
order=['Tableau de bord','Paramètres','ALP','LLP','LST Moteur','LST Reverse','Concentration','MST Données','MST Scénarios','Intégré','Sources']
wb._sheets=[wb[n] for n in order]
for w in wb.worksheets:
    w.sheet_view.showGridLines=False; assert w.freeze_panes is None
    w.page_setup.orientation='landscape'; w.page_setup.fitToWidth=1; w.page_setup.fitToHeight=0; w.sheet_properties.pageSetUpPr.fitToPage=True
out='LST_MST_Openstone_Infraworld.xlsx'; wb.save(out); print('saved',out)
