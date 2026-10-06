# -*- coding: utf-8 -*-
from lib import *
from params import name

def build_limites(wb):
    ws=wb.create_sheet('1 Limites')
    S=Sheet(ws,'RÉFÉRENTIEL DES LIMITES - OPENSTONE INFRAWORLD, COMPARTIMENT I',
            'Limites du prospectus (projet du 30/09/2026), seuils internes de la note de liquidité et points de contrôle. Statut à la date d\'arrêté : le Fonds n\'est pas constitué.',
            {'A':2,'B':8,'C':40,'D':12,'E':12,'F':12,'G':22,'H':30,'I':52})
    S.hdr(['Réf.','Limite','Type','Limite','Alerte','Statut','Base et contrôle','Source et commentaire'])
    L=[]
    def row(ref,lab,typ,lim,al_,stat,base,src,fmt='0 %'):
        r=S.row([ref,lab,typ,lim,al_,stat,base,src],fmts=[None,None,None,fmt,fmt],h=30); L.append(r); return r
    S.block('A. LIMITES DU PROSPECTUS - opposables')
    row('IR-01','Fonds Cibles principalement d\'infrastructure','Qualitative','-','-','Non testable','Chaque investissement','Art. 19 et 21.2 ; « principalement », sans seuil chiffré.',None)
    row('IR-03','Zone : Europe privilégiée, Amérique du Nord et Asie admises','Qualitative','-','-','Non testable','Chaque investissement','Art. 21.2 ; aucun seuil. Allocation cible : environ 80 % Amérique du Nord et Europe par transparence.',None)
    row('IR-04','Entreprises Cibles (co-investissements directs)','Dure','=Entreprises_max','=Entreprises_max-0.05','Non testable','Actif net, chaque VL','Art. 22 ; traitement d\'un dépassement passif non précisé.')
    row('IR-06','Parts de fonds (OPCVM, FIA, fonds étrangers)','Dure',1,'-','Non testable','Actif net, chaque VL','Art. 22 ; sous le plafond de sur-engagement.')
    row('IR-07','Instruments dérivés','Dure',0,'-','Paramétré','Pré-négociation','Art. 22 : néant. Exclut toute couverture de change ou de taux.')
    row('IR-08','Actifs immobiliers, acquisitions et cessions temporaires de titres','Dure',0,'-','Paramétré','Pré-négociation','Art. 22 : néant.')
    row('IR-10','Fonds gérés par la Société de Gestion','Procédure','-','-','Paramétré','Chaque investissement','Art. 21.2 ; politique de conflits d\'intérêts à appliquer avant investissement.',None)
    row('RL-01','Sur-engagement : engagements appelés et non appelés / actif net','Dure','=Surengagement_max','=S_sureng_orange','Non testable','Avant chaque engagement et chaque VL','Art. 23 ; dépassement passif toléré sans nouvel engagement.')
    row('RL-02','Plancher d\'actifs liquides','Dure','=Liq_min','=Liq_cible','Non testable','Actif, chaque VL','Art. 21.8 ; écrit comme obligation permanente, assiette en actif (registre n° 2 et 8).')
    row('RL-03','Engagements hors bilan (garanties, nantissements)','Dure','=Surengagement_max','-','Paramétré','Actif net','Art. 3.1.1 ; cumul avec RL-01 non précisé (registre n° 6).')
    row('RL-04','Emprunts d\'espèces','Dure','=Emprunt_max','-','Paramétré','Pré-négociation','Art. 22 : néant ; art. 3.1.1, 14.2 et 25.2.6 à nettoyer (registre n° 1).')
    row('RL-05','Détention par une personne physique','Dure','=Pers_phys_max','-','Non testable','Parts émises, souscription et transfert','Art. 5.1 et 8.1.')
    row('RL-06','Plafonnement des Rachats par Date de Centralisation','Dure','=Gate','-','Non testable','Actif net de la VL d\'exécution','Art. 7.2.6 ; nette des souscriptions ; 4 reports au plus.')
    row('RL-07','Contreparties : établissements OCDE notés au moins BBB+','Dure','-','-','Paramétré','À la mise en place','Art. 15 ; revue spécifique sous A-.',None)
    S.block('B. SEUILS INTERNES - note de liquidité du 24/09/2026, § 5, à valider en Comité des risques')
    row('KRI-01','Demande nette de rachat par date','Interne',0.10,0.05,'À valider','Actif net','Rouge au-delà de 10 % ; orange de 5 à 10 %.')
    row('KRI-02','File d\'attente de rachats reportés','Interne','=S_file_rouge',0.0001,'À valider','Actif net','Rouge au-delà de 25 % ; orange dès la première file.')
    row('KRI-03','Dates plafonnées consécutives','Interne','=Gate_dates',1,'À valider','Nombre','Au-delà de 4 : suspension examinée.','0')
    row('KRI-04','Couverture de liquidité à douze mois','Interne','=S_couv_rouge','=S_couv_vert','À valider','Ressources 12 mois / besoins 12 mois','Rouge en deçà de 1,0 ; orange de 1,0 à 1,5.','0.0')
    row('KRI-05','Levier, méthode de l\'engagement','Interne','=Levier_max','-','À valider','Actif net','Niveau maximal à réintroduire au prospectus.')
    row('KRI-06','Premier porteur','Interne',0.15,'=S_top1_orange','À valider','Actif net','Rouge au-delà de 15 %.')
    row('KRI-07','Cinq premiers porteurs','Interne',0.40,'=S_top5_orange','À valider','Actif net','Rouge au-delà de 40 %.')
    row('KRI-08','Concentration par gérant de fonds cibles','Interne','À arrêter','À arrêter','Seuil à arrêter','Actif net','Aucune limite au prospectus ; 51 % chez Macquarie à l\'allocation cible. Proposition : 40 %, alerte 35 %.',None)
    row('KRI-09','Concentration par fonds cible','Interne','À arrêter','À arrêter','Seuil à arrêter','Actif net','34 % sur MAPIF II à l\'allocation cible. Proposition : 35 %, alerte 30 %.',None)
    row('KRI-10','Exposition au dollar non couverte','Interne','À arrêter','À arrêter','Seuil à arrêter','Actif net','34 % à 68 % selon les véhicules retenus ; aucune couverture possible (IR-07).',None)
    row('KRI-11','Péremption des VL reçues des fonds cibles','Interne','À arrêter','À arrêter','Seuil à arrêter','Jours depuis la VL du fonds cible','VL trimestrielles et décalées pour MAPIF II et MSIG 3 ; mensuelles à 15-21 j.o. pour Ares et PG NGI.',None)
    S.block('C. SYNTHÈSE')
    r1=S.row(['','Limites et seuils recensés','=COUNTA(B7:B35)','','','','',''],fmts=[None,None,'0']); name(wb,ws,'Lim_total',f'D{r1}')
    r2=S.row(['','Seuils internes à arrêter','=COUNTIF(G7:G35,"Seuil à arrêter")','','','','',''],fmts=[None,None,'0']); name(wb,ws,'Lim_a_arreter',f'D{r2}')
    r3=S.row(['','Seuils internes à valider','=COUNTIF(G7:G35,"À valider")','','','','',''],fmts=[None,None,'0']); name(wb,ws,'Lim_a_valider',f'D{r3}')
    r4=S.row(['','Limites non testables avant lancement','=COUNTIF(G7:G35,"Non testable")','','','','',''],fmts=[None,None,'0']); name(wb,ws,'Lim_non_testables',f'D{r4}')
    r5=S.row(['','Franchissements de limite du prospectus','=COUNTIF(G7:G35,"Dépassée")','','','','',''],fmts=[None,None,'0']); name(wb,ws,'Lim_depassees',f'D{r5}')
    S.blank(); S.ncols=8
    S.note('Statuts : Paramétré (limite inscrite au dispositif de contrôle, sans position à tester) ; Non testable (nécessite une valeur liquidative ou un registre) ; À valider (seuil interne proposé par la Fonction Risques, non encore arrêté par le Comité) ; Seuil à arrêter (aucune valeur proposée dans les documents) ; Dépassée (franchissement constaté).')

def build_liquidite(wb):
    ws=wb.create_sheet('2 Liquidité')
    S=Sheet(ws,'PROFIL DE LIQUIDITÉ À L\'ALLOCATION CIBLE - actif, passif, couverture',
            'Délais de liquidation par ligne (classeur de stress tests du 24/09/2026), passif recalculé sur le préavis du prospectus du 30/09/2026 (91 jours). Format Annexe IV.',
            {'A':2,'B':44,'C':13,'D':13,'E':13,'F':13,'G':13,'H':48})
    S.hdr(['Ligne','Poids cible','Devise','Délai normal (j)','Délai stress (j)','Rachat au fonds cible','Source et hypothèse'])
    rows=[('MAPIF II - secondaires, fermé','=W_MAPIF','USD',1825,1825,'Néant','PPM avril 2026 : aucun retrait ; durée 10 ans + 2 + 2. Sortie par cession agréée par le GP.'),
          ('MSIG 3 - dette, fermé hybride','=W_MSIG','USD / EUR',1825,1825,'Annuel, run-off','PPM février 2026 : lock-up 3 ans puis retrait au 31/03 payé par écoulement, sans date.'),
          ('PG NGI - evergreen','=W_PG','USD',100,270,'5 % / trimestre','Supplément 2 : rachat mensuel, préavis 45 j.o., gate 5 % distributions comprises ; Special Dealing à 180 j.'),
          ('Ares AGI ELTIF - evergreen','=W_ARES','EUR',90,270,'5 % / trimestre','Annexe 7 : rachat trimestriel, cut-off 1er j.o. du dernier mois, gate 5 % net, limite ELTIF.'),
          ('Actifs liquides','=W_LIQ','EUR',3,7,'Quotidien','Art. 21.8 : monétaire, dépôts, OPC obligataires à VL quotidienne.')]
    first=S.r
    for t in rows: S.row(list(t),fmts=[None,'0 %',None,'0','0'],h=28,input_cols=(3,4))
    last=S.r-1
    r=S.row(['Délai moyen pondéré de liquidation de l\'actif (TTL actif)',f'=SUMPRODUCT(C{first}:C{last},E{first}:E{last})','jours',f'=SUMPRODUCT(C{first}:C{last},F{first}:F{last})','jours (stress)','',''],fmts=[None,'0',None,'0'],key_cols=(1,3))
    name(wb,ws,'TTL_actif',f'C{r}'); name(wb,ws,'TTL_actif_stress',f'E{r}')
    S.blank()
    S.block('PASSIF - délai de sortie d\'un porteur après la Période de Blocage')
    S.hdr(['Mesure','Jours','','','','','Lecture'])
    r=S.row(['Règlement après la date de VL (35 j.o.)','=ROUND(Reglement_JO*7/5,0)','','','','','Art. 7.2.5 : publication sous 30 j.o. puis paiement sous 5 j.o.'],fmts=[None,'0']); name(wb,ws,'Reglement_cal',f'C{r}')
    r=S.row(['Délai minimal de sortie (demande déposée à la centralisation)','=Preavis_j+Reglement_cal','','','','','Centralisation le 31/03 ou le 30/09 pour la VL semestrielle suivante.'],fmts=[None,'0']); name(wb,ws,'TTL_passif_min',f'C{r}')
    r=S.row(['Délai moyen de sortie','=TTL_passif_min+91','','','','','Attente moyenne d\'une date de centralisation : un trimestre.'],fmts=[None,'0'],key_cols=(1,)); name(wb,ws,'TTL_passif',f'C{r}')
    r=S.row(['Délai maximal de sortie','=TTL_passif_min+182','','','','','Demande déposée juste après une centralisation.'],fmts=[None,'0']); name(wb,ws,'TTL_passif_max',f'C{r}')
    r=S.row(['Délai moyen avec préavis prolongé (180 jours)','=Preavis_max_j+Reglement_cal+91','','','','','Art. 7.2.7 : outil de gestion de la liquidité, 12 mois au plus.'],fmts=[None,'0']); name(wb,ws,'TTL_passif_prolonge',f'C{r}')
    r=S.row(['Écart actif - passif (délais moyens)','=TTL_actif-TTL_passif','','','','','Positif : l\'actif se liquide bien après la sortie des porteurs ; le plafonnement absorbe l\'écart.'],fmts=[None,'0'],key_cols=(1,)); name(wb,ws,'Ecart_TTL',f'C{r}')
    S.blank()
    S.block('ACTIF LIQUIDABLE PAR TRANCHE - format Annexe IV, à l\'allocation cible')
    S.hdr(['Tranche de liquidation','% actif','','','','','Contenu'])
    b=[('0 à 1 jour',0,'Néant : la trésorerie est placée en monétaire et dépôts à J+2.'),
       ('2 à 7 jours','=W_LIQ','Actifs liquides (art. 21.8).'),
       ('8 à 30 jours',0,'Néant.'),
       ('31 à 90 jours','=W_ARES','Ares AGI : rachat trimestriel sous gate.'),
       ('91 à 180 jours','=W_PG','PG NGI : préavis de 45 j.o. puis règlement.'),
       ('181 à 365 jours',0,'Néant.'),
       ('Plus de 365 jours','=W_FERMES','MAPIF II et MSIG 3 : aucune fenêtre avant échéance.')]
    f=S.r
    for t in b: S.row([t[0],t[1],'','','','',t[2]],fmts=[None,'0 %'])
    r=S.row(['Total',f'=SUM(C{f}:C{S.r-1})','','','','','Doit boucler à 100 %.'],fmts=[None,'0 %'],key_cols=(1,)); name(wb,ws,'Buckets_total',f'C{r}')
    name(wb,ws,'B_2_7',f'C{f+1}'); name(wb,ws,'B_31_90',f'C{f+3}'); name(wb,ws,'B_91_180',f'C{f+4}'); name(wb,ws,'B_365',f'C{f+6}')
    S.blank()
    S.block('COUVERTURE DES RACHATS - ratios structurels à l\'allocation cible')
    S.hdr(['Indicateur','Valeur','Cible','Statut','','','Définition'])
    r=S.row(['Plafond de rachat annuel (deux dates)','=2*Gate','','','','','Deux Dates de Centralisation par an à 5 % chacune.'],fmts=[None,'0 %']); name(wb,ws,'Gate_annuel',f'C{r}')
    r=S.row(['Poche liquide / plafond par date','=W_LIQ/Gate',1.5,'=IF(C{0}>=D{0},"Respectée","Surveillance")'.format(S.r),'','','La poche cible couvre trois dates de plafond sans autre ressource.'],fmts=[None,'0.0"x"','0.0"x"'],key_cols=(1,)); name(wb,ws,'Couv_poche_gate',f'C{r}'); name(wb,ws,'Couv_poche_gate_c',f'D{r}'); name(wb,ws,'Couv_poche_gate_s',f'E{r}')
    r=S.row(['Poche liquide / plafond annuel','=W_LIQ/Gate_annuel',1.0,'=IF(C{0}>=D{0},"Respectée","Surveillance")'.format(S.r),'','','Une année de plafond servie sur la seule poche.'],fmts=[None,'0.0"x"','0.0"x"']); name(wb,ws,'Couv_poche_annuel',f'C{r}'); name(wb,ws,'Couv_poche_annuel_c',f'D{r}'); name(wb,ws,'Couv_poche_annuel_s',f'E{r}')
    r=S.row(['Capacité de rachat semestrielle des fonds evergreen','=W_EVER*0.05*2','','','','','Deux trimestres à 5 % de chaque ligne evergreen, hors gate chez les fonds cibles.'],fmts=[None,'0.0 %']); name(wb,ws,'Cap_evergreen',f'C{r}')
    r=S.row(['Capacité evergreen / plafond par date','=Cap_evergreen/Gate',1.0,'=IF(C{0}>=D{0},"Respectée","Surveillance")'.format(S.r),'','','Les deux lignes evergreen ne financent pas seules une date de plafond.'],fmts=[None,'0.0"x"','0.0"x"'],key_cols=(1,)); name(wb,ws,'Couv_evergreen_gate',f'C{r}'); name(wb,ws,'Couv_evergreen_gate_c',f'D{r}'); name(wb,ws,'Couv_evergreen_gate_s',f'E{r}')
    r=S.row(['Ressources à douze mois / plafond annuel (classeur de stress tests)',2.96,1.5,'=IF(C{0}>=D{0},"Respectée","Surveillance")'.format(S.r),'','','Classeur du 24/09/2026 : poche, distributions et rachats evergreen sur 12 mois, départ 30/09/2028.'],fmts=[None,'0.00"x"','0.0"x"'],input_cols=(1,)); name(wb,ws,'Ressources_12m',f'C{r}'); name(wb,ws,'Ressources_12m_c',f'D{r}'); name(wb,ws,'Ressources_12m_s',f'E{r}')
    r=S.row(['Couverture de liquidité à douze mois au départ (nette des appels, frais et acomptes)',2.38,'=S_couv_vert','=IF(C{0}>=D{0},"Respectée","Surveillance")'.format(S.r),'','','Même classeur ; seuil vert 1,5.'],fmts=[None,'0.00"x"','0.0"x"'],input_cols=(1,)); name(wb,ws,'Couv_12m',f'C{r}'); name(wb,ws,'Couv_12m_c',f'D{r}'); name(wb,ws,'Couv_12m_s',f'E{r}')
    S.blank(); S.ncols=7
    S.note('Convention TTL : délai complet pour disposer du cash, préavis et règlement compris ; une position sans fenêtre garantie est classée au-delà de 365 jours. Les délais des fonds cibles sont ceux du classeur de stress tests du 24/09/2026 ; le passif est recalculé sur le préavis contractuel du projet du 30/09/2026 (91 jours), là où la note de liquidité retenait 180 jours.')

def build_scenarios(wb):
    ws=wb.create_sheet('3 Scénarios')
    S=Sheet(ws,'STRESS TESTS DE LIQUIDITÉ - résultats du classeur du 24/09/2026 et verdicts',
            'Projection trimestrielle sur 36 mois à partir du 30/09/2028, actif net de 50 M€, allocation cible, plafond 5 % par date, 4 dates au plus. Valeurs importées ; verdicts recalculés sur les seuils de la note de liquidité.',
            {'A':2,'B':40,'C':11,'D':11,'E':12,'F':11,'G':12,'H':12,'I':11,'J':11,'K':12,'L':10,'M':44})
    S.hdr(['Scénario','Dates plaf.','File max','Résorption (mois)','Poche min','Couv. 12 m min','Sur-eng. max','Change max','Cession (M€)','Suspension','Verdict','Lecture'])
    def verdict(r):
        return (f'=IF(OR(D{r}>S_file_rouge,F{r}<S_poche_rouge,G{r}<S_couv_rouge,E{r}>S_resorption_rouge,K{r}<>"-"),"Rouge",'
                f'IF(OR(C{r}>0,F{r}<Liq_min,G{r}<S_couv_vert,H{r}>S_sureng_orange),"Orange","Vert"))')
    data=[('Base : rachats 2 % par date, collecte 5 % par trimestre',0,0,0,0.1637,2.47,0.968,0.220,0,'-','Régime de croisière : aucune date plafonnée, poche au-dessus du plancher.'),
          ('Adverse : rachats de 10 % sur deux dates',4,0.1049,24,0.0761,1.148,1.022,0.343,0,'-','Quatre dates plafonnées, file résorbée en 24 mois ; poche sous le plancher pendant la résorption.'),
          ('Extrême : rachats de 15 % puis 20 %',4,0.2616,99,0.0761,0.683,1.022,0.342,0,'12/2030','File de 26 % non résorbée ; suspension examinée à la 5e date (fin 2030).'),
          ('Concentration : sortie des cinq premiers porteurs (35 %)',0,0,0,0.158,3.69,0.978,0.289,0,'-','Servi sous plafonnement ordinaire ; les autres porteurs demandent 2 % par date.'),
          ('Distributions des fonds cibles divisées par deux et retardées d\'un an',0,0,0,0.1218,2.10,0.999,0.348,0,'-','Poche entamée par les appels de fonds non compensés.'),
          ('Appels de fonds accélérés, dollar + 15 %',0,0,0,0.1323,1.99,1.005,0.358,0,'-','Appels de MAPIF II renchéris par le dollar ; sur-engagement au-dessus de 100 %.'),
          ('Valorisation - 20 % (capital), - 10 % (dette), dollar - 15 %',0,0,0,0.1489,2.19,0.973,0.288,0,'-','Poche juste sous le plancher par effet de dénominateur.'),
          ('Crise combinée : chocs de passif et d\'actif simultanés',4,0.2609,99,0,0.172,1.078,0.334,6.61,'12/2030','Poche épuisée dès 2029, cessions secondaires de 6,6 M€ à 10 % de décote, suspension fin 2030, triple gel dès 12/2028.')]
    f=S.r
    for t in data:
        r=S.r
        S.row([t[0],t[1],t[2],(t[3] if t[3]!=99 else 99),t[4],t[5],t[6],t[7],t[8],t[9],verdict(r),t[10]],
              fmts=[None,'0','0.0 %','0','0.0 %','0.00','0 %','0 %','0.00',None,None],h=30,input_cols=range(1,10))
    l=S.r-1
    r=S.row(['Scénarios verts',f'=COUNTIF(L{f}:L{l},"Vert")','','','','','','','','','',''],fmts=[None,'0']); name(wb,ws,'N_vert',f'C{r}')
    r=S.row(['Scénarios orange',f'=COUNTIF(L{f}:L{l},"Orange")','','','','','','','','','',''],fmts=[None,'0']); name(wb,ws,'N_orange',f'C{r}')
    r=S.row(['Scénarios rouges',f'=COUNTIF(L{f}:L{l},"Rouge")','','','','','','','','','',''],fmts=[None,'0'],key_cols=(1,)); name(wb,ws,'N_rouge',f'C{r}')
    for i,nm in enumerate(['Sc_base','Sc_adverse','Sc_extreme','Sc_top5','Sc_distrib','Sc_appels','Sc_valo','Sc_combine']):
        name(wb,ws,nm,f'L{f+i}')
    name(wb,ws,'Sc_combine_cession',f'J{f+7}'); name(wb,ws,'Sc_extreme_file',f'D{f+2}'); name(wb,ws,'Sc_adverse_poche',f'F{f+1}'); name(wb,ws,'Sc_adverse_resorption',f'E{f+1}')
    S.note('Résorption « 99 » : file non résorbée à l\'horizon de 36 mois. Les scénarios reprennent la note de liquidité § 6 ; le scénario « fonds evergreen au prorata » n\'est pas restitué par le tableau de bord du classeur et reste à documenter.',12)
    S.blank()
    S.block('TESTS INVERSÉS')
    S.hdr(['Rachats persistants par date (% AN)','Dates plaf.','Suspension','Poche min (gate)','Cession sans gate (M€)','Décote (M€)','','','','','Verdict','Lecture'])
    rev=[(0.05,0,'-',0.0338,0,0,'Demandes égales au plafond : servies sans plafonnement, la poche passe sous 5 % dès la deuxième année.'),
         (0.075,4,'12/2030',0.0761,5.16,0.52,'Point de rupture : quatre dates plafonnées puis suspension ; sans gate, 5,2 M€ de cessions.'),
         (0.10,4,'12/2030',0.0761,11.58,1.16,''),(0.125,4,'12/2030',0.0761,17.31,1.73,''),(0.15,4,'12/2030',0.0761,22.33,2.23,'')]
    f=S.r
    for t in rev:
        r=S.r
        S.row([t[0],t[1],t[2],t[3],t[4],t[5],'','','','',f'=IF(OR(D{r}<>"-",E{r}<S_poche_rouge),"Rouge",IF(C{r}>0,"Orange","Vert"))',t[6]],fmts=['0.0 %','0',None,'0.0 %','0.00','0.00'],h=26,input_cols=range(0,6))
        if t[0]==0.075: name(wb,ws,'Rupture_cession',f'F{r}')
    r=S.row(['Point de rupture (suspension ou cession)',0.075,'','','','','','','','','',''],fmts=[None,'0.0 %'],key_cols=(1,)); name(wb,ws,'Rupture_rachats',f'C{r}')
    S.blank()
    S.hdr(['Appels de fonds à 12 mois (% AN)','Non-appelé départ','Poche min','Couv. 12 m min','Cession (M€)','','','','','','Verdict','Lecture'])
    app=[(0.075,0.15,0.1397,2.08,0,''),(0.125,0.25,0.0747,1.50,0,'Poche sous le plancher de 15 %.'),(0.175,0.35,0,0.93,0.68,'Poche épuisée, cessions.'),(0.225,0.45,0,0.61,6.41,''),(0.45,0.45,0,0.45,12.39,'Tout le non-appelé tiré en un an.')]
    for t in app:
        r=S.r
        S.row([t[0],t[1],t[2],t[3],t[4],'','','','','',f'=IF(OR(D{r}<S_poche_rouge,E{r}<S_couv_rouge),"Rouge",IF(OR(D{r}<Liq_min,E{r}<S_couv_vert),"Orange","Vert"))',t[5]],fmts=['0.0 %','0 %','0.0 %','0.00','0.00'],h=22,input_cols=range(0,5))
    r=S.row(['Point de rupture : poche sous le plancher',0.125,'','','','','','','','','',''],fmts=[None,'0.0 %'],key_cols=(1,)); name(wb,ws,'Rupture_appels',f'C{r}')
    S.blank()
    S.hdr(['Sortie d\'un porteur unique (% AN)','Dates plaf.','File max','Résorption (mois)','Poche min','Suspension','','','','','Verdict','Lecture'])
    port=[(0.10,2,0.052,12,0.101,'-',''),(0.15,4,0.104,24,0.0706,'-','Taille maximale sans suspension ni cession.'),(0.20,4,0.156,99,0.0761,'12/2030',''),(0.25,4,0.208,99,0.0761,'12/2030','')]
    for t in port:
        r=S.r
        S.row([t[0],t[1],t[2],t[3],t[4],t[5],'','','','',f'=IF(OR(G{r}<>"-",E{r}>S_resorption_rouge),"Rouge",IF(C{r}>0,"Orange","Vert"))',t[6]],fmts=['0 %','0','0.0 %','0','0.0 %',None],h=22,input_cols=range(0,6))
    r=S.row(['Sortie maximale d\'un porteur sans suspension',0.15,'','','','','','','','','',''],fmts=[None,'0 %'],key_cols=(1,)); name(wb,ws,'Porteur_max',f'C{r}')
    S.blank(); S.ncols=12
    S.note('Hypothèses du classeur : poche de 15 % au départ, non-appelé 15 % tiré sur trois ans, rachats evergreen à 5 % par trimestre, décote de cession 10 %, frais 2 % par an, acomptes 5 % sur 20 % de parts de distribution, toutes les parts rachetables dès le départ (prudent). Préavis retenu : 180 jours, à recaler sur les 91 jours du projet du 30/09/2026 ; un préavis plus court ne change ni la file ni les verdicts, il réduit le délai d\'information de la Société de Gestion avant chaque date.')
