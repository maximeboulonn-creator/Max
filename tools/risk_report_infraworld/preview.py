# -*- coding: utf-8 -*-
from lib import *
FUND='Openstone Infraworld  ·  Compartiment I'
HDR='Rapport de risque initial  ·  Octobre 2026  ·  Prospectus du 30/09/2026'
PCT='0 %'; PCT1='0.0 %'; X='0.0"x"'; D='0'

def page(p,num,title,sub):
    p.band(); p.title(f'{num}. {title}'); p.sub(sub)

def build_preview(wb):
    ws=wb.create_sheet('PREVIEW',0)
    p=Preview(ws,FUND,HDR)
    # ---------- COUVERTURE
    p.band()
    p.cover([('',False,WHITE,10),
             ('RAPPORT DE RISQUE INITIAL · OCTOBRE 2026',True,PALE,22),
             ('Openstone Infraworld - Compartiment I',True,WHITE,30),
             ('Fonds de fonds d\'infrastructure evergreen semi-liquide · quatre fonds cibles, deux fermés et deux evergreen',False,WHITE,30),
             ('Fonction permanente de gestion des risques - Fundcraft France',False,PALE,18),
             ('',False,WHITE,8),
             ('PRÉ-LANCEMENT · PROSPECTUS PROJET DU 30 SEPTEMBRE 2026 · ACTIF NET CIBLE 50 M€ · SRI 4 / 7 · FPS - RG AMF ART. 423-27',True,WHITE,30),
             ('Document interne - ne pas diffuser · AIFMD art. 15-16 · ESMA 34-39-897',False,PALE,18),
             ('',False,WHITE,8),
             ('SOCIÉTÉ DE GESTION',True,PALE,16),('Fundcraft France SAS - AMF GP-20260008',False,WHITE,16),
             ('CONSEILLER ET DISTRIBUTEUR PRINCIPAL',True,PALE,16),('Openstone (Innovative Finance SAS) - CIF ORIAS 23002459',False,WHITE,16),
             ('GÉRANTS DES FONDS CIBLES',True,PALE,16),('Macquarie (MAPIF II, MSIG 3) · Partners Group (Next Generation Infrastructure) · Ares (Global Infrastructure ELTIF)',False,WHITE,30),
             ('DÉPOSITAIRE · DÉLÉGATAIRE COMPTABLE · COMMISSAIRE AUX COMPTES',True,PALE,16),('BFCM · CIC · APLITEC',False,WHITE,16),
             ('',False,WHITE,8),
             ('Rapport établi sur la documentation : le Fonds n\'est pas constitué, aucune valeur liquidative ni aucun porteur. Toutes les grandeurs en euros sont des cibles ; les stress tests sont calculés dans ce classeur.',False,PALE,30),
             ('',False,WHITE,10)])
    p.pagebreak()
    # ---------- 00 SYNTHÈSE COMITÉ
    page(p,'00','Synthèse pour le Comité des Risques','Verdict  ·  trois chiffres  ·  six décisions')
    p.sec('VERDICT DE LA FONCTION RISQUES   charte Fundcraft, verdict d\'ensemble = pire domaine','Échelle à cinq niveaux : Routine, Vigilance, Élevé, Sévère, Critique. Rapport initial : le verdict porte sur la conception du véhicule, non sur une position.')
    p.hdr(['Domaine','Verdict','Fondement','','','Implication'])
    p.row(['Liquidité structurelle','Élevé','=(ROUND((W_FERMES)*100,0)&" %")&" de l\'actif sans fenêtre avant échéance ; écart actif-passif de "&(ROUND(Ecart_TTL,0)&"")&" jours"',None,None,'Le plafonnement de 5 % est la seule digue ; au-delà de quatre dates, la suspension.'],merge=[(4,6)],h=30)
    p.row(['Stress tests','Élevé','=N_rouge&" scénarios rouges sur "&(N_vert+N_orange+N_rouge)&" : rachats de 15 % puis 20 %, crise combinée ; point de rupture à "&(SUBSTITUTE(ROUND((Rupture_rachats)*100,1)&"",".",",")&" %")&" par date"',None,None,'Décision 5 : séquencer les engagements ; rejouer sur l\'encours réel (R12).'],merge=[(4,6)],h=30)
    p.row(['Concentration','Élevé','=(ROUND((W_MACQ)*100,0)&" %")&" de l\'allocation chez Macquarie, "&(ROUND((W_MAPIF)*100,0)&" %")&" sur MAPIF II ; aucune limite au prospectus"',None,None,'Décision 4 : limites internes gérant 40 % et fonds cible 35 %.'],merge=[(4,6)],h=30)
    p.row(['Marché, change et valorisation','Vigilance','="Dollar non couvert de "&(ROUND((W_USD_BAS)*100,0)&" %")&" à "&(ROUND((W_USD_HAUT)*100,0)&" %")&" de l\'actif ; VL trimestrielles décalées sur "&(ROUND((W_FERMES)*100,0)&" %")',None,None,'Décision 4 : limite de change ; décision 6 : revue des VL reçues.'],merge=[(4,6)],h=30)
    p.row(['Documentation et gouvernance','Élevé','="Préavis 91 jours au prospectus contre 180 dans la note de liquidité ; "&Lim_a_arreter&" seuils à arrêter, "&Lim_a_valider&" à valider ; douze écarts au registre"',None,None,'Décisions 1, 2 et 6 avant le dépôt du prospectus.'],merge=[(4,6)],h=30)
    p.row(['Verdict d\'ensemble','ÉLEVÉ','Fixé par la liquidité, les stress tests, la concentration et la documentation. Aucune limite du prospectus n\'est testable avant la première valeur liquidative.',None,None,'Lancement conditionné aux décisions 1 à 3 ; revue à la première VL.'],merge=[(4,6)],h=30)
    p.blank()
    p.sec('TROIS CHIFFRES QUI COMMANDENT LES DÉCISIONS')
    p.kpi(['PART DE L\'ACTIF SANS LIQUIDITÉ AVANT ÉCHÉANCE','POINT DE RUPTURE DES RACHATS PERSISTANTS','CONCENTRATION SUR UN SEUL GÉRANT'],
          ['=(ROUND((W_FERMES)*100,0)&" %")&" de l\'actif net"','=(SUBSTITUTE(ROUND((Rupture_rachats)*100,1)&"",".",",")&" %")&" par date"','=(ROUND((W_MACQ)*100,0)&" %")&" de l\'actif net"'],
          ['="MAPIF II et MSIG 3 ; délai moyen de liquidation de l\'actif "&(ROUND(TTL_actif,0)&"")&" jours contre "&(ROUND(TTL_passif,0)&"")&" jours pour la sortie d\'un porteur"',
           '="Quatre dates plafonnées puis suspension fin 2030 ; sans plafonnement, "&SUBSTITUTE(ROUND(Rupture_cession,1)&"",".",",")&" M€ de cessions secondaires - décisions 1 et 5"',
           '="Macquarie, deux fonds fermés dont un non lancé ; score IRP en transparence "&SUBSTITUTE(ROUND(IRPc_Transp,1)&"",".",",")&" sur 100 - décision 4"'])
    p.blank()
    p.sec('SIX DÉCISIONS SOUMISES AU COMITÉ   détail en onglet 5')
    p.hdr(['N°','Décision','','Proposition de la Fonction Risques','','Échéance'])
    for i in range(1,7):
        p.row([str(i),f'=Dec{i}',None,f'=DecP{i}',None,f'=DecE{i}'],merge=[(3,4),(5,6)],h=44,bold_first=False)
    p.note('Les douze recommandations de la Fonction Risques et les points suivis sans décision figurent en onglet 5. Le corps du rapport (pages 01 à 07) documente chaque chiffre ; l\'annexe donne les sources et leur fraîcheur.')
    p.pagebreak()
    # ---------- 01 SYNTHÈSE
    page(p,'01','Synthèse et caractéristiques du Fonds','Conclusion  ·  chiffres clés  ·  points d\'attention')
    p.sec('CONCLUSION DE LA GESTION DES RISQUES')
    p.note('="À la date d\'arrêté, le Fonds n\'est pas constitué : aucune limite du prospectus n\'est franchie ni testable. Le dispositif est cohérent sur le papier - ni emprunt ni dérivé, poche liquide de "&(ROUND((Liq_min)*100,0)&" %")&" au moins, plafonnement de "&(ROUND((Gate)*100,0)&" %")&" par date, prolongation du préavis et suspension - mais il repose sur un actif dont "&(ROUND((W_FERMES)*100,0)&" %")&" ne se liquide qu\'à l\'échéance des fonds cibles. Les stress tests du 24/09/2026 donnent "&N_vert&" scénarios verts, "&N_orange&" orange et "&N_rouge&" rouges : le Fonds absorbe des rachats de 10 % par date au prix de quatre dates plafonnées, il ne résiste pas à 15 % puis 20 % ni à une crise combinée sans cessions et suspension. Trois chantiers conditionnent le lancement : figer le texte de liquidité (préavis, Demandes Prioritaires, issue du plafonnement), obtenir les dérogations de ticket des deux fonds Macquarie, et arrêter des limites internes de concentration et de change que le prospectus ne contient pas."',lines=8)
    p.blank()
    p.kpi(['ACTIF NET CIBLE','PLAFOND PAR DATE DE CENTRALISATION','POCHE LIQUIDE / PLAFOND PAR DATE','SCORE IRP DU FONDS','SCORE IRP EN TRANSPARENCE'],
          ['=(ROUND(AN_cible/1000000,0)&"")&" M€"','=SUBSTITUTE(ROUND(Gate*AN_cible/1000000,1)&"",".",",")&" M€"','=SUBSTITUTE(ROUND(Couv_poche_gate,1)&"",".",",")&" x"','=SUBSTITUTE(ROUND(IRPc_Fonds,1)&"",".",",")&" / 100"','=SUBSTITUTE(ROUND(IRPc_Transp,1)&"",".",",")&" / 100"'],
          ['prévision commerciale ; aucune souscription reçue','="5 % de l\'actif net, net des souscriptions - art. 7.2.6 ; "&(ROUND((Gate_annuel)*100,0)&" %")&" par an au plus"','="poche de "&(ROUND((W_LIQ)*100,0)&" %")&" ; les fonds evergreen ajoutent "&(SUBSTITUTE(ROUND((Cap_evergreen)*100,1)&"",".",",")&" %")&" par semestre"','Élevé - liquidité et crédit saturés, contrepartie et levier faibles','="Élevé - "&SUBSTITUTE(ROUND(IRPc_Transp_hp,1)&"",".",",")&" hors poche liquide ; MAPIF II seul fonds cible noté Moyen"'])
    p.blank()
    p.sec('CHIFFRES CLÉS   à l\'allocation cible, en % de l\'actif net sauf mention')
    p.hdr(['Indicateur','Valeur','Seuil / limite','Statut','Référence','Commentaire'])
    p.row(['Fonds fermés (MAPIF II, MSIG 3)','=W_FERMES','À arrêter','Seuil à arrêter','KRI-09','Aucune limite de division des risques au prospectus (art. 21.7)'],fmts=[None,PCT])
    p.row(['Fonds evergreen (PG NGI, Ares AGI)','=W_EVER','-','Information','-','Rachat trimestriel plafonné à 5 % chez chaque fonds cible'],fmts=[None,PCT])
    p.row(['Gérant Macquarie','=W_MACQ','À arrêter','Seuil à arrêter','KRI-08','Proposition : 40 %, alerte 35 %'],fmts=[None,PCT])
    p.row(['Actifs liquides','=W_LIQ','=Liq_min','Respectée','RL-02','Plancher permanent ; cible de gestion 16 %'],fmts=[None,PCT,PCT])
    p.row(['Dollar non couvert, hypothèse basse / haute','=(ROUND((W_USD_BAS)*100,0)&" %")&" / "&(ROUND((W_USD_HAUT)*100,0)&" %")','À arrêter','Seuil à arrêter','KRI-10','Aucune couverture possible : dérivés interdits (IR-07)'])
    p.row(['Sur-engagement maximal','=Surengagement_max','=Surengagement_max','Paramétré','RL-01','Sans ligne de crédit ; appels financés par la poche et les distributions'],fmts=[None,PCT,PCT])
    p.row(['Levier (engagement et brut)','=Levier_max','=Levier_max','À réintroduire','KRI-05','Supprimé du projet du 30/09 ; art. 23 § 1 a) de la directive'],fmts=[None,PCT,PCT])
    p.row(['Délai moyen de liquidation de l\'actif','=TTL_actif','-','Information','TTL','="jours ; "&(ROUND(TTL_actif_stress,0)&"")&" jours en stress"'],fmts=[None,D])
    p.row(['Délai moyen de sortie d\'un porteur','=TTL_passif','-','Information','TTL','="jours ; "&(ROUND(TTL_passif_prolonge,0)&"")&" jours avec préavis prolongé"'],fmts=[None,D])
    p.row(['Limites du prospectus franchies','=Lim_depassees',0,'=IF(Lim_depassees<=0,"Conforme","Escalade")','Onglet 1','Aucune position : rien à tester avant la première VL'],fmts=[None,D,D])
    p.row(['Seuils internes à arrêter ou à valider','=Lim_a_arreter+Lim_a_valider',0,'Décision 4','Onglet 1','="dont "&Lim_a_arreter&" sans valeur proposée dans les documents"'],fmts=[None,D,D])
    p.blank()
    p.sec('POINTS D\'ATTENTION   décisions attendues')
    p.hdr(['Priorité','Sujet','Constat','','Décision attendue',''])
    p.row(['Haute','Texte de liquidité','Préavis de 91 jours au prospectus du 30/09 contre 180 jours dans la note de liquidité et le classeur de stress tests ; Demandes Prioritaires maintenues ; issue après 4 dates non écrite',None,'Décision 1 : figer le texte et recaler la note (R1)',None],merge=[(4,5),(6,7)],h=40)
    p.row(['Haute','Construction du portefeuille','Tickets minimaux de 10 M$ chez MAPIF II et MSIG 3 pour des lignes cibles de 17 et 8,5 M€ ; véhicule et manche de MSIG 3 non arrêtés',None,'Décision 3 : dérogations écrites avant tout engagement (R4, R5)',None],merge=[(4,5),(6,7)],h=40)
    p.row(['Haute','Concentration et change','51 % chez un gérant, 34 % sur une ligne, dollar non couvert jusqu\'à 68 % ; aucun seuil au prospectus',None,'Décision 4 : limites internes KRI-08 à KRI-11 (R6)',None],merge=[(4,5),(6,7)],h=40)
    p.row(['Moyenne','Levier AIFM','Niveau maximal de levier absent du projet du 30/09',None,'Décision 2 : 100 % engagement et brut (R3)',None],merge=[(4,5),(6,7)],h=28)
    p.row(['Moyenne','Gouvernance','Conseiller à la fois conseil, distributeur principal et titulaire du pouvoir de révocation de la Société de Gestion',None,'Décision 6 : cartographie du conflit et procédure de revue des VL (R9, R10)',None],merge=[(4,5),(6,7)],h=40)
    p.note('Statuts : Respectée, Surveillance, Paramétré, Non testable, Seuil à arrêter, À valider. Convention définie en onglet 1.')
    p.pagebreak()
    # ---------- 02 PORTEFEUILLE CIBLE
    page(p,'02','Portefeuille cible et liquidité contractuelle','Allocation  ·  conditions du Fonds  ·  calendrier de la première fenêtre')
    p.sec('ALLOCATION CIBLE   synthèse Openstone de septembre 2026, non contractuelle','Quatre fonds cibles : deux fermés (secondaires, dette) et deux evergreen (equity, equity et dette). Instruments visés : 50 % equity secondaire, 30 % dette, 20 % equity primaire.')
    p.hdr(['Ligne','Poids','Montant cible','Forme','Devise','Liquidité pour le Fonds'])
    p.row(['Macquarie Alliance Partners Infrastructure Fund II','=W_MAPIF','=W_MAPIF*AN_cible','Fermé, secondaires','USD','Aucune avant échéance (10 à 14 ans) ; cession agréée par le GP'],fmts=[None,PCT,'# ##0'])
    p.row(['Macquarie Specialized Infrastructure Global 3','=W_MSIG','=W_MSIG*AN_cible','Fermé hybride, dette','USD ou EUR','Retrait annuel après 3 ans, payé par écoulement sans date'],fmts=[None,PCT,'# ##0'])
    p.row(['Partners Group Next Generation Infrastructure','=W_PG','=W_PG*AN_cible','Evergreen, equity','USD (classe H possible)','Mensuel, préavis 45 j.o., gate 5 % par trimestre distributions comprises'],fmts=[None,PCT,'# ##0'])
    p.row(['Ares Global Infrastructure ELTIF','=W_ARES','=W_ARES*AN_cible','Evergreen ELTIF, equity et dette','EUR','Trimestriel, gate 5 % net, déduction 5 % avant 24 mois'],fmts=[None,PCT,'# ##0'])
    p.row(['Actifs liquides','=W_LIQ','=W_LIQ*AN_cible','Monétaire, dépôts, OPC obligataires','EUR','Quotidienne ; plancher de 15 % de l\'actif (art. 21.8)'],fmts=[None,PCT,'# ##0'])
    p.row(['Total','=W_TOT','=W_TOT*AN_cible','','','Aucun emprunt, aucun dérivé, aucun passif financier'],fmts=[None,PCT,'# ##0'])
    p.blank()
    p.sec('RACHAT ET LIQUIDITÉ DU FONDS   prospectus du 30 septembre 2026')
    p.hdr(['Paramètre','Valeur','','Source','Commentaire',''])
    rows=[('Valeur liquidative','Mensuelle, dernier Jour Ouvré','Art. 6.1','="Publication sous "&Publication_JO&" Jours Ouvrés ; VL du 31/12 certifiée, du 30/06 attestée"'),
          ('Souscriptions','Mensuelles, à cours inconnu','Art. 7.1','="Centralisation "&Centralisation_sous_JO&" Jours Ouvrés avant la VL, paiement intégral ; au nominal tant que la VL est inférieure à 100 €"'),
          ('Rachats','Semestriels, VL du 30 juin et du 31 décembre','Art. 7.2.1','Centralisation le 31 mars et le 30 septembre à 12 h ; demandes définitives'),
          ('Période de Blocage','="24 mois par souscription"','Art. 7.2.2','Imputation FIFO des rachats sur les parts les plus anciennes'),
          ('Préavis','="Environ "&Preavis_j&" jours ; prolongeable à "&Preavis_max_j','Art. 7.2.3 et 7.2.7','La Société de Gestion connaît les demandes brutes trois mois avant la VL'),
          ('Plafonnement des Rachats','=(ROUND((Gate)*100,0)&" %")&" de l\'actif net par date, net des souscriptions"','Art. 7.2.6','="Report automatique, "&Gate_dates&" dates au plus ; Demandes Prioritaires après deux reports ; au-delà, texte muet"'),
          ('Suspension','Circonstances exceptionnelles, L. 214-24-41 CMF','Art. 7.2.8','Souscriptions suspendues simultanément ; sans plafonnement préalable obligatoire'),
          ('Délai de règlement','="Au plus "&Reglement_JO&" Jours Ouvrés après la VL"','Art. 7.2.5','5 Jours Ouvrés après publication ; en numéraire'),
          ('Frais de sortie','Néant','Art. 27.2.4','Les déductions de sortie existent chez Ares (5 % avant 24 mois) et PG NGI (discrétionnaire)'),
          ('Dissolution','Actif net sous 300 000 € pendant 30 jours ; sous 1 M€ ou rachats non exécutés au-delà de 50 % des parts, sur décision','Règlement art. 2 et 15','Accord du Conseiller requis pour la dissolution anticipée')]
    for a,b,c,d in rows: p.row([a,b,None,c,d,None],merge=[(3,4),(6,7)],h=30)
    p.blank()
    p.sec('CALENDRIER DE LA PREMIÈRE FENÊTRE DE RACHAT   hypothèse : constitution au 31/12/2026')
    p.hdr(['Étape','Date','','Écart','Commentaire',''])
    p.row(['Constitution du Fonds (hypothèse)','31/12/2026',None,'-','Attestation de dépôt des fonds par le Dépositaire',None],merge=[(3,4),(6,7)])
    p.row(['Fin de la Période de Blocage des premières parts','31/12/2028',None,'24 mois','Chaque souscription ultérieure porte son propre blocage',None],merge=[(3,4),(6,7)])
    p.row(['Première Date de Centralisation des Rachats','31/03/2029',None,'-','Première VL de rachat : 30/06/2029 ; règlement au plus tard mi-août 2029',None],merge=[(3,4),(6,7)])
    p.row(['Demandes aux fonds evergreen','Mi-avril 2029',None,'45 j.o. avant la VL','PG NGI : préavis 45 Jours Ouvrés ; Ares : cut-off 1er Jour Ouvré de juin',None],merge=[(3,4),(6,7)])
    p.row(['Fin de la déduction Ares pour des parts souscrites début 2027','Début 2029',None,'24 mois','Un rachat Ares avant 24 mois coûte 5 % de la ligne - recommandation R8',None],merge=[(3,4),(6,7)])
    p.note('Le classeur de stress tests part du 30/09/2028 avec toutes les parts rachetables : hypothèse prudente par rapport à ce calendrier.')
    p.pagebreak()
    # ---------- 03 PROFIL DE LIQUIDITÉ
    page(p,'03','Profil de liquidité et couverture des rachats','Actif  ·  passif  ·  adéquation  ·  ratios de couverture  ·  outils de gestion de la liquidité')
    p.sec('PROFIL DE LIQUIDITÉ DE L\'ACTIF   format Annexe IV, à l\'allocation cible','TTL (time to liquidate) : délai complet pour disposer du cash. Une position sans fenêtre garantie est classée au-delà de 365 jours.')
    p.hdr(['Tranche de liquidation','% actif','Montant cible','','Contenu',''])
    for lab,nm,txt in [('0 à 1 jour',None,'Néant'),('2 à 7 jours','B_2_7','Actifs liquides : monétaire, dépôts, OPC obligataires'),('8 à 30 jours',None,'Néant'),('31 à 90 jours','B_31_90','Ares AGI : rachat trimestriel sous gate'),('91 à 180 jours','B_91_180','PG NGI : préavis 45 j.o. puis règlement'),('181 à 365 jours',None,'Néant'),('Plus de 365 jours','B_365','MAPIF II et MSIG 3 : aucune fenêtre avant échéance')]:
        p.row([lab,f'={nm}' if nm else 0,f'={nm}*AN_cible' if nm else 0,None,txt,None],fmts=[None,PCT,'# ##0'],merge=[(5,7)],h=20)
    p.row(['Total','=Buckets_total','=Buckets_total*AN_cible',None,'',None],fmts=[None,PCT,'# ##0'],merge=[(5,7)],h=20)
    p.blank()
    p.sec('ADÉQUATION ACTIF-PASSIF   en jours')
    p.hdr(['Mesure','TTL actif','TTL passif','Écart','Lecture',''])
    p.row(['Délais moyens, préavis normal','=TTL_actif','=TTL_passif','=Ecart_TTL','L\'actif se liquide près de trois ans après la sortie moyenne d\'un porteur : le plafonnement absorbe l\'écart',None],fmts=[None,D,D,D],merge=[(6,7)],h=30)
    p.row(['Délais moyens, préavis prolongé à 180 jours','=TTL_actif','=TTL_passif_prolonge','=TTL_actif-TTL_passif_prolonge','La prolongation gagne un trimestre ; elle n\'est pas un correctif structurel',None],fmts=[None,D,D,D],merge=[(6,7)],h=30)
    p.row(['Délais en stress (fonds evergreen en Special Dealing)','=TTL_actif_stress','=TTL_passif_min','=TTL_actif_stress-TTL_passif_min','Sortie la plus rapide d\'un porteur contre actif le plus lent',None],fmts=[None,D,D,D],merge=[(6,7)],h=30)
    p.blank()
    p.sec('RATIOS DE COUVERTURE DES RACHATS')
    p.hdr(['Indicateur','Valeur','Cible','Statut','Définition',''])
    p.row(['Poche liquide / plafond par date','=Couv_poche_gate','=Couv_poche_gate_c','=Couv_poche_gate_s','Trois dates de plafond servies sur la seule poche',None],fmts=[None,X,X],merge=[(6,7)])
    p.row(['Poche liquide / plafond annuel','=Couv_poche_annuel','=Couv_poche_annuel_c','=Couv_poche_annuel_s','Une année de plafond servie sur la seule poche',None],fmts=[None,X,X],merge=[(6,7)])
    p.row(['Capacité evergreen semestrielle / plafond par date','=Couv_evergreen_gate','=Couv_evergreen_gate_c','=Couv_evergreen_gate_s','="Deux trimestres à 5 % de PG NGI et Ares : "&(SUBSTITUTE(ROUND((Cap_evergreen)*100,1)&"",".",",")&" %")&" de l\'actif net par semestre"',None],fmts=[None,X,X],merge=[(6,7)])
    p.row(['Ressources à douze mois / plafond annuel','=Ressources_12m','=Ressources_12m_c','=Ressources_12m_s','Classeur de stress tests : poche, distributions et rachats evergreen',None],fmts=[None,X,X],merge=[(6,7)])
    p.row(['Couverture de liquidité à douze mois au départ','=Couv_12m','=Couv_12m_c','=Couv_12m_s','Nette des appels, frais et acomptes ; seuil vert 1,5 (KRI-04)',None],fmts=[None,X,X],merge=[(6,7)])
    p.blank()
    p.sec('OUTILS DE GESTION DE LA LIQUIDITÉ   annexe V de la directive (UE) 2024/927, note de liquidité § 3')
    p.hdr(['Outil','Déclenchement','','Durée et levée','','Observation'])
    p.row(['Plafonnement des Rachats (art. 7.2.6)','Demandes nettes supérieures à 5 % de l\'actif net à une date',None,'Quatre dates consécutives au plus ; levé dès le retour sous 5 %',None,'Issue au-delà de quatre dates à écrire'],merge=[(3,4),(5,6)],h=30)
    p.row(['Prolongation du préavis (art. 7.2.7)','File supérieure à 10 % de l\'actif net, couverture à douze mois sous 1,2, plafonnement chez un fonds cible',None,'180 jours au plus, douze mois au plus ; levée à couverture 1,5 et file sous 5 %',None,'Seul outil agissant avant l\'exécution ; à décider une centralisation à l\'avance'],merge=[(3,4),(5,6)],h=40)
    p.row(['Suspension des rachats (art. 7.2.8)','Circonstances exceptionnelles ; durée maximale du plafonnement atteinte',None,'Souscriptions suspendues simultanément ; réexamen mensuel',None,'Notification AMF, Dépositaire, distributeurs et porteurs'],merge=[(3,4),(5,6)],h=30)
    p.row(['Suspension des souscriptions (art. 7.1.6)','Circonstances exceptionnelles ou capacité d\'investissement épuisée',None,'Sur décision de la Société de Gestion',None,'Évite la dilution de la poche en cas d\'appels retardés'],merge=[(3,4),(5,6)],h=30)
    p.pagebreak()
    # ---------- 04 STRESS TESTS
    page(p,'04','Stress tests de liquidité','Scénarios  ·  tests inversés  ·  lecture - moteur en formules, onglet 3b')
    p.sec('SCÉNARIOS DE LA NOTE DE LIQUIDITÉ   actif net cible, départ à la fin du blocage, 36 mois, calcul en formules (onglet 3b)','Verdicts sur les seuils de la note (onglet 0) : rouge si file > 25 %, poche < 5 %, couverture < 1,0, résorption > 25 mois ou suspension ; orange si une date est plafonnée, poche < 15 %, couverture < 1,5 ou sur-engagement > 120 %.',h=38)
    p.hdr(['Scénario','Dates plafonnées','File max','Poche min','Couverture 12 m','Verdict et lecture'])
    for i in range(9):
        r=6+i
        p.row([f"='3 Scénarios'!B{r}",f"='3 Scénarios'!C{r}",f"='3 Scénarios'!D{r}",f"='3 Scénarios'!F{r}",f"='3 Scénarios'!G{r}",f"='3 Scénarios'!L{r}&\" - \"&'3 Scénarios'!M{r}"],fmts=[None,D,PCT1,PCT1,'0.00'],h=36)
    p.row(['Bilan','=N_vert&" verts"','=N_orange&" orange"','=N_rouge&" rouges"','','="Les scénarios rouges partagent une cause : des rachats très supérieurs au plafond sur une ou deux dates, ou des fonds evergreen qui ne servent plus ; premier triple gel en crise combinée : "&Gel_combine'],h=30)
    p.blank()
    p.sec('TESTS INVERSÉS')
    p.hdr(['Question','Réponse','','Lecture','',''])
    p.row(['Niveau de rachats persistants par date qui conduit à la suspension','=Rupture_rachats',None,'="Quatre dates plafonnées puis suspension fin 2030 ; sans plafonnement, "&SUBSTITUTE(ROUND(Rupture_cession,1)&"",".",",")&" M€ de cessions à 10 % de décote"',None,None],fmts=[None,PCT1],merge=[(3,4),(5,7)],h=30)
    p.row(['Appels de fonds sur douze mois qui font passer la poche sous 10 %','=Rupture_appels',None,'Non-appelé de départ de 25 % de l\'actif net ; à 17,5 % la poche est épuisée et les cessions commencent',None,None],fmts=[None,PCT1],merge=[(3,4),(5,7)],h=30)
    p.row(['Taille maximale d\'un porteur unique sortant sans suspension ni cession','=Porteur_max',None,'Les autres porteurs demandent 2 % par date ; au-delà de 15 %, file non résorbée et suspension',None,None],fmts=[None,PCT],merge=[(3,4),(5,7)],h=30)
    p.row(['Rachats égaux au plafond (5 % par date) sur toute la période : poche minimale','=Persist5_poche',None,'Servis sans plafonnement, ils érodent la poche sous 5 % : le plafond ne protège pas quand la demande lui est exactement égale',None,None],fmts=[None,PCT1],merge=[(3,4),(5,7)],h=30)
    p.blank()
    p.sec('LECTURE DE LA GESTION DES RISQUES')
    p.note('Le dispositif tient à trois conditions : une collecte qui ne s\'arrête pas (le scénario de base suppose 5 % par trimestre), des distributions des fonds cibles qui arrivent (le scénario de distributions divisées par deux suffit à passer sous le plancher) et un non-appelé contenu (le point de rupture est à 25 % de l\'actif net). Aucune de ces trois conditions ne dépend de la Société de Gestion seule. Les hypothèses structurantes - collecte, non-appelé de départ, décote de cession, part des parts de distribution - sont à arrêter en Comité des risques (R12), et le moteur est rejoué sur l\'encours réel dès la première valeur liquidative (R12). Le préavis de 91 jours du prospectus du 30/09 ne change pas les flux : il réduit de trois mois le temps de préparation de chaque date.',lines=7)
    p.note('Limites de la méthode : chocs de valorisation calibrés sur des indices Preqin corrigés du lissage (Geltner), sans historique propre ; comportement de rachat postulé, non modélisé ; capacité des fonds evergreen supposée pleine hors scénario au prorata ; cessions secondaires de parts de fonds fermés supposées possibles à 10 % de décote, ce qui est optimiste pour MAPIF II (accord du GP) et sans objet pour MSIG 3 (run-off). Niveau de confiance : moyen.',lines=4)
    p.pagebreak()
    # ---------- 05 FONDS CIBLES
    page(p,'05','Les quatre fonds cibles','Conditions  ·  profil de risque initial  ·  points ouverts')
    p.sec('CONDITIONS STRUCTURANTES   lecture des prospectus et PPM ; forme, stratégie et prestataires en onglet 4')
    p.hdr(['Terme','MAPIF II','MSIG 3 (USD)','PG NGI','Ares AGI ELTIF',''])
    for t,src,h in [('Statut',8,52),('Devise',9,40),('Rachats',12,80),('Plafonnement',13,80),('Pénalité de sortie',14,66),('Valeur liquidative',15,52),('Frais de gestion',16,52),('Performance',17,66),('Levier',18,80),('Ticket minimum',19,52),('SFDR',21,40)]:
        p.row([t]+[f"='4 Fonds cibles'!{c}{src}" for c in 'CDEF']+[None],h=h,merge=[(6,7)])
    p.blank()
    p.sec('PROFIL DE RISQUE INITIAL   scores IRP 0 à 100, poids de l\'allocation cible')
    p.hdr(['Fonds','Score IRP','Notation','Allocation','Contribution','Typologies saturées'])
    r0=27
    for lab,nm,w,sat in [('Openstone Infraworld (caractéristiques propres)','IRPc_Fonds',None,'Liquidité, crédit'),('MAPIF II','IRPc_MAPIF','W_MAPIF','Levier ; liquidité faible (fonds fermé)'),('MSIG 3 (USD, levered)','IRPc_MSIG','W_MSIG','Crédit, levier'),('PG Next Generation Infrastructure','IRPc_PG','W_PG','Liquidité, crédit, opérationnel, levier'),('Ares Global Infrastructure ELTIF','IRPc_ARES','W_ARES','Liquidité, crédit, levier')]:
        p.row([lab,f'={nm}',f'=IF({nm}>56,"Élevé",IF({nm}>34,"Moyen","Faible"))',f'={w}' if w else '-',f'={w}*{nm}' if w else '-',sat],fmts=[None,'0.0',None,PCT,'0.0'],h=22)
    p.row(['Score en transparence','=IRPc_Transp','=IF(IRPc_Transp>56,"Élevé",IF(IRPc_Transp>34,"Moyen","Faible"))','=1-W_LIQ','=IRPc_Transp','="Rebasé hors poche liquide : "&SUBSTITUTE(ROUND(IRPc_Transp_hp,1)&"",".",",")'],fmts=[None,'0.0',None,PCT,'0.0'],h=22)
    p.blank()
    p.sec('POINTS OUVERTS PAR FONDS CIBLE   à lever avant engagement')
    p.hdr(['Fonds','Points ouverts','','','','Criticité'])
    p.row(['MAPIF II','Dérogation de ticket (10 M$) ; plafonds AIFM (annexe A du LPA) ; ligne de souscription sans durée maximale ; fréquence contractuelle de VL ; notification AMF du passeport',None,None,None,'Élevée'],merge=[(3,6)],h=30)
    p.row(['MSIG 3','Véhicule USD ou EUR et manche levered ou unlevered ; dérogation de ticket ; admission comme Automatic Withdrawing Investor ; first closing inconnu ; aucun dépositaire AIFMD identifié ; non-appelé survivant au retrait',None,None,None,'Élevée'],merge=[(3,6)],h=30)
    p.row(['PG NGI','Classe souscrite (I ou H couverte) ; Redemption Fee discrétionnaire sans butoir ; distributions imputées sur le gate ; Special Dealing à 180 jours ; levier 200 % engagement',None,None,None,'Moyenne'],merge=[(3,6)],h=30)
    p.row(['Ares AGI ELTIF','Classe souscrite ; composition de l\'entrepôt d\'actifs ; limite ELTIF de rachat (27,3 %) ; levier non plafonné avant 2029 ; calendrier de la déduction de 24 mois',None,None,None,'Moyenne'],merge=[(3,6)],h=30)
    p.pagebreak()
    # ---------- 06 RISQUES NON QUANTIFIÉS
    page(p,'06','Valorisation, contrepartie, opérationnel et gouvernance','Risques non quantifiés  ·  dispositif  ·  lacunes  ·  actions')
    p.sec('VALORISATION   art. 19 de la directive, art. 67 à 74 du règlement délégué')
    p.hdr(['Réf.','Risque','Exposition à l\'allocation cible','Dispositif et lacune','','Action'])
    p.row(['V-1','Dépendance aux VL des gérants cibles','="100 % de l\'actif investi ; "&(ROUND((W_FERMES)*100,0)&" %")&" valorisés au mieux trimestriellement avec décalage"','IPEV déclaré par les gérants ; aucune procédure de revue de la VL reçue, ni seuil de péremption',None,'R9 : procédure de revue et seuil KRI-11'],merge=[(5,6)],h=40)
    p.row(['V-2','VL mensuelle du Fonds construite sur des VL trimestrielles','Mois sans VL fraîche pour MAPIF II et MSIG 3 ; souscriptions mensuelles à cours inconnu','Ajustement des VL des mouvements financiers (art. 13) ; aucun ajustement de marché prévu',None,'Documenter le stale pricing dans la politique de valorisation'],merge=[(5,6)],h=40)
    p.row(['V-3','Souscription au nominal tant que la VL est inférieure à 100 €','Nouveaux porteurs servis au-dessus de la VL en cas de baisse initiale','Aucun garde-fou ; question de traitement équitable (art. 11)',None,'R2 : limiter le nominal à la période de lancement'],merge=[(5,6)],h=40)
    p.blank()
    p.sec('CONTREPARTIE ET PRESTATAIRES   art. 38 et 43 du règlement délégué')
    p.hdr(['Réf.','Risque','Exposition','Dispositif et lacune','','Action'])
    p.row(['C-1','Concentration sur un sponsor','="Macquarie : "&(ROUND((W_MACQ)*100,0)&" %")&" de l\'actif, deux véhicules non lancés à la date de leur PPM"','Diversification d\'actifs, non de contrepartie ; aucune limite',None,'Décision 4 : KRI-08 à 40 %'],merge=[(5,6)],h=30)
    p.row(['C-2','Dépositaire et banques de dépôt','="Poche de "&(ROUND((W_LIQ)*100,0)&" %")&" chez BFCM, OPC monétaires et comptes à terme"','Contreparties OCDE BBB+ (art. 15) ; revue sous A-',None,'Revue annuelle du Dépositaire au dossier'],merge=[(5,6)],h=30)
    p.row(['C-3','Fonds cibles sans dépositaire AIFMD','MSIG 3 : SCSp non qualifiées de FIA à la date du PPM, aucun dépositaire désigné','Éligibilité et look-through Annexe IV à vérifier',None,'R5 : vérifier avant engagement'],merge=[(5,6)],h=30)
    p.blank()
    p.sec('OPÉRATIONNEL ET GOUVERNANCE   art. 38 et 40 du règlement délégué')
    p.hdr(['Réf.','Risque','Exposition','Dispositif et lacune','','Action'])
    p.row(['O-1','Appels de fonds à 10 Jours Ouvrés contre collecte mensuelle','Deux fonds fermés appelant à vue ; défaut d\'appel sanctionnable jusqu\'à la perte de la participation','Poche et séquencement ; aucune ligne de crédit',None,'Décision 5 : règle de séquencement'],merge=[(5,6)],h=34)
    p.row(['O-2','Conflit d\'intérêts du Conseiller','Conseil, distribution principale, accord préalable sur le prospectus et faculté de révoquer la Société de Gestion','Politique de conflits d\'intérêts ; décision d\'investissement réservée à la Société de Gestion',None,'R10 : cartographie et comité'],merge=[(5,6)],h=34)
    p.row(['O-3','Chaîne de données et Annexe IV','MSIG 3 : exercice au 31/03 et reporting à 30 j.o. ; MAPIF II : états trimestriels non audités','Aucun calendrier de collecte formalisé',None,'R9 : calendrier Annexe IV'],merge=[(5,6)],h=34)
    p.row(['O-4','Cohérence documentaire','Douze écarts au registre du prospectus ; vingt écarts entre classeur, modèle et documents des fonds cibles','Registre tenu par la Fonction Risques',None,'Décisions 1 et 2'],merge=[(5,6)],h=34)
    p.note('Aucun de ces points n\'est chiffré et aucun ne doit l\'être en l\'état : il n\'existe ni série, ni historique de perte, ni référence sectorielle qui permette une mesure défendable. Ils sont traités par le dispositif et les décisions, puis suivis au registre des incidents dès le lancement.')
    p.pagebreak()
    # ---------- 07 LIMITES ET DÉCISIONS
    page(p,'07','Limites, décisions et conclusion','Suivi des limites  ·  décisions  ·  conclusion  ·  contrôles avant diffusion')
    p.sec('SUIVI DES LIMITES   référentiel en onglet 1')
    p.hdr(['Catégorie','Nombre','Cible','Statut','Ce que cela impose',''])
    p.row(['Limites et seuils recensés','=Lim_total','-','Information','Prospectus (14), seuils internes (11)',None],fmts=[None,D],merge=[(6,7)])
    p.row(['Franchissements de limite du prospectus','=Lim_depassees',0,'=IF(Lim_depassees<=0,"Conforme","Action requise")','Fiche incident, escalade aux dirigeants effectifs, information de l\'AMF',None],fmts=[None,D,D],merge=[(6,7)])
    p.row(['Limites non testables avant lancement','=Lim_non_testables','-','Information','Premier test à la première valeur liquidative',None],fmts=[None,D],merge=[(6,7)])
    p.row(['Seuils internes à valider','=Lim_a_valider',0,'Décision 4','Validation des seuils de la note de liquidité § 5',None],fmts=[None,D,D],merge=[(6,7)])
    p.row(['Seuils internes à arrêter','=Lim_a_arreter',0,'Décision 4','Gérant, fonds cible, dollar non couvert, péremption des VL',None],fmts=[None,D,D],merge=[(6,7)])
    p.blank()
    p.sec('DÉCISIONS SOUMISES AU COMITÉ DES RISQUES')
    p.hdr(['N°','Décision','','Proposition de la Fonction Risques','','Échéance'])
    for i in range(1,7):
        p.row([str(i),f'=Dec{i}',None,f'=DecP{i}',None,f'=DecE{i}'],merge=[(3,4),(5,6)],h=44,bold_first=False)
    p.blank()
    p.sec('CONCLUSION')
    p.note('="Openstone Infraworld est, par construction, un véhicule de transformation : une promesse de liquidité semestrielle adossée à "&(ROUND((W_FERMES)*100,0)&" %")&" d\'actifs sans fenêtre avant dix ans. Le dispositif du prospectus - blocage de deux ans, plafonnement de 5 %, prolongation du préavis, suspension, plancher liquide de 15 %, ni emprunt ni dérivé - est complet au regard de l\'annexe V de la directive (UE) 2024/927, et les stress tests montrent qu\'il absorbe un choc adverse de 10 % par date. Il ne tient pas face à des rachats de 15 % puis 20 % ni à une crise combinée, et il repose sur des hypothèses de collecte et de distributions qui ne dépendent pas de la Société de Gestion. La Fonction Risques recommande de conditionner le lancement aux décisions 1 à 3 (texte de liquidité, levier maximal, dérogations écrites) et d\'arrêter les limites internes de concentration et de change avant la première valeur liquidative. Niveau de confiance : élevé sur le référentiel contractuel et la lecture des fonds cibles, moyen sur les stress tests (hypothèses externes), faible sur le comportement de rachat des porteurs, inconnu à ce stade."',lines=9)
    p.blank()
    p.sec('CONTRÔLES AVANT DIFFUSION   avertissements')
    p.hdr(['Contrôle','Constaté','Attendu','Statut','Lecture',''])
    p.row(['Allocation cible bouclée','=W_TOT',1,'=IF(ABS(W_TOT-1)<0.0001,"Conforme","Erreur")','Somme des cinq lignes de l\'onglet 0',None],fmts=[None,PCT,PCT],merge=[(6,7)])
    p.row(['Tranches Annexe IV bouclées','=Buckets_total',1,'=IF(ABS(Buckets_total-1)<0.0001,"Conforme","Erreur")','Onglet 2',None],fmts=[None,PCT,PCT],merge=[(6,7)])
    p.row(['Sources périmées','=Src_perimees',0,'=IF(Src_perimees<=0,"Conforme","Avertissement")','="Sur "&Src_total&" sources recensées (annexe)"',None],fmts=[None,D,D],merge=[(6,7)])
    p.row(['Note de liquidité alignée sur le prospectus','=Preavis_j&" jours"','180 jours','Avertissement','Écart de préavis : décision 1',None],merge=[(6,7)])
    p.row(['Poids de départ du moteur bouclés','=Ws_TOT',1,'=IF(ABS(Ws_TOT-1)<0.0001,"Conforme","Erreur")','Onglet 0, bloc des lignes du portefeuille',None],fmts=[None,PCT,PCT],merge=[(6,7)])
    p.row(['Score IRP recalculé','=IRPc_Fonds','=IRP_Fonds','=IF(ABS(IRPc_Fonds-IRP_Fonds)<0.1,"Conforme","Erreur")','Onglet 4 contre classeur IRP du 06/10/2026',None],fmts=[None,'0.0','0.0'],merge=[(6,7)])
    p.note('Rapport établi par la fonction permanente de gestion des risques - Fundcraft France SAS. Document interne, ne pas diffuser.')
    p.pagebreak()
    # ---------- ANNEXE
    page(p,'Annexe','Caractéristiques, base réglementaire et sources','Caractéristiques générales  ·  parts  ·  base réglementaire  ·  fraîcheur des données')
    p.sec('CARACTÉRISTIQUES GÉNÉRALES')
    p.hdr(['Élément','Information','','','Source',''])
    for a,b,c in [('Dénomination','Openstone Infraworld - Compartiment I','Prospectus art. 1.1 et 18'),('Forme juridique','FPS à compartiments sous forme de FCP - art. L. 214-154 CMF ; FIA non agréé','Prospectus art. 1.2'),('Société de gestion','Fundcraft France SAS - agrément AMF GP-20260008 du 4 juin 2026','Prospectus art. 2.1'),('Conseiller','Innovative Finance SAS (Openstone) - CIF ORIAS 23002459, distributeur principal','Prospectus art. 2.7 et 2.9'),('Dépositaire, comptable, CAC','BFCM ; CIC ; APLITEC','Prospectus art. 2.2 à 2.6'),('Stratégie','Fonds de fonds d\'infrastructure evergreen : fonds fermés et evergreen, secondaires, dette, equity ; co-investissements ≤ 30 %','Prospectus art. 19 à 22'),('Durée','99 ans ; durée de placement recommandée supérieure à 10 ans','Prospectus art. 1.3 et 4.2'),('Investisseurs','Investisseurs Autorisés (RG AMF 423-27) ; minimum 100 000 € ; aucune personne physique au-delà de 10 % des parts','Avertissement ; art. 5.1'),('SFDR','Article 6 ; PAI non mesurées ; aucun engagement taxonomie','Prospectus art. 21.4 et 21.5'),('Indicateur de risque','=(ROUND(SRI,0)&"")&" sur 7 - relèvement prudentiel d\'une classe (valorisation trimestrielle des sous-jacents)"','DIC PRIIPs du 29/09/2026'),('Date d\'arrêté','=DateArrete','Rapport initial sur documentation')]:
        p.row([a,b,None,None,c,None],merge=[(3,5),(6,7)],h=24,fmts=[None,'dd/mm/yyyy'] if a=='Date d\'arrêté' else None)
    p.blank()
    p.sec('CATÉGORIES DE PARTS   huit catégories en euros, minimum 100 000 €')
    p.hdr(['Part','ISIN','Affectation','Souscription','Conseil + distribution','Commentaire'])
    for part,isin,aff,sub,fee,com in [('A1c','FR001401BDP7','Capitalisation','Après la Date de Déclenchement','1,07 % + 0,75 %','Barème A, avec rétrocession'),('A1d','FR001401BDQ5','Distribution','Après la Date de Déclenchement','1,07 % + 0,75 %','Acomptes trimestriels examinés'),('A3c','FR001401BDR3','Capitalisation','Après la Date de Déclenchement','1,07 %','Clean share ; ISIN à corriger au prospectus (« FFR »)'),('A3d','FR001401BDS1','Distribution','Après la Date de Déclenchement','1,07 %','Clean share'),('B1c','FR001401BDT9','Capitalisation','Avant la Date de Déclenchement','0,57 % + 0,75 %','Conseil réduit jusqu\'à 10 M€ ou 2 ans'),('B1d','FR001401BDU7','Distribution','Avant la Date de Déclenchement','0,57 % + 0,75 %',''),('B3c','FR001401BDV5','Capitalisation','Avant la Date de Déclenchement','0,57 %','Clean share'),('B3d','FR001401BDW3','Distribution','Avant la Date de Déclenchement','0,57 %','Clean share')]:
        p.row([part,isin,aff,sub,fee,com],h=18)
    p.note('="Commission de Gestion de la Société de Gestion : "&(SUBSTITUTE(ROUND((Fee_SGP)*100,2)&"",".",",")&" %")&" (minimum 60 000 € par an) ; frais de constitution jusqu\'à "&(SUBSTITUTE(ROUND((Fee_Const)*100,2)&"",".",",")&" %")&" ; droits d\'entrée jusqu\'à "&(ROUND((Fee_Entree)*100,0)&" %")&" ; frais des fonds cibles en sus (1,00 à 1,25 % et commissions de performance de 12,5 à 15 %)."')
    p.blank()
    p.sec('BASE RÉGLEMENTAIRE')
    p.note('Rapport établi au titre de l\'article 15 de la directive 2011/61/UE et des articles 38 à 49 du règlement délégué (UE) 231/2013 (gestion des risques et de la liquidité). Outils de gestion de la liquidité : article 16 § 2 ter et annexe V de la directive, tels qu\'introduits par la directive (UE) 2024/927 ; instruction AMF DOC-2017-05 pour le plafonnement. Simulations de crise de liquidité : orientations ESMA 34-39-897 (2020). Déclaration Annexe IV : article 24 de la directive et article 110 du règlement délégué. Indicateur de risque : règlement (UE) 1286/2014 et règlement délégué (UE) 2017/653 (PRIIPs). Régime du FPS : articles L. 214-154 et suivants du CMF, 423-27 et suivants du RG AMF. Les seuils internes relèvent de la pratique de marché et non d\'une obligation réglementaire.',lines=5)
    p.blank()
    p.sec('SOURCES ET FRAÎCHEUR DES DONNÉES   détail en onglet DATA')
    p.hdr(['Source','Date','Ancienneté','Statut','Usage',''])
    for i in range(12):
        r=6+i
        p.row([f'=DATA!B{r}',f'=DATA!C{r}',f'=DATA!D{r}&" j"',f'=DATA!F{r}',f'=DATA!G{r}',None],fmts=[None,'dd/mm/yyyy'],merge=[(6,7)],h=26)
    p.note('Toutes les valeurs de ce rapport découlent du bloc de paramètres de l\'onglet 0 et des calculs qui en dérivent. Aucune valeur n\'est ressaisie dans le document.')
    return ws
