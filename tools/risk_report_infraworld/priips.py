# -*- coding: utf-8 -*-
"""Entrées du classeur KID PRIIPs v2.3 (22/09/2026) : proxies, volatilités, SRI / VEV, scénarios, coûts."""
from lib import *
from params import name
NB=' '
CLASSES=['A1c','A1d','A3c','A3d','B1c','B1d','B3c','B3d']
# (frais de gestion du Fonds, RIY 2 ans, RIY 10 ans, coûts 2 ans €, coûts 10 ans €, tensions 2 ans €, tensions ann., tensions 10 ans €, ann., défav 10 ans €, ann., interm 10 ans €, ann., fav 10 ans €, ann.)
KID={
 'A1c':(0.02,0.07808,0.056,1721.56,15001.29,7434.16,-0.13778,4237.80,-0.08227,10792.84,0.00766,22979.51,0.08676,23739.72,0.0903),
 'A1d':(0.02,0.0764,0.0495,1676.80,11256.62,7426.45,-0.13823,4215.86,-0.08275,10664.38,0.00645,19685.48,0.07008,20247.67,0.07309),
 'A3c':(0.0125,0.06973,0.04748,1543.39,13133.27,7551.29,-0.13102,4582.29,-0.07507,11670.19,0.01557,24847.53,0.09529,25669.53,0.09886),
 'A3d':(0.0125,0.06841,0.04223,1506.80,9876.24,7543.52,-0.13147,4558.75,-0.07555,11409.51,0.01327,21065.86,0.07735,21671.14,0.08041),
 'B1c':(0.02,0.07251,0.05487,1602.93,14760.24,7512.15,-0.13327,4282.25,-0.08131,10906.05,0.00871,23220.56,0.0879,23988.74,0.09144),
 'B1d':(0.02,0.07107,0.04854,1563.61,11078.16,7504.39,-0.13372,4260.11,-0.08179,10763.82,0.00739,19863.95,0.07104,20431.70,0.07406),
 'B3c':(0.0125,0.06417,0.04634,1423.83,12874.66,7629.88,-0.12651,4629.98,-0.07411,11791.66,0.01662,25106.14,0.09642,25936.69,0.1),
 'B3d':(0.0125,0.06308,0.04126,1392.75,9685.60,7622.07,-0.12696,4606.23,-0.07459,11508.28,0.01415,21256.51,0.07832,21867.74,0.08139)}

def build_params_priips(wb):
    ws=wb['0 Paramètres']; r=ws.max_row+2
    def block(text):
        nonlocal r
        c=ws.cell(r,2,text); c.font=F(True,WHITE); c.fill=fill(NAVY)
        for col in range(3,8): ws.cell(r,col).fill=fill(NAVY)
        r+=1
    def row(key,label,val,unit,src,fiab,res='',fmt=None,derived=False):
        nonlocal r
        for i,v in enumerate([label,val,unit,src,fiab,res]):
            c=ws.cell(r,2+i,v); c.font=F(False,INK); c.alignment=al('left',v='top')
            if i==1 and fmt: c.number_format=fmt
            if i==1 and not derived: c.fill=fill('FFF8F2')
        for col in range(2,8): ws.cell(r,col).border=B_THIN
        name(wb,ws,key,f'C{r}'); r+=1
    S='Classeur KID PRIIPs Openstone Infraworld v2.3 (DIC du 22/09/2026)'
    block('ENTRÉES DU DIC PRIIPs  ·  '+S)
    row('Liq_mon_w','Poche liquide : part en OPC monétaires',0.05,'% AN',S+', Parametres (allocation cible et proxies)','Moyenne','Au moins 5 % en monétaire et dépôts selon la règle interne.','0 %')
    row('Liq_hy_w','Poche liquide : part en dette à haut rendement (OPC / ETF)',0.10,'% AN',S+', Parametres','Moyenne','Obligataire limité à 10 % de l\'actif selon la règle interne ; sans contrainte de notation au prospectus.','0 %')
    row('TauxMon','Taux monétaire',0.025,'par an',S+', Parametres C11','Moyenne','','0.0 %')
    row('Rend_HY','Rendement net cible de la poche à haut rendement',0.06,'par an',S+', Fonds cibles : fourchette 5 à 7 %','Faible','Proxy Morningstar LSTA, rendement brut observé 4,9 %.','0.0 %')
    row('Vol_Infra','Volatilité annualisée désmoothée - Preqin Infrastructure',0.0334,'par an',S+', DATA : Geltner ordre 1, correction de fréquence','Moyenne','Volatilité brute 3,0 % ; autocorrélation 0,12.','0.00 %')
    row('Vol_VA','Volatilité désmoothée - Preqin Infrastructure Value Added',0.0427,'par an',S+', DATA','Moyenne','','0.00 %')
    row('Vol_DL','Volatilité désmoothée - Preqin Private Debt Direct Lending',0.0609,'par an',S+', DATA','Moyenne','','0.00 %')
    row('Vol_HY','Volatilité - Morningstar LSTA (coté)',0.0568,'par an',S+', DATA','Élevée','Indice coté, non lissé.','0.00 %')
    row('Vol_USD','Volatilité annualisée EUR/USD',0.0782,'par an',S+', DATA X5','Élevée','','0.00 %')
    row('Z99','Quantile à 99 % de la loi normale',2.326,'écart-type','Convention','Élevée','Choc à 99 % sur un an = 2,326 × volatilité ; hypothèse de normalité, sans queue épaisse.','0.000')
    row('Choc99_Infra','Choc à 99 % sur un an - infrastructure (equity)','=Z99*Vol_Infra','%','dérivé','-','Le choc retenu dans les stress tests (20 %) est 2,6 fois supérieur.','0.0 %',derived=True)
    row('Choc99_VA','Choc à 99 % sur un an - infrastructure value added','=Z99*Vol_VA','%','dérivé','-','','0.0 %',derived=True)
    row('Choc99_DL','Choc à 99 % sur un an - dette privée','=Z99*Vol_DL','%','dérivé','-','Le choc dette retenu (10 %) est inférieur : voir réserve.','0.0 %',derived=True)
    row('Choc99_HY','Choc à 99 % sur un an - haut rendement coté','=Z99*Vol_HY','%','dérivé','-','Appliqué à la poche HY dans le scénario de valorisation (onglet 3b).','0.0 %',derived=True)
    row('Choc99_USD','Choc à 99 % sur un an - EUR/USD','=Z99*Vol_USD','%','dérivé','-','Le choc dollar retenu (15 %) est inférieur.','0.0 %',derived=True)
    row('VEV_priv','VEV du proxy composite privé (désmoothé)',0.0502,'%',S+', Scenarios C24','Moyenne','VaR à 10 ans dans l\'espace des rendements : - 32,3 %.','0.00 %')
    row('VEV_cote','VEV du contrôle coté (S&P Global Infrastructure)',0.1088,'%',S+', Scenarios D24','Élevée','Retenue comme VEV du produit (annexe IV point 16 du RD 2017/653).','0.00 %')
    row('MRM','Classe de risque de marché (MRM) retenue',3,'1-7',S+', Scenarios C27','Élevée','Relèvement de + 1 pour valorisation trimestrielle des sous-jacents : MRM 4.')
    row('SRI_mec','SRI mécanique (MRM × CRM 1)',3,'1-7',S+', Scenarios C29','Élevée','Publié 4 après relèvement prudentiel ; aucun relèvement pour liquidité.')
    row('Stress_10','Facteur brut du scénario de tensions à 10 ans',0.700,'facteur',S+', Scenarios C17 (Cornish-Fisher, σ stressée P95)','Moyenne','Perte brute de 30 % avant coûts sur l\'horizon recommandé.','0.000')
    row('Stress_2','Facteur brut du scénario de tensions à 2 ans',0.857,'facteur',S+', Scenarios F17','Moyenne','Perte brute de 14 % à la fin du blocage.','0.000')
    row('LT_rec','Frais récurrents des fonds cibles en transparence',0.0177,'% AN / an',S+', Fonds cibles C48 : DIC des gérants quand ils existent, sinon PPM','Moyenne','MAPIF II 1,55 %, PG NGI 2,70 %, MSIG 3 1,30 %, Ares 3,00 % (dont levier 1,25 %).','0.00 %')
    row('LT_carry','Carried interest ex ante des fonds cibles',0.0053,'% AN / an',S+', Fonds cibles C49 : excédent au-dessus du hurdle à la RHP','Faible','Estimation haute avec catch-up intégral : 0,91 %.','0.00 %')
    row('LT_txn','Coûts de transaction en transparence',0.0002,'% AN / an',S+', Fonds cibles C50','Moyenne','','0.00 %')
    row('Fee_fixe','Frais fixes annuels des prestataires (BFCM, CIC, CAC)',55192,'euros',S+', Parametres C72 : offres du 25/03/2026','Élevée','Soit 0,11 % à 50 M€ et 1,1 % à 5 M€ : la taille du Fonds commande la charge.','# ##0')
    row('Fee_Fonds_A1','Frais récurrents du Fonds, part A1, prestataires compris',0.0211,'% AN / an',S+', Fees H20','Élevée','Hors fonds cibles. Part A3 : 1,36 %.','0.00 %')
    row('Fee_Fonds_A3','Frais récurrents du Fonds, part A3, prestataires compris',0.0136,'% AN / an',S+', Fees','Élevée','','0.00 %')
    row('RIY_A1','Réduction du rendement à 10 ans, part A1c',0.056,'par an',S+', Fees D49','Élevée','7,81 % à 2 ans (droits d\'entrée non amortis).','0.00 %')
    row('RIY_A3','Réduction du rendement à 10 ans, part A3c',0.0475,'par an',S+', Fees F49','Élevée','6,97 % à 2 ans.','0.00 %')
    row('Rdt_brut','Rendement annuel moyen avant coûts, scénario intermédiaire à 10 ans',0.1428,'par an',S+', Fees D47','Faible','Proxy composite calibré sur les TRI nets cibles des gérants, relevés de 0,5 point.','0.00 %')
    row('KID_lancement','Date de lancement retenue par le DIC (première VL)','=DATE(2026,9,30)','date',S+', Parametres C16','Élevée','Le Fonds n\'est pas constitué au 06/10/2026 : le DIC retient une date dépassée (R11).','dd/mm/yyyy',derived=True)

def build_priips_sheet(wb):
    ws=wb.create_sheet('4b PRIIPs')
    S=Sheet(ws,'DIC PRIIPs - PROXY COMPOSITE, INDICATEUR DE RISQUE, SCÉNARIOS ET COÛTS',
            'Entrées du classeur KID PRIIPs Openstone Infraworld v2.3 (DIC du 22/09/2026), catégorie 2, proxy composite Preqin désmoothé et contrôle coté, calibré sur les TRI nets cibles. Les colonnes crème sont des saisies reprises du classeur ; les autres sont des formules.',
            {'A':2,'B':40,'C':12,'D':12,'E':12,'F':12,'G':12,'H':12,'I':12,'J':12,'K':46})
    S.block('A. PROXY COMPOSITE PAR LIGNE  ·  allocation cible')
    S.hdr(['Ligne','Poids','Proxy','TRI net cible','Vol. désmoothée','Choc 99 % 1 an','Perte sur l\'actif net (choc 99 %)','Devise','Choc stress test retenu','Perte sur l\'actif net (stress)','Lecture'])
    rows=[('MAPIF II','=W_MAPIF','Preqin Infrastructure','=Rend_MAPIF','=Vol_Infra','=Choc99_Infra','USD non couvert','=Choc_cap','Fonds fermé de secondaires : la VL suit le lissage des GP sous-jacents.'),
          ('MSIG 3','=W_MSIG','Preqin Private Debt Direct Lending','=Rend_MSIG','=Vol_DL','=Choc99_DL','USD couvert (hypothèse DIC)','=Choc_dette','Dette senior et subordonnée ; le choc 99 % de l\'indice dépasse le choc retenu.'),
          ('PG NGI','=W_PG','Preqin Infrastructure Value Added','=Rend_PG','=Vol_VA','=Choc99_VA','EUR couvert (classe H, hypothèse DIC)','=Choc_cap','Evergreen value add, levier jusqu\'à 40 % de la VL.'),
          ('Ares AGI','=W_ARES','60 % Infra, 20 % Direct Lending, 10 % monétaire, 10 % HY','=Rend_ARES','=0.6*Vol_Infra+0.2*Vol_DL+0.1*Vol_HY','=Z99*F{r}','EUR','=Cap_ARES*Choc_cap+(1-Cap_ARES)*Choc_dette','Mix equity et dette ; volatilité composite sans effet de diversification (prudent).'),
          ('Actifs liquides : monétaire','=Liq_mon_w','Taux monétaire','=TauxMon',0,0,'EUR',0,'Sans risque de marché.'),
          ('Actifs liquides : haut rendement','=Liq_hy_w','Morningstar LSTA (coté)','=Rend_HY','=Vol_HY','=Choc99_HY','EUR couvert','=Choc_dette','Seule ligne cotée ; poche limitée à 10 % par la règle interne.')]
    f=S.r
    for t in rows:
        r=S.r
        vals=[t[0],t[1],t[2],t[3],t[4],(t[5].replace('{r}',str(r)) if isinstance(t[5],str) else t[5]),f'=C{r}*G{r}',t[6],t[7],f'=C{r}*J{r}',t[8]]
        S.row(vals,fmts=[None,'0 %',None,'0.0 %','0.00 %','0.0 %','0.0 %',None,'0.0 %','0.0 %'],h=30)
    l=S.r-1
    r=S.row(['Total / pondéré',f'=SUM(C{f}:C{l})','',f'=SUMPRODUCT(C{f}:C{l},E{f}:E{l})',f'=SUMPRODUCT(C{f}:C{l},F{f}:F{l})','',f'=SUM(H{f}:H{l})','','',f'=SUM(K{f}:K{l})','Somme des pertes par ligne : borne haute, corrélations supposées égales à 1.'],fmts=[None,'0 %',None,'0.0 %','0.00 %',None,'0.0 %',None,None,'0.0 %'],key_cols=(6,9))
    name(wb,ws,'Rdt_cible_pond',f'E{r}'); name(wb,ws,'Vol_pond',f'F{r}'); name(wb,ws,'Perte99_AN',f'H{r}'); name(wb,ws,'Perte_stress_AN',f'K{r}')
    r=S.row(['Perte de change sur l\'actif net - dollar - 15 % (hypothèse basse de change)','=W_USD_BAS*Choc_usd','','','','','','','','','MAPIF II seul en dollars non couverts.'],fmts=[None,'0.0 %']); name(wb,ws,'Perte_USD_bas',f'C{r}')
    r=S.row(['Perte de change sur l\'actif net - dollar - 15 % (hypothèse haute de change)','=W_USD_HAUT*Choc_usd','','','','','','','','','MAPIF II, MSIG 3 USD et PG NGI non couverts.'],fmts=[None,'0.0 %']); name(wb,ws,'Perte_USD_haut',f'C{r}')
    r=S.row(['Perte combinée valorisation et change (hypothèse basse)','=Perte_stress_AN+Perte_USD_bas','','','','','','','','','Scénario « valorisation et dollar » du moteur, avant effet de dénominateur.'],fmts=[None,'0.0 %'],key_cols=(1,)); name(wb,ws,'Perte_combinee',f'C{r}')
    S.blank(); S.ncols=11
    S.block('B. INDICATEUR DE RISQUE  ·  annexe II du règlement délégué (UE) 2017/653')
    S.hdr(['Mesure','Proxy privé','Contrôle coté','Retenu','','','','','','','Lecture'])
    S.row(['VEV (volatilité équivalente)','=VEV_priv','=VEV_cote','=MAX(VEV_priv,VEV_cote)','','','','','','','Le contrôle coté l\'emporte : le proxy privé désmoothé reste trop lisse pour fonder le SRI.'],fmts=[None,'0.00 %','0.00 %','0.00 %'])
    S.row(['MRM','=MRM','=MRM','=MRM+1','','','','','','','Relèvement de + 1 pour la valorisation trimestrielle et décalée des fonds fermés.'],fmts=[None,'0','0','0'])
    S.row(['SRI','=SRI_mec','=SRI_mec','=SRI','','','','','','','Publié 4 sur 7 ; aucun relèvement pour le risque de liquidité, traité dans les mentions.'],fmts=[None,'0','0','0'])
    S.row(['Facteur brut de tensions (Cornish-Fisher)','=Stress_2','','=Stress_10','','','','','','','2 ans puis 10 ans : perte brute de 14 % puis 30 % avant coûts.'],fmts=[None,'0.000',None,'0.000'])
    S.blank()
    S.block('C. SCÉNARIOS DE PERFORMANCE ET COÛTS PAR PART  ·  10 000 € investis')
    S.hdr(['Part','Frais du Fonds','RIY 2 ans','RIY 10 ans','Coûts 2 ans (€)','Coûts 10 ans (€)','Tensions 2 ans (€)','Tensions 10 ans (€)','Défavorable 10 ans (€)','Intermédiaire 10 ans (€)','Favorable 10 ans (€) et rendement annuel'])
    f=S.r
    for cl in CLASSES:
        k=KID[cl]; r=S.r
        S.row([cl,k[0],k[1],k[2],k[3],k[4],k[5],k[7],k[9],k[11],f'{k[13]:,.0f} € ; {k[14]*100:.1f} % par an'.replace(',',' ')],fmts=[None,'0.00 %','0.00 %','0.00 %','# ##0','# ##0','# ##0','# ##0','# ##0','# ##0'],input_cols=range(1,10),h=18)
    l=S.r-1
    name(wb,ws,'KID_A1_tens10',f'I{f}'); name(wb,ws,'KID_A1_tens2',f'H{f}'); name(wb,ws,'KID_A1_defav10',f'J{f}'); name(wb,ws,'KID_A1_interm10',f'K{f}')
    name(wb,ws,'KID_A1_cout10',f'G{f}')
    r=S.row(['Écart de RIY à 10 ans entre la part la plus chère (A1) et la moins chère (B3)',f'=MAX(E{f}:E{l})-MIN(E{f}:E{l})','','','','','','','','','Commission de distribution de 0,75 % et conseil réduit des parts B.'],fmts=[None,'0.00 %']); name(wb,ws,'RIY_ecart',f'C{r}')
    S.blank()
    S.block('D. EMPILEMENT DES COÛTS  ·  part A1, % de l\'actif net par an')
    S.hdr(['Couche','Taux','','','','','','','','','Source'])
    S.row(['Commission de Gestion (Société de Gestion)','=Fee_SGP','','','','','','','','','Prospectus art. 27.2.1'],fmts=[None,'0.00 %'])
    S.row(['Commission de Conseil (Openstone), parts A','=Fee_CIF_A','','','','','','','','','Prospectus art. 27.2.2'],fmts=[None,'0.00 %'])
    S.row(['Commission de Distribution, parts A1 et B1','=Fee_Distrib','','','','','','','','','Prospectus art. 27.2.3'],fmts=[None,'0.00 %'])
    S.row(['Prestataires et frais de fonctionnement (à 50 M€)','=Fee_fixe/AN_cible','','','','','','','','','Offres BFCM et CIC du 25/03/2026, CAC'],fmts=[None,'0.00 %'])
    r0=S.r
    S.row(['Frais récurrents des fonds cibles (transparence)','=LT_rec','','','','','','','','','DIC des gérants (PG NGI), PPM sinon ; Ares : levier 1,25 % inclus'],fmts=[None,'0.00 %'])
    S.row(['Carried interest ex ante des fonds cibles','=LT_carry','','','','','','','','','Cascade à la RHP sur les TRI cibles, excédent au-dessus du hurdle'],fmts=[None,'0.00 %'])
    S.row(['Coûts de transaction','=LT_txn','','','','','','','','','DIC des fonds cibles et poche HY'],fmts=[None,'0.00 %'])
    r=S.row(['Total annuel, part A1',f'=SUM(C{r0-4}:C{r0+2})','','','','','','','','','Vue investisseur, hors droits d\'entrée (5 % au plus, non acquis au Fonds).'],fmts=[None,'0.00 %'],key_cols=(1,)); name(wb,ws,'Cout_total_A1',f'C{r}')
    S.row(['Rendement brut du proxy moins coûts totaux (ordre de grandeur du rendement net)','=Rdt_brut-Cout_total_A1','','','','','','','','','À comparer à l\'objectif commercial de 8 à 10 % net ; le DIC donne 8,7 % en intermédiaire pour A1c.'],fmts=[None,'0.00 %'])
    S.blank()
    S.note('Lecture risque : le SRI 4 repose sur un contrôle coté, car le proxy privé désmoothé (VEV 5 %) donnerait 3 ; la Société de Gestion relève d\'une classe pour la valorisation trimestrielle. Les chocs à 99 % à un an des indices désmoothés restent inférieurs aux chocs retenus dans les stress tests pour l\'equity (7,8 % contre 20 %) mais supérieurs pour la dette privée (14,2 % contre 10 %) et le change (18,2 % contre 15 %) : le choc dette et le choc dollar sont à relever lors de la prochaine revue. Les coûts cumulés (4,4 % par an pour A1, dont 2,3 % en transparence) absorbent près du tiers du rendement brut du proxy ; ils sont la première cause d\'écart entre l\'objectif commercial et le résultat net.')
