# -*- coding: utf-8 -*-
from lib import *
from params import name
NB=' '
def build_scenarios(wb, SM):
    ws=wb.create_sheet('3 Scénarios')
    S=Sheet(ws,'STRESS TESTS DE LIQUIDITÉ - synthèse des scénarios calculés dans l\'onglet 3b et verdicts',
            'Projection trimestrielle sur 36 mois à partir de la date de départ, actif net cible, allocation cible, plafond et durée du plafonnement du prospectus du 30/09/2026. Verdicts sur les seuils de la note de liquidité (onglet 0).',
            {'A':2,'B':46,'C':11,'D':11,'E':12,'F':11,'G':12,'H':12,'I':11,'J':11,'K':12,'L':10,'M':52})
    S.hdr(['Scénario','Dates plaf.','File max','Résorption (mois)','Poche min','Couv. 12 m min','Sur-eng. max','Change max','Cession (M€)','Suspension','Verdict','Lecture'])
    def verdict(r):
        return (f'=IF(OR(D{r}>S_file_rouge,F{r}<S_poche_rouge,G{r}<S_couv_rouge,E{r}>S_resorption_rouge,K{r}<>"-"),"Rouge",'
                f'IF(OR(C{r}>0,F{r}<Liq_min,G{r}<S_couv_vert,H{r}>S_sureng_orange),"Orange","Vert"))')
    FM=[None,'0','0.0 %','0','0.0 %','0.00','0 %','0 %','0.00',None,None]
    def srow(i,h=30):
        m=SM[i]; r=S.r
        S.row([m['nom'],m['dates'],m['file'],m['resorb'],m['poche'],m['couv'],m['sureng'],m['change'],m['cession'],m['susp'],verdict(r),m['lecture']],fmts=FM,h=h)
        return r
    f=S.r
    rows=[srow(i) for i in range(9)]
    l=S.r-1
    r=S.row(['Scénarios verts',f'=COUNTIF(L{f}:L{l},"Vert")']+['']*10,fmts=[None,'0']); name(wb,ws,'N_vert',f'C{r}')
    r=S.row(['Scénarios orange',f'=COUNTIF(L{f}:L{l},"Orange")']+['']*10,fmts=[None,'0']); name(wb,ws,'N_orange',f'C{r}')
    r=S.row(['Scénarios rouges',f'=COUNTIF(L{f}:L{l},"Rouge")']+['']*10,fmts=[None,'0'],key_cols=(1,)); name(wb,ws,'N_rouge',f'C{r}')
    for i,nm in enumerate(['Sc_base','Sc_adverse','Sc_extreme','Sc_top5','Sc_distrib','Sc_appels','Sc_valo','Sc_combine','Sc_prorata']):
        name(wb,ws,nm,f'L{rows[i]}')
    name(wb,ws,'Sc_combine_cession',f'J{rows[7]}'); name(wb,ws,'Sc_extreme_file',f'D{rows[2]}'); name(wb,ws,'Sc_adverse_poche',f'F{rows[1]}'); name(wb,ws,'Sc_adverse_resorption',f'E{rows[1]}')
    name(wb,ws,'Sc_combine_gel',f"'3 Scénarios'!$M${rows[7]}".split('!')[1].replace('$',''))
    # ressources et couverture de départ (scénario de base)
    r=S.row(['Ressources à douze mois / plafond annuel, au départ (scénario de base)',f"={SM[0]['res0']}/(2*Gate*AN_cible)"]+['']*10,fmts=[None,'0.00']); name(wb,ws,'Eng_Res0',f'C{r}')
    r=S.row(['Couverture de liquidité à douze mois au départ, nette des appels, frais et acomptes',f"={SM[0]['couv0']}"]+['']*10,fmts=[None,'0.00']); name(wb,ws,'Eng_Couv0',f'C{r}')
    r=S.row(['Premier triple gel (crise combinée)',SM[7]['gel']]+['']*10); name(wb,ws,'Gel_combine',f'C{r}')
    S.ncols=12
    S.note('Résorption « 99 » : file non résorbée à l\'horizon de 36 mois. Scénarios de la note de liquidité § 6, calculés dans l\'onglet 3b ; le scénario de concentration applique bien la sortie des cinq premiers porteurs (le classeur du 24/09/2026 laissait cette demande à zéro).',12)
    S.blank()
    S.block('TESTS INVERSÉS  ·  calculés dans l\'onglet 3b')
    # rachats persistants (gate appliqué : 9..13 ; sans gate : 14..18)
    S.hdr(['Rachats persistants par date (% AN)','Dates plaf.','Suspension','Poche min (plafond appliqué)','Cession sans plafond (M€)','Décote (M€)','Poche min (sans plafond)','aide','','','Verdict','Lecture'])
    f=S.r
    for k,x in enumerate([0.05,0.075,0.10,0.125,0.15]):
        g=SM[9+k]; n=SM[14+k]; r=S.r
        S.row([x,g['dates'],g['susp'],g['poche'],n['cession'],f'={n["cession"][1:]}*Decote',n['poche'],f'=IF(OR(D{r}<>"-",F{r}>0.000001),B{r},9)','','',f'=IF(OR(D{r}<>"-",E{r}<S_poche_rouge),"Rouge",IF(C{r}>0,"Orange","Vert"))',
               'Demandes égales au plafond : servies sans plafonnement, elles érodent la poche sous 5 %.' if k==0 else ''],fmts=['0.0 %','0',None,'0.0 %','0.00','0.00','0.0 %','0.000'],h=24)
    l=S.r-1
    r=S.row(['Point de rupture : premier taux conduisant à la suspension ou à des cessions',f'=MIN(I{f}:I{l})']+['']*10,fmts=[None,'0.0 %'],key_cols=(1,)); name(wb,ws,'Rupture_rachats',f'C{r}')
    r=S.row(['Cessions secondaires sans plafonnement à ce taux (M€)',f'=SUMPRODUCT((ABS(B{f}:B{l}-C{r})<0.00001)*F{f}:F{l})']+['']*10,fmts=[None,'0.00']); name(wb,ws,'Rupture_cession',f'C{r}')
    r=S.row(['Poche minimale avec des rachats égaux au plafond (5 %)',f'=E{f}']+['']*10,fmts=[None,'0.0 %']); name(wb,ws,'Persist5_poche',f'C{r}')
    S.blank()
    S.hdr(['Appels de fonds à douze mois (% AN)','Non-appelé départ','Poche min','Couv. 12 m min','Cession (M€)','aide','','','','','Verdict','Lecture'])
    f=S.r
    for k,(x,na) in enumerate([(0.075,0.15),(0.125,0.25),(0.175,0.35),(0.225,0.45),(0.45,0.45)]):
        m=SM[19+k]; r=S.r
        S.row([x,na,m['poche'],m['couv'],m['cession'],f'=IF(D{r}<Poche_min_exec,B{r},9)','','','','',f'=IF(OR(D{r}<S_poche_rouge,E{r}<S_couv_rouge),"Rouge",IF(OR(D{r}<Liq_min,E{r}<S_couv_vert),"Orange","Vert"))',''],fmts=['0.0 %','0 %','0.0 %','0.00','0.00','0.000'],h=22)
    l=S.r-1
    r=S.row(['Point de rupture : premier niveau d\'appels qui fait passer la poche sous 10 % de l\'actif net',f'=MIN(G{f}:G{l})']+['']*10,fmts=[None,'0.0 %'],key_cols=(1,)); name(wb,ws,'Rupture_appels',f'C{r}')
    S.blank()
    S.hdr(['Sortie d\'un porteur unique (% AN)','Dates plaf.','File max','Résorption (mois)','Poche min','Suspension','Cession (M€)','aide','','','Verdict','Lecture'])
    f=S.r
    for k,x in enumerate([0.10,0.15,0.20,0.25]):
        m=SM[24+k]; r=S.r
        S.row([x,m['dates'],m['file'],m['resorb'],m['poche'],m['susp'],m['cession'],f'=IF(AND(G{r}="-",H{r}<0.000001),B{r},0)','','',f'=IF(OR(G{r}<>"-",E{r}>S_resorption_rouge),"Rouge",IF(C{r}>0,"Orange","Vert"))',''],fmts=['0 %','0','0.0 %','0','0.0 %',None,'0.00','0.000'],h=22)
    l=S.r-1
    r=S.row(['Sortie maximale d\'un porteur unique sans suspension ni cession',f'=MAX(I{f}:I{l})']+['']*10,fmts=[None,'0 %'],key_cols=(1,)); name(wb,ws,'Porteur_max',f'C{r}')
    S.blank()
    S.note('Hypothèses (onglet 0) : poche de 15 % au départ, non-appelé 15 % tiré sur trois ans, rachats evergreen à 5 % par trimestre, décote de cession 10 %, frais 2 % par an, acomptes 5 % sur 20 % de parts de distribution, toutes les parts rachetables dès le départ (prudent). Le préavis de 91 jours ne modifie pas les flux : il réduit à trois mois le délai d\'information de la Société de Gestion avant chaque date.')
