# -*- coding: utf-8 -*-
"""Jeu de données fictif pour la version de démonstration du registre.

Aucune donnée réelle : fonds « test », montants et dates générés aléatoirement
(graine fixe) pour vérifier les formules et illustrer le tableau de bord.
"""
import random
import datetime as dt

import content as C

NOW = dt.datetime(2026, 10, 8, 9, 0)
OPEN, CLOSE = dt.time(9, 0), dt.time(18, 0)

DESCRIPTIONS = {
    "Valorisation et VL": ["Écart de VL détecté au contrôle de cohérence", "Valorisation d'un fonds cible intégrée sur une base erronée",
                           "Erreur de change dans le calcul de la VL"],
    "Passif (souscriptions, rachats, registre)": ["Ordre de souscription traité sur une mauvaise part", "Retard de mise à jour du registre des porteurs"],
    "Appels de fonds et distributions": ["Avis d'appel de fonds envoyé avec un montant erroné", "Distribution versée à la mauvaise date de valeur"],
    "Trésorerie et paiements": ["Paiement fournisseur émis en double", "Rapprochement bancaire non réalisé à J+1"],
    "Investissements et limites (ratios)": ["Dépassement passif d'une limite de concentration", "Engagement souscrit sans contrôle préalable des limites"],
    "Liquidité (gates, outils de gestion)": ["Seuil de gate mal paramétré dans l'outil de suivi", "Prévision de liquidité non mise à jour avant la fenêtre de rachat"],
    "Conformité et LCB-FT": ["Dossier KYC incomplet constaté après souscription", "Revue périodique LCB-FT non réalisée à échéance"],
    "Reporting réglementaire (Annex IV, PRIIPs, EET)": ["Fichier Annex IV rejeté par l'AMF (schéma XSD)", "KID PRIIPs publié avec un SRI non mis à jour"],
    "Reporting investisseurs": ["Reporting trimestriel envoyé avec un tableau erroné", "Relevé de position transmis en retard"],
    "Systèmes d'information et cybersécurité (TIC)": ["Indisponibilité de l'outil de gestion du passif", "Tentative d'hameçonnage ayant compromis une messagerie"],
    "Délégataires et prestataires": ["Retard de production de la VL par l'administrateur", "Prestataire n'ayant pas transmis son rapport de contrôle"],
    "Juridique et documentation des fonds": ["Version non à jour du règlement transmise à un investisseur", "Délai de notification d'une modification du règlement dépassé"],
    "Autre": ["Erreur d'archivage d'un procès-verbal", "Accès d'un collaborateur non révoqué à son départ"],
}
RESPONSABLES = ["Middle-office", "Gestion", "Risk Management", "Conformité", "Administrateur de fonds", "Agent de transfert",
                "Prestataire informatique"]
PRESTA = {"Valorisation et VL": "Administrateur de fonds", "Passif (souscriptions, rachats, registre)": "Agent de transfert",
          "Délégataires et prestataires": "Administrateur de fonds", "Systèmes d'information et cybersécurité (TIC)": "Prestataire informatique"}


def holidays():
    from build import jours_feries
    return {d for d, _ in jours_feries()}


HOL = None


def is_wd(d):
    return d.weekday() < 5 and d not in HOL


def next_open(t):
    d = t.date() + dt.timedelta(days=1)
    while not is_wd(d):
        d += dt.timedelta(days=1)
    return dt.datetime.combine(d, OPEN)


def add_bh(t, hours):
    """Ajoute des heures ouvrées (9 h - 18 h, jours ouvrés)."""
    rem = hours * 60.0
    cur = t
    while True:
        o, c = dt.datetime.combine(cur.date(), OPEN), dt.datetime.combine(cur.date(), CLOSE)
        if not is_wd(cur.date()) or cur >= c:
            cur = next_open(cur)
            continue
        if cur < o:
            cur = o
        avail = (c - cur).total_seconds() / 60
        if rem <= avail:
            res = cur + dt.timedelta(minutes=rem)
            return res.replace(second=0, microsecond=0) - dt.timedelta(minutes=res.minute % 5)
        rem -= avail
        cur = next_open(cur)


def pick(rng, items, weights):
    return rng.choices(items, weights=weights, k=1)[0]


def business_days(a, b):
    d, out = a, []
    while d <= b:
        if is_wd(d):
            out.append(d)
        d += dt.timedelta(days=1)
    return out


def generate(fonds):
    global HOL
    HOL = holidays()
    rng = random.Random(20261008)
    days = business_days(dt.date(2025, 10, 1), dt.date(2026, 10, 6))
    dom_w = [16, 10, 10, 7, 8, 5, 8, 9, 7, 8, 7, 3, 2]
    src_w = [35, 25, 10, 15, 6, 5, 3, 1]
    fonds_w = [26, 20, 24, 8, 8, 14]
    rows = []

    def base_row(enr, crit, typ="Incident"):
        dom = pick(rng, C.DOMAINES, dom_w)
        if dom.startswith("Systèmes"):
            bale = pick(rng, C.BALE, [0, 2, 0, 0, 1, 7, 1])
        else:
            bale = pick(rng, C.BALE, [1, 1, 0, 8, 0, 4, 86])
        det = enr - dt.timedelta(minutes=rng.choice([10, 20, 30, 45, 60, 90, 120]))
        if det.time() < dt.time(8, 0):
            det = enr - dt.timedelta(minutes=10)
        u = rng.random()
        if u < 0.55:
            surv = det - dt.timedelta(hours=rng.uniform(0.5, 6))
        elif u < 0.92:
            surv = det - dt.timedelta(days=rng.randint(1, 5))
        else:
            surv = det - dt.timedelta(days=rng.randint(15, 40))
        surv = surv.replace(second=0, microsecond=0) - dt.timedelta(minutes=surv.minute % 5)
        return dict(Type=typ, Fonds=pick(rng, fonds, fonds_w), Dom=dom, Bale=bale,
                    Src=pick(rng, [s for s, _ in C.SOURCES], src_w), Presta=PRESTA.get(dom),
                    Desc=rng.choice(DESCRIPTIONS[dom]), Crit=C.CRITICITES[crit - 1][0], DORA="Non",
                    Resp=rng.choice(RESPONSABLES), Surv=surv, Det=det, Enr=enr, Reo="Non", RCA="Non",
                    NotReq="Non", ActReq="Non")

    # Incidents aléatoires
    for _ in range(54):
        d = rng.choice(days)
        enr = dt.datetime.combine(d, dt.time(rng.randint(9, 17), rng.choice(range(0, 60, 5))))
        crit = pick(rng, [1, 2, 3, 4], [4, 16, 46, 34])
        r = base_row(enr, crit)
        pec_h = {1: (0.1, 1.2), 2: (0.3, 2.3), 3: (0.5, 9.5), 4: (1, 19)}[crit]
        r["PEC"] = add_bh(enr, rng.uniform(*pec_h))
        res_d = {1: (0.2, 1.2), 2: (0.8, 3.4), 3: (1, 10.5), 4: (2, 21)}[crit]
        res = add_bh(r["PEC"], rng.uniform(*res_d) * 9)
        if res > NOW or (enr.date() > dt.date(2026, 8, 20) and rng.random() < 0.35):
            res = None
        r["Res"] = res
        if res is not None:
            clo = add_bh(res, rng.uniform(2, 24) * 9)
            r["Clo"] = clo if clo <= NOW and rng.random() < 0.92 else None
        # Escalade
        if crit == 1:
            r["EscN2"] = add_bh(r["Det"], rng.uniform(0.2, 0.8))
            r["EscN3"] = add_bh(r["Det"], rng.uniform(0.6, 2.6))
        elif crit == 2:
            if rng.random() < 0.9:
                r["EscN2"] = add_bh(r["Det"], rng.uniform(0.5, 5.5))
            if rng.random() < 0.2:
                r["EscN3"] = add_bh(r["Det"], rng.uniform(3, 12))
        elif crit == 3 and rng.random() < 0.15:
            r["EscN2"] = add_bh(r["Det"], rng.uniform(2, 20))
        elif crit == 4 and rng.random() < 0.05:
            r["EscN2"] = add_bh(r["Det"], rng.uniform(4, 30))
        for k in ("EscN2", "EscN3"):
            if r.get(k) and res and r[k] > res:
                r[k] = r["PEC"]
            if r.get(k) and r[k] > NOW:
                r[k] = None
        r["Reaf"] = pick(rng, [0, 1, 2, 3], [55, 30, 10, 5])
        if res and rng.random() < 0.06:
            r["Reo"] = "Oui"
        if crit <= 2:
            r["RCA"] = "Oui" if rng.random() < 0.85 else "Non"
        elif rng.random() < 0.2:
            r["RCA"] = "Oui"
        if crit == 1 or (crit == 2 and rng.random() < 0.3):
            r["NotReq"] = "Oui"
            lim = dt.datetime.combine((r["Det"] + dt.timedelta(days=2)).date(), dt.time(18, 0))
            r["NotLim"] = lim
            late = rng.random() < 0.15
            nt = r["Det"] + dt.timedelta(hours=rng.uniform(58, 70) if late else rng.uniform(3, 30))
            r["Not"] = nt.replace(second=0, microsecond=0) if nt <= NOW else None
        if rng.random() < 0.22:
            pb = round(rng.uniform(1000, 60000), -2)
            r["PB"] = pb
            u = rng.random()
            r["Recup"] = 0 if u < 0.6 else (pb if u < 0.8 else round(pb * rng.uniform(0.2, 0.8), -2))
        if crit <= 2 or (crit == 3 and rng.random() < 0.4) or (crit == 4 and rng.random() < 0.1):
            r["ActReq"] = "Oui"
            ech = enr.date() + dt.timedelta(days=rng.randint(20, 60))
            r["Ech"] = ech
            u = rng.random()
            real = ech - dt.timedelta(days=rng.randint(0, 10)) if u < 0.82 else (ech + dt.timedelta(days=rng.randint(1, 15)) if u < 0.93 else None)
            r["Real"] = real if real and real <= NOW.date() else None
        rows.append(r)

    # Incident TIC majeur DORA conforme (septembre 2026)
    enr = dt.datetime(2026, 9, 2, 10, 40)
    r = base_row(enr, 1)
    r.update(Dom="Systèmes d'information et cybersécurité (TIC)", Bale="Interruption d'activité et dysfonctionnement des systèmes",
             Src="Alerte automatique (outil, rapprochement)", Presta="Prestataire informatique",
             Desc="Indisponibilité de l'outil de gestion du passif", DORA="Oui", Surv=dt.datetime(2026, 9, 2, 9, 50),
             Det=dt.datetime(2026, 9, 2, 10, 15), PEC=dt.datetime(2026, 9, 2, 10, 55), Res=dt.datetime(2026, 9, 3, 9, 30),
             Clo=dt.datetime(2026, 9, 17, 16, 0), EscN2=dt.datetime(2026, 9, 2, 10, 50), EscN3=dt.datetime(2026, 9, 2, 11, 30),
             Reaf=1, RCA="Oui", DClas=dt.datetime(2026, 9, 2, 13, 30), DInit=dt.datetime(2026, 9, 2, 15, 45),
             DInter=dt.datetime(2026, 9, 5, 10, 0), DFin=dt.date(2026, 9, 30), PB=8000, Recup=0, ActReq="Oui",
             Ech=dt.date(2026, 10, 30))
    rows.append(r)
    # Incident TIC majeur DORA avec rapport intermédiaire tardif (février 2026)
    enr = dt.datetime(2026, 2, 10, 14, 20)
    r = base_row(enr, 1)
    r.update(Dom="Systèmes d'information et cybersécurité (TIC)", Bale="Fraude externe",
             Src="Contrôle de 2nd niveau (Risk / Conformité)", Presta="Prestataire informatique",
             Desc="Tentative d'hameçonnage ayant compromis une messagerie", DORA="Oui", Surv=dt.datetime(2026, 2, 9, 16, 0),
             Det=dt.datetime(2026, 2, 10, 13, 50), PEC=dt.datetime(2026, 2, 10, 14, 30), Res=dt.datetime(2026, 2, 11, 11, 0),
             Clo=dt.datetime(2026, 3, 6, 12, 0), EscN2=dt.datetime(2026, 2, 10, 14, 10), EscN3=dt.datetime(2026, 2, 10, 15, 0),
             Reaf=2, RCA="Oui", DClas=dt.datetime(2026, 2, 10, 18, 0), DInit=dt.datetime(2026, 2, 10, 21, 30),
             DInter=dt.datetime(2026, 2, 14, 6, 0), DFin=dt.date(2026, 3, 10), PB=25000, Recup=15000, ActReq="Oui",
             Ech=dt.date(2026, 3, 31), Real=dt.date(2026, 3, 27))
    rows.append(r)

    # Quasi-incidents
    for _ in range(30):
        d = rng.choice(days)
        enr = dt.datetime.combine(d, dt.time(rng.randint(9, 17), rng.choice(range(0, 60, 5))))
        crit = rng.choice([3, 4])
        r = base_row(enr, crit, "Quasi-incident")
        r["PEC"] = add_bh(enr, rng.uniform(1, 8))
        r["Res"] = add_bh(r["PEC"], rng.uniform(0.5, 5) * 9)
        r["Clo"] = add_bh(r["Res"], rng.uniform(1, 8) * 9)
        for k in ("Res", "Clo"):
            if r[k] > NOW:
                r[k] = None
        r["Reaf"] = 0
        if rng.random() < 0.3:
            r["ActReq"] = "Oui"
            r["Ech"] = enr.date() + dt.timedelta(days=30)
            r["Real"] = r["Ech"] - dt.timedelta(days=3) if r["Ech"] - dt.timedelta(days=3) <= NOW.date() else None
        rows.append(r)

    # Lignes annulées
    for d in (dt.date(2026, 1, 14), dt.date(2026, 7, 21)):
        r = base_row(dt.datetime.combine(d, dt.time(11, 0)), 4)
        r.update(PEC=dt.datetime.combine(d, dt.time(11, 30)), Desc="Doublon d'une déclaration existante")
        r["_statut"] = "Annulé"
        rows.append(r)

    # Ligne en anomalie de saisie (prise en compte antérieure à l'enregistrement)
    r = base_row(dt.datetime(2026, 10, 5, 15, 0), 3)
    r.update(PEC=dt.datetime(2026, 10, 5, 14, 0))
    rows.append(r)

    rows.sort(key=lambda x: x["Enr"])
    counters, ids = {}, []
    for r in rows:
        y = r["Enr"].year
        counters[y] = counters.get(y, 0) + 1
        r["ID"] = f"INC-{y}-{counters[y]:03d}"
        ids.append((r["ID"], r["Dom"]))
    for i, r in enumerate(rows):
        if r["Type"] == "Incident" and i > 5 and rng.random() < 0.07:
            prev = [pid for pid, dom in ids[:i] if dom == r["Dom"]] or [ids[i - 3][0]]
            r["Lien"] = rng.choice(prev)
        if "_statut" in r:
            r["Statut"] = r.pop("_statut")
        elif r.get("Clo"):
            r["Statut"] = "Clos"
        elif r.get("Res"):
            r["Statut"] = "Résolu"
        elif r.get("EscN2") or r.get("EscN3"):
            r["Statut"] = "En escalade"
        else:
            r["Statut"] = "En cours"
        if r.get("Reo") == "Oui":
            r["Com"] = "Rouvert après contrôle de second niveau"
    return rows
