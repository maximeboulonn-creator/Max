/* Risk Studio - socle commun aux trois propositions d'interface.
   Données de démonstration (valeurs fictives) et moteur de calcul : statuts,
   tendances, marges, scénarios de liquidité. Les trois interfaces lisent ce
   même bloc, la comparaison porte donc sur l'interface seule.
   Modifier ce fichier puis lancer build.py pour le recopier dans les pages. */
(function () {
"use strict";

const DATA = {
  meta: {
    entite: "Fundcraft France",
    aujourdhui: "2026-10-05",
    produit: "05/10/2026 08:12",
    arretes: ["2026-06-30", "2026-07-31", "2026-08-31", "2026-09-30"]
  },

  domaines: [
    { id: "LIQ", nom: "Liquidité" },
    { id: "VAL", nom: "Valorisation" },
    { id: "CON", nom: "Concentration" },
    { id: "LEV", nom: "Levier" },
    { id: "MKT", nom: "Marché et change" },
    { id: "OPS", nom: "Opérationnel" },
    { id: "REG", nom: "Réglementaire" }
  ],

  fonds: [
    { id: "AIR", nom: "AirFund Conviction Value Capital", court: "AirFund CVC",
      nature: "Fonds de fonds de private equity, evergreen ouvert",
      liquidite: "Rachats périodiques plafonnés par une gate ; préavis à confirmer",
      freq: "Mensuelle", devise: "EUR", encours: 142.6, dateVl: "2026-08-31" },
    { id: "ALE", nom: "Aletheon Growth III", court: "Aletheon III",
      nature: "Fonds fermé, co-investissements aux côtés de Partech, en direct ou via SPV",
      liquidite: "Aucun rachat ; appels de capitaux et ligne de crédit relais",
      freq: "Trimestrielle", devise: "EUR", encours: 118.3, dateVl: "2026-06-30" },
    { id: "OPN", nom: "FPS Openstone Infraworld", court: "Openstone",
      nature: "Fonds de fonds d'infrastructure semi-liquide, evergreen : 2 fonds cibles ouverts, 2 fermés",
      liquidite: "Rachats périodiques plafonnés par une gate ; préavis à confirmer",
      freq: "Mensuelle", devise: "EUR", encours: 96.4, dateVl: "2026-08-31" }
  ],

  /* sens "max" : plafond (alerte quand la valeur atteint al, dépassement au-delà de lim)
     sens "min" : plancher (alerte quand la valeur descend à al, dépassement sous lim)
     h : valeurs aux quatre arrêtés, null = donnée non reçue */
  indicateurs: [
    { id: "AIR-LIQ-01", f: "AIR", d: "LIQ", lib: "Ratio de couverture de liquidité", type: "KRI", nat: "Appétence interne",
      def: "(Poche liquide + facilité de crédit disponible) / (rachats notifiés + appels des fonds cibles à 3 mois)",
      u: "x", sens: "min", al: 1.5, lim: 1.0, h: [1.92, 1.71, 1.58, 1.35],
      src: "Suivi de trésorerie, registre de l'agent de transfert", ref: "Orientations ESMA34-39-897 (2019)",
      com: "Baisse continue sur le trimestre : hausse des rachats notifiés et appel attendu sur le premier fonds cible." },
    { id: "AIR-LIQ-02", f: "AIR", d: "LIQ", lib: "Poche liquide", type: "Limite", nat: "Prospectus",
      def: "(Trésorerie + fonds monétaires) / VL",
      u: "%", sens: "min", al: 7, lim: 5, h: [9.4, 8.7, 7.5, 6.6],
      src: "Comptabilité du fonds", ref: "Prospectus, poche de liquidité ; plancher interne à confirmer",
      com: "Pré-alerte franchie en septembre. Au rythme actuel, le plancher serait atteint en novembre sans nouvelle souscription.", act: "A5" },
    { id: "AIR-LIQ-03", f: "AIR", d: "LIQ", lib: "Proximité de la gate", type: "KRI", nat: "Prospectus",
      def: "Rachats nets de la fenêtre / seuil de gate",
      u: "%", sens: "max", al: 50, lim: 80, h: [22, 35, 38, 41],
      src: "Registre de l'agent de transfert", ref: "Prospectus, plafonnement des rachats", com: "" },
    { id: "AIR-LIQ-04", f: "AIR", d: "LIQ", lib: "Engagements non appelés / VL", type: "KRI", nat: "Appétence interne",
      def: "Engagements non appelés envers les fonds cibles / VL",
      u: "%", sens: "max", al: 50, lim: 70, h: [41, 40, 39, 38],
      src: "Suivi des engagements", ref: "Pratique de marché (surengagement)", com: "" },
    { id: "AIR-VAL-01", f: "AIR", d: "VAL", lib: "Part de la VL en valorisation estimée", type: "KRI", nat: "Appétence interne",
      def: "Part de la VL reposant sur des VL de fonds cibles non définitives",
      u: "%", sens: "max", al: 70, lim: 90, h: [58, 66, 61, 64],
      src: "Rapports des gérants des fonds cibles", ref: "AIFMD art. 19 ; Règlement délégué (UE) 231/2013, art. 67 à 74", com: "" },
    { id: "AIR-VAL-02", f: "AIR", d: "VAL", lib: "Ancienneté des VL définitives", type: "KRI", nat: "Appétence interne",
      def: "Ancienneté moyenne pondérée des dernières VL définitives des fonds cibles, en jours",
      u: "j", sens: "max", al: 120, lim: 180, h: [96, 121, 79, 123],
      src: "Rapports des gérants des fonds cibles", ref: "AIFMD art. 19",
      com: "VL définitives du deuxième trimestre non reçues pour le second fonds cible.", act: "A7" },
    { id: "AIR-VAL-03", f: "AIR", d: "VAL", lib: "Écart VL estimée / VL définitive", type: "KRI", nat: "Appétence interne",
      def: "Moyenne des écarts absolus sur les 4 dernières observations",
      u: "pb", sens: "max", al: 50, lim: 100, h: [38, 40, 44, 42],
      src: "Suivi des VL estimées et définitives", ref: "Règlement délégué (UE) 231/2013, art. 67 à 74", com: "" },
    { id: "AIR-CON-01", f: "AIR", d: "CON", lib: "Premier fonds cible", type: "Limite", nat: "Prospectus",
      def: "Exposition au premier fonds cible / VL",
      u: "%", sens: "max", al: 65, lim: 75, h: [58, 59, 61, 61],
      src: "Inventaire", ref: "Prospectus, à confirmer", com: "Concentration structurelle : architecture à deux fonds cibles." },
    { id: "AIR-CON-02", f: "AIR", d: "CON", lib: "Concentration du passif, 5 premiers porteurs", type: "KRI", nat: "Appétence interne",
      def: "Part des 5 premiers porteurs dans la VL",
      u: "%", sens: "max", al: 40, lim: 60, h: [37, 38, 38, 39],
      src: "Registre des porteurs", ref: "Orientations ESMA34-39-897 (2019)", com: "" },
    { id: "AIR-LEV-01", f: "AIR", d: "LEV", lib: "Levier, méthode de l'engagement", type: "Limite", nat: "Réglementaire",
      def: "Exposition selon la méthode de l'engagement / VL",
      u: "x", sens: "max", al: 1.3, lim: 1.5, h: [1.02, 1.03, 1.02, 1.02],
      src: "Calcul Annex IV", ref: "Règlement délégué (UE) 231/2013, art. 8", com: "Limite interne, à rapprocher du levier déclaré dans l'Annex IV." },
    { id: "AIR-MKT-01", f: "AIR", d: "MKT", lib: "Écart de couverture de change, part couverte USD", type: "Limite", nat: "Prospectus",
      def: "Valeur absolue de (ratio de couverture de la part couverte - 100 %), maximum du mois",
      u: "%", sens: "max", al: 3, lim: 5, h: [2.1, 2.8, 3.9, 5.6],
      src: "Suivi de couverture, confirmations de la contrepartie", ref: "Prospectus, politique de couverture de change",
      com: "Hausse de la VL des fonds cibles en USD sans réajustement du notionnel des contrats à terme.", act: "A1" },
    { id: "AIR-OPS-01", f: "AIR", d: "OPS", lib: "Écarts de rapprochement significatifs ouverts", type: "KRI", nat: "Appétence interne",
      def: "Nombre d'écarts de rapprochement au-delà du seuil de matérialité non résolus",
      u: "nb", sens: "max", al: 1, lim: 3, h: [0, 1, 0, 0],
      src: "Rapprochements dépositaire et valorisateur", ref: "Règlement délégué (UE) 231/2013, art. 13", com: "" },
    { id: "AIR-REG-01", f: "AIR", d: "REG", lib: "Ancienneté du dernier test de résistance de liquidité", type: "KRI", nat: "Réglementaire",
      def: "Jours écoulés depuis le dernier LST/MST validé",
      u: "j", sens: "max", al: 100, lim: 190, h: [0, 31, 62, 92],
      src: "Classeurs LST/MST", ref: "Orientations ESMA34-39-897 : au moins annuel, trimestriel recommandé", com: "" },

    { id: "ALE-LIQ-01", f: "ALE", d: "LIQ", lib: "Couverture des décaissements à 12 mois", type: "KRI", nat: "Appétence interne",
      def: "(Trésorerie + appels de capitaux programmés + ligne relais disponible) / décaissements prévus à 12 mois",
      u: "x", sens: "min", al: 1.2, lim: 1.0, h: [1.31, 1.27, 1.22, 1.17],
      src: "Plan de trésorerie, calendrier d'investissement", ref: "AIFMD art. 16 ; pratique de marché",
      com: "La marge repose sur la ligne relais : sans elle, la couverture tombe sous 1,00x." },
    { id: "ALE-LIQ-02", f: "ALE", d: "LIQ", lib: "Défaut des investisseurs sur appels de capitaux", type: "KRI", nat: "Règlement du fonds",
      def: "Montants appelés non réglés à 10 jours ouvrés / montants appelés",
      u: "%", sens: "max", al: 1, lim: 3, h: [0, 0, 0, 0],
      src: "Suivi des appels de capitaux", ref: "Règlement du fonds, clause de défaut", com: "" },
    { id: "ALE-VAL-01", f: "ALE", d: "VAL", lib: "Part de la VL valorisée sur un tour de plus de 12 mois", type: "KRI", nat: "Appétence interne",
      def: "Participations valorisées sur un dernier tour de financement de plus de 12 mois / VL",
      u: "%", sens: "max", al: 30, lim: 50, h: [31, 31, 31, 38],
      src: "Dossiers de valorisation trimestriels", ref: "Lignes directrices IPEV (2022) ; AIFMD art. 19",
      com: "Risque de valorisation obsolète, accentué par la rareté des tours de financement." },
    { id: "ALE-VAL-02", f: "ALE", d: "VAL", lib: "Ancienneté de la dernière VL", type: "KRI", nat: "Appétence interne",
      def: "Jours écoulés depuis la date de la dernière VL validée",
      u: "j", sens: "max", al: 120, lim: 180, h: [91, 31, 62, 92],
      src: "Valorisateur", ref: "Règlement du fonds", com: "" },
    { id: "ALE-CON-01", f: "ALE", d: "CON", lib: "Première participation", type: "Limite", nat: "Règlement du fonds",
      def: "Première participation, directe ou via SPV en transparence / engagements totaux",
      u: "%", sens: "max", al: 15, lim: 20, h: [12.4, 12.4, 12.4, 13.1],
      src: "Inventaire en transparence", ref: "Règlement du fonds, ratio d'emprise à confirmer", com: "" },
    { id: "ALE-CON-02", f: "ALE", d: "CON", lib: "Cinq premières participations", type: "KRI", nat: "Appétence interne",
      def: "Cinq premières participations en transparence / VL",
      u: "%", sens: "max", al: 55, lim: null, h: [52, 52, 52, 57],
      src: "Inventaire en transparence", ref: "Pratique de marché", com: "Pas de limite contractuelle : seuil d'alerte interne uniquement." },
    { id: "ALE-LEV-01", f: "ALE", d: "LEV", lib: "Utilisation de la ligne de crédit relais", type: "Limite", nat: "Règlement du fonds",
      def: "Encours tiré / montant de la ligne",
      u: "%", sens: "max", al: 75, lim: 100, h: [48, 61, 61, 72],
      src: "Relevés de la banque", ref: "Convention de crédit", com: "" },
    { id: "ALE-LEV-02", f: "ALE", d: "LEV", lib: "Ancienneté du plus ancien tirage", type: "Limite", nat: "Règlement du fonds",
      def: "Jours écoulés depuis le plus ancien tirage non remboursé",
      u: "j", sens: "max", al: 150, lim: 180, h: [79, 110, 141, 171],
      src: "Relevés de la banque", ref: "Règlement du fonds et convention de crédit, plafond à confirmer",
      com: "Plafond de 180 jours atteint le 09/10/2026. Avec un préavis d'appel de 10 jours ouvrés, un appel seul ne suffit plus à rembourser à temps.", act: "A4" },
    { id: "ALE-OPS-01", f: "ALE", d: "OPS", lib: "Délai de production de la VL trimestrielle", type: "KRI", nat: "Appétence interne",
      def: "Jours calendaires entre la date de VL et sa validation",
      u: "j", sens: "max", al: 60, lim: 90, h: [55, 48, 48, 48],
      src: "Valorisateur", ref: "Procédure de valorisation", com: "" },
    { id: "ALE-OPS-02", f: "ALE", d: "OPS", lib: "Rapports trimestriels des SPV reçus", type: "KRI", nat: "Appétence interne",
      def: "Rapports reçus / SPV en portefeuille",
      u: "%", sens: "min", al: 95, lim: 80, h: [100, 100, 88, null],
      src: "Administrateur des SPV", ref: "Procédure de valorisation",
      com: "Rapports du deuxième trimestre manquants pour 2 SPV sur 8 ; donnée de septembre non transmise.", act: "A3" },
    { id: "ALE-REG-01", f: "ALE", d: "REG", lib: "Dépôts réglementaires hors délai, 12 mois glissants", type: "KRI", nat: "Réglementaire",
      def: "Nombre de dépôts réglementaires effectués après l'échéance",
      u: "nb", sens: "max", al: 1, lim: 2, h: [0, 0, 0, 0],
      src: "Échéancier réglementaire", ref: "Règlement délégué (UE) 231/2013, art. 110", com: "" },

    { id: "OPN-LIQ-01", f: "OPN", d: "LIQ", lib: "Ratio de couverture de liquidité", type: "KRI", nat: "Appétence interne",
      def: "(Poche liquide + rachats obtenables des fonds cibles ouverts dans la fenêtre) / (rachats notifiés + appels des fonds cibles fermés à 3 mois)",
      u: "x", sens: "min", al: 1.5, lim: 1.0, h: [1.41, 1.22, 1.08, 0.93],
      src: "Suivi de trésorerie, registre de l'agent de transfert", ref: "Orientations ESMA34-39-897 (2019)",
      com: "Plancher franchi en septembre : rachats notifiés en hausse, appel attendu sur un fonds cible fermé, aucune facilité de crédit.", act: "A2" },
    { id: "OPN-LIQ-02", f: "OPN", d: "LIQ", lib: "Part de la VL investie en fonds cibles fermés", type: "KRI", nat: "Appétence interne",
      def: "Exposition aux fonds cibles fermés / VL",
      u: "%", sens: "max", al: 45, lim: 55, h: [43, 44, 46, 47],
      src: "Inventaire", ref: "Prospectus, allocation cible", com: "Les fonds cibles fermés ne contribuent pas à la liquidité du Fonds." },
    { id: "OPN-LIQ-03", f: "OPN", d: "LIQ", lib: "Proximité de la gate", type: "KRI", nat: "Prospectus",
      def: "Rachats nets de la fenêtre / seuil de gate",
      u: "%", sens: "max", al: 50, lim: 80, h: [31, 44, 52, 63],
      src: "Registre de l'agent de transfert", ref: "Prospectus, plafonnement des rachats", com: "" },
    { id: "OPN-LIQ-04", f: "OPN", d: "LIQ", lib: "Écart de préavis avec les fonds cibles ouverts", type: "KRI", nat: "Structure du fonds",
      def: "Préavis de rachat le plus long des fonds cibles ouverts - préavis du Fonds, en jours",
      u: "j", sens: "max", al: 0, lim: 30, h: [15, 15, 15, 15],
      src: "Documentation des fonds cibles", ref: "Orientations ESMA sur les outils de gestion de la liquidité (2025)",
      com: "Asymétrie structurelle : le Fonds ne peut pas obtenir la liquidité des fonds cibles ouverts dans sa propre fenêtre de rachat." },
    { id: "OPN-VAL-01", f: "OPN", d: "VAL", lib: "Part de la VL en valorisation estimée", type: "KRI", nat: "Appétence interne",
      def: "Part de la VL reposant sur des VL de fonds cibles non définitives",
      u: "%", sens: "max", al: 70, lim: 90, h: [72, 75, 71, 78],
      src: "Rapports des gérants des fonds cibles", ref: "AIFMD art. 19", com: "" },
    { id: "OPN-VAL-02", f: "OPN", d: "VAL", lib: "Écart VL estimée / VL définitive", type: "KRI", nat: "Appétence interne",
      def: "Moyenne des écarts absolus sur les 4 dernières observations",
      u: "pb", sens: "max", al: 50, lim: 100, h: [27, 29, 30, 31],
      src: "Suivi des VL estimées et définitives", ref: "Règlement délégué (UE) 231/2013, art. 67 à 74", com: "" },
    { id: "OPN-CON-01", f: "OPN", d: "CON", lib: "Premier fonds cible", type: "Limite", nat: "Prospectus",
      def: "Exposition au premier fonds cible / VL",
      u: "%", sens: "max", al: 35, lim: 40, h: [32, 33, 33, 34],
      src: "Inventaire", ref: "Prospectus, à confirmer", com: "" },
    { id: "OPN-CON-02", f: "OPN", d: "CON", lib: "Concentration du passif, 5 premiers porteurs", type: "KRI", nat: "Appétence interne",
      def: "Part des 5 premiers porteurs dans la VL",
      u: "%", sens: "max", al: 40, lim: 60, h: [36, 37, 37, 38],
      src: "Registre des porteurs", ref: "Orientations ESMA34-39-897 (2019)", com: "" },
    { id: "OPN-LEV-01", f: "OPN", d: "LEV", lib: "Levier, méthode de l'engagement", type: "Limite", nat: "Réglementaire",
      def: "Exposition selon la méthode de l'engagement / VL",
      u: "x", sens: "max", al: 1.3, lim: 1.5, h: [1.0, 1.0, 1.0, 1.0],
      src: "Calcul Annex IV", ref: "Règlement délégué (UE) 231/2013, art. 8", com: "" },
    { id: "OPN-MKT-01", f: "OPN", d: "MKT", lib: "Exposition hors EUR non couverte", type: "KRI", nat: "Appétence interne",
      def: "Exposition en devises hors EUR non couverte / VL",
      u: "%", sens: "max", al: 15, lim: 25, h: [11, 11, 12, 12],
      src: "Inventaire en transparence", ref: "Politique de couverture de change", com: "" },
    { id: "OPN-OPS-01", f: "OPN", d: "OPS", lib: "Incidents opérationnels ouverts", type: "KRI", nat: "Appétence interne",
      def: "Nombre d'incidents non clôturés au registre",
      u: "nb", sens: "max", al: 2, lim: 5, h: [1, 1, 2, 2],
      src: "Registre des incidents", ref: "Règlement délégué (UE) 231/2013, art. 13", com: "" },
    { id: "OPN-REG-01", f: "OPN", d: "REG", lib: "Ancienneté du dernier test de résistance de liquidité", type: "KRI", nat: "Réglementaire",
      def: "Jours écoulés depuis le dernier LST/MST validé",
      u: "j", sens: "max", al: 100, lim: 190, h: [85, 116, 147, 177],
      src: "Classeurs LST/MST", ref: "Orientations ESMA34-39-897 : au moins annuel, trimestriel recommandé",
      com: "Dernier LST/MST validé le 06/04/2026 ; exercice du troisième trimestre non lancé.", act: "A6" }
  ],

  actions: [
    { id: "A1", f: "AIR", ind: "AIR-MKT-01", prio: "Haute", ech: "2026-10-07", resp: "Gestion et Risques", statut: "Ouverte",
      objet: "Qualifier le dépassement (actif ou passif), réajuster le notionnel de couverture et inscrire au registre des dépassements" },
    { id: "A2", f: "OPN", ind: "OPN-LIQ-01", prio: "Haute", ech: "2026-10-06", resp: "Risques", statut: "Ouverte",
      objet: "Informer la Direction générale et présenter un plan de rétablissement : calendrier des appels, cession secondaire, facilité de crédit, activation éventuelle de la gate" },
    { id: "A4", f: "ALE", ind: "ALE-LEV-02", prio: "Haute", ech: "2026-10-05", resp: "Gestion", statut: "Ouverte",
      objet: "Vérifier qu'un appel de capitaux a été émis ; à défaut, obtenir de la banque une prorogation du tirage au-delà du 09/10/2026" },
    { id: "A3", f: "ALE", ind: "ALE-OPS-02", prio: "Moyenne", ech: "2026-10-09", resp: "Risques et middle office", statut: "Ouverte",
      objet: "Obtenir de l'administrateur des SPV les rapports manquants et la donnée de septembre" },
    { id: "A5", f: "AIR", ind: "AIR-LIQ-02", prio: "Moyenne", ech: "2026-10-14", resp: "Gestion", statut: "Ouverte",
      objet: "Valider avec la gestion un plan de reconstitution de la poche liquide avant la prochaine fenêtre de rachat" },
    { id: "A8", f: null, pipe: "PG", prio: "Moyenne", ech: "2026-10-14", resp: "Risques", statut: "En cours",
      objet: "Comité des risques : valider les paramètres de liquidité du Nourricier Partners Group (gate, préavis, poche liquide)" },
    { id: "A7", f: "AIR", ind: "AIR-VAL-02", prio: "Basse", ech: "2026-10-15", resp: "Risques", statut: "Ouverte",
      objet: "Relancer le gérant du second fonds cible pour les VL définitives du deuxième trimestre" },
    { id: "A6", f: "OPN", ind: "OPN-REG-01", prio: "Moyenne", ech: "2026-10-20", resp: "Risques", statut: "Ouverte",
      objet: "Lancer le LST/MST du troisième trimestre sur données au 30/09/2026" }
  ],

  echeances: [
    { date: "2026-10-09", f: ["ALE"], objet: "Remboursement du plus ancien tirage de la ligne relais", ref: "Convention de crédit", statut: "À risque" },
    { date: "2026-10-14", f: ["AIR", "ALE", "OPN"], objet: "Comité des risques du troisième trimestre", ref: "Procédure interne", statut: "En préparation" },
    { date: "2026-10-15", f: ["AIR", "OPN"], objet: "Mise à jour trimestrielle EMT et EET", ref: "Standard FinDatEx (pratique de marché)", statut: "À venir" },
    { date: "2026-10-20", f: ["OPN"], objet: "LST/MST du troisième trimestre", ref: "Orientations ESMA34-39-897 (2019)", statut: "Non lancé" },
    { date: "2026-10-30", f: ["ALE"], objet: "Annex IV AIFMD, troisième trimestre", ref: "Règlement délégué (UE) 231/2013, art. 110", statut: "À venir" },
    { date: "2026-11-14", f: ["AIR", "OPN"], objet: "Annex IV AIFMD, troisième trimestre (fonds de fonds : délai de 15 jours supplémentaires)", ref: "Règlement délégué (UE) 231/2013, art. 110", statut: "À venir" },
    { date: "2026-12-15", f: ["AIR"], objet: "Réexamen annuel du KID PRIIPs", ref: "Règlement délégué (UE) 2017/653, art. 15", statut: "À venir" },
    { date: "2026-12-31", f: ["AIR", "ALE", "OPN"], objet: "Réexamen annuel de la politique de gestion des risques", ref: "Règlement délégué (UE) 231/2013, art. 41", statut: "À venir" }
  ],

  pipeline: [
    { id: "PG", nom: "Partners Group Private Equity Opportunities ELTIF Feeder", nature: "Onboarding",
      etapes: ["Réception", "Revue documentaire", "Initial Risk Profile", "Comité des risques", "Paramétrage des limites", "Lancement"], cur: 3,
      prochaine: "Validation des paramètres de liquidité", ech: "2026-10-14",
      point: "Agrément ELTIF en instruction ; préavis du Nourricier à aligner sur celui du Maître.",
      url: "https://claude.ai/artifact/WJQcKuqhLSSDEsTBCBPZhi" },
    { id: "OTQ", nom: "Otentiq Private Equity X SLP", nature: "Onboarding",
      etapes: ["Réception", "Revue documentaire", "Initial Risk Profile", "Comité des risques", "Paramétrage des limites", "Lancement"], cur: 1,
      prochaine: "Initial Risk Profile", ech: "2026-10-23",
      point: "Structure de SLP : vérifier le régime des appels de capitaux et la clause de défaut." },
    { id: "ACF", nom: "ACF X Growth Buy-Out Europe", nature: "Initial Risk Profile",
      etapes: ["Collecte", "Rédaction", "Revue", "Validation"], cur: 1,
      prochaine: "Revue de l'IRP", ech: "2026-10-21",
      point: "Track record du gérant à obtenir par millésime." },
    { id: "ACP", nom: "Access Capital Partners", nature: "Due diligence",
      etapes: ["Collecte", "Analyse", "Benchmarking Preqin", "Note de due diligence", "Avis"], cur: 2,
      prochaine: "Note de due diligence", ech: "2026-10-30",
      point: "Échantillon de comparables Preqin à figer : millésime, stratégie, zone." }
  ],

  studios: [
    { nom: "KRI Studio", objet: "Référentiel et suivi des indicateurs clés de risque", etat: "En service", url: "https://claude.ai/artifact/L77zVHUjzfDRBUcSD3FJbF" },
    { nom: "Risk Limit Studio", objet: "Limites, contrôle avant opération, registre des dépassements", etat: "En service", url: "https://claude.ai/artifact/7mpwx6X9YX3SuMrDTu1VjZ" },
    { nom: "Annex IV Studio", objet: "Validation et génération XML de l'Annex IV AIFMD", etat: "En développement", url: "https://claude.ai/artifact/7bsZWA9dku9pQEQ7Ct6sot" },
    { nom: "KID PRIIPs Studio", objet: "SRI, MRM, VEV, scénarios de performance, EPT, EMT et EET", etat: "En développement", url: "" },
    { nom: "Cash Monitoring Studio", objet: "Trésorerie quotidienne, fonds monétaires, appels et distributions, gate", etat: "En développement", url: "" },
    { nom: "Annexe SFDR Studio", objet: "Annexes précontractuelles et périodiques SFDR", etat: "En service", url: "https://claude.ai/artifact/Ezw2BDs8sLdhANu8BPdS3z" },
    { nom: "Operating Memorandum", objet: "Operating Memorandum interactif de FPS Openstone Infraworld", etat: "En service", url: "https://claude.ai/artifact/B3hEgGQeLgP5T8otAj5C9s" },
    { nom: "Monitoring des closings", objet: "Suivi des closings et des souscriptions", etat: "En service", url: "https://claude.ai/artifact/B1XPvm9B9o9ugjxxiSgxd8" },
    { nom: "Plan de charge Risque", objet: "Charge de travail de la fonction de gestion des risques", etat: "En service", url: "https://claude.ai/artifact/FFCUwnB1FqyNaRCtfFsG51" }
  ],

  /* Scénarios de liquidité, calibrés au 30/09/2026.
     Fonds ouverts : le scénario de base reprend les rachats notifiés, d'où l'égalité
     entre la couverture de base et le ratio de couverture de liquidité. */
  liquidite: {
    AIR: {
      type: "ouvert",
      params: [
        { k: "vl", lib: "Valeur liquidative", v: 142.6, u: "EUR m", src: "VL au 31/08/2026" },
        { k: "poche", lib: "Poche liquide (trésorerie et fonds monétaires)", v: 6.6, u: "% VL", src: "Comptabilité du fonds" },
        { k: "obtenable", lib: "Rachats obtenables des fonds cibles dans la fenêtre", v: 0, u: "% VL", src: "Hypothèse : fonds cibles sans rachat" },
        { k: "credit", lib: "Facilité de crédit disponible", v: 5.0, u: "EUR m", src: "Convention de crédit, à confirmer" },
        { k: "gate", lib: "Gate par fenêtre de rachat", v: 5, u: "% VL", src: "Prospectus, à confirmer" }
      ],
      scen: [
        { k: "rachats", lib: "Rachats demandés", u: "% VL", v: { base: 3.4, adverse: 6, extreme: 12 }, src: "Base : rachats notifiés ; adverse et extrême : hypothèses" },
        { k: "appels", lib: "Appels des fonds cibles à 3 mois", u: "% VL", v: { base: 4.1, adverse: 5.0, extreme: 6.5 }, src: "Plan d'engagements ; stress : accélération des appels" }
      ]
    },
    OPN: {
      type: "ouvert",
      params: [
        { k: "vl", lib: "Valeur liquidative", v: 96.4, u: "EUR m", src: "VL au 31/08/2026" },
        { k: "poche", lib: "Poche liquide (trésorerie et fonds monétaires)", v: 4.2, u: "% VL", src: "Comptabilité du fonds" },
        { k: "obtenable", lib: "Rachats obtenables des fonds cibles ouverts dans la fenêtre", v: 1.3, u: "% VL", src: "Gates et préavis des fonds cibles" },
        { k: "credit", lib: "Facilité de crédit disponible", v: 0, u: "EUR m", src: "Aucune facilité en place" },
        { k: "gate", lib: "Gate par fenêtre de rachat", v: 5, u: "% VL", src: "Prospectus, à confirmer" }
      ],
      scen: [
        { k: "rachats", lib: "Rachats demandés", u: "% VL", v: { base: 3.9, adverse: 6, extreme: 10 }, src: "Base : rachats notifiés ; adverse et extrême : hypothèses" },
        { k: "appels", lib: "Appels des fonds cibles fermés à 3 mois", u: "% VL", v: { base: 2.0, adverse: 2.5, extreme: 3.0 }, src: "Plan d'engagements ; stress : accélération des appels" }
      ]
    },
    ALE: {
      type: "ferme",
      params: [
        { k: "treso", lib: "Trésorerie disponible", v: 6.2, u: "EUR m", src: "Relevés bancaires au 30/09/2026" },
        { k: "appels", lib: "Appels de capitaux programmés à 12 mois", v: 24.0, u: "EUR m", src: "Calendrier des appels" },
        { k: "ligne", lib: "Ligne relais disponible", v: 4.2, u: "EUR m", src: "Relevés de la banque" },
        { k: "decaiss", lib: "Décaissements prévus à 12 mois", v: 29.5, u: "EUR m", src: "Calendrier d'investissement et frais" }
      ],
      scen: [
        { k: "defaut", lib: "Défaut des investisseurs sur les appels", u: "%", v: { base: 0, adverse: 10, extreme: 25 }, src: "Hypothèses ; aucun défaut observé à ce jour" },
        { k: "maintien", lib: "Part de la ligne relais maintenue", u: "%", v: { base: 100, adverse: 100, extreme: 0 }, src: "Extrême : non-renouvellement par la banque" }
      ]
    }
  },

  hypothesesLiquidite: [
    "Le scénario de base reprend les rachats notifiés au 30/09/2026 ; les scénarios adverse et extrême sont calibrés à dire d'expert, faute d'historique de rachats suffisant pour une approche statistique.",
    "Les ressources comprennent la poche liquide, les rachats obtenables des fonds cibles ouverts dans la fenêtre et la facilité de crédit disponible. Aucune cession secondaire n'est retenue : décote et délai trop incertains.",
    "La gate est appliquée à son niveau prospectus, à confirmer ; les rachats différés reportent le besoin sur la fenêtre suivante sans l'éteindre.",
    "Les VL des fonds cibles sont reprises sans retraitement du lissage (desmoothing) : la couverture est surestimée si ces VL sont en retard sur le marché."
  ],

  references: [
    { texte: "Directive 2011/61/UE (AIFMD)", niveau: "Directive UE, 2011", objet: "Art. 15 gestion des risques, art. 16 liquidité, art. 19 évaluation, art. 24 reporting" },
    { texte: "Règlement délégué (UE) 231/2013", niveau: "Règlement UE, 2013", objet: "Art. 38 à 49 risques et liquidité, art. 67 à 74 évaluation, art. 110 Annex IV" },
    { texte: "Directive (UE) 2024/927 (AIFMD 2)", niveau: "Directive UE, 2024", objet: "Outils de gestion de la liquidité des FIA ouverts ; transposition au 16/04/2026" },
    { texte: "ESMA34-39-897", niveau: "Orientations ESMA, 2019", objet: "Tests de résistance de liquidité des OPCVM et des FIA" },
    { texte: "Orientations sur les outils de gestion de la liquidité", niveau: "Orientations ESMA, 2025", objet: "Sélection et calibrage des outils de gestion de la liquidité" },
    { texte: "Règlement délégué (UE) 2017/653", niveau: "Règlement UE, 2017", objet: "KID PRIIPs, art. 15 réexamen au moins annuel" },
    { texte: "Lignes directrices IPEV", niveau: "Pratique de marché, 2022", objet: "Évaluation des participations non cotées" }
  ]
};

/* ------------------------------------------------------------------ outils */

const STATUTS = {
  DEP:  { lib: "Dépassement",       court: "Dépassement",  rang: 4 },
  MAN:  { lib: "Donnée manquante",  court: "Manquant",     rang: 3 },
  SURV: { lib: "Sous surveillance", court: "Surveillance", rang: 2 },
  OK:   { lib: "Conforme",          court: "Conforme",     rang: 1 },
  NA:   { lib: "Non applicable",    court: "n.a.",         rang: 0 }
};
const ORDRE_STATUTS = ["DEP", "MAN", "SURV", "OK"];
const MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
              "septembre", "octobre", "novembre", "décembre"];
const EPS = 1e-9;

const fondsById = {}; DATA.fonds.forEach(f => { fondsById[f.id] = f; });
const domById = {}; DATA.domaines.forEach(d => { domById[d.id] = d; });
const indById = {}; DATA.indicateurs.forEach(i => { indById[i.id] = i; });
const actById = {}; DATA.actions.forEach(a => { actById[a.id] = a; });
const pipeById = {}; DATA.pipeline.forEach(p => { pipeById[p.id] = p; });

function esc(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
/* typographie française : espace insécable avant : ; ? ! % */
function fr(s) {
  return String(s == null ? "" : s).replace(/ ([:;?!%])/g, " $1");
}
function t(s) { return esc(fr(s)); }

function nombre(v, dec) {
  return new Intl.NumberFormat("fr-FR", { minimumFractionDigits: dec, maximumFractionDigits: dec }).format(v);
}
const DEC = { "%": 1, "x": 2, "pb": 0, "j": 0, "nb": 0, "EUR m": 1, "% VL": 1 };
function val(v, u) {
  if (v === null || v === undefined || isNaN(v)) return "n.d.";
  const d = DEC[u] !== undefined ? DEC[u] : 1;
  const n = nombre(v, d);
  if (u === "x") return n + "x";
  if (u === "nb") return n;
  if (u === "% VL") return n + " %";
  return n + " " + u;
}
function signe(v, u) {
  if (v === null || v === undefined || isNaN(v)) return "-";
  const s = val(Math.abs(v), u);
  if (Math.abs(v) < EPS) return s;
  return (v > 0 ? "+" : "-") + s;
}

function date(iso) { const p = iso.split("-"); return p[2] + "/" + p[1] + "/" + p[0]; }
function dateCourte(iso) { const p = iso.split("-"); return p[2] + "/" + p[1]; }
function dateLongue(iso) {
  const p = iso.split("-"); return Number(p[2]) + " " + MOIS[Number(p[1]) - 1] + " " + p[0];
}
function moisAnnee(iso) { const p = iso.split("-"); return MOIS[Number(p[1]) - 1] + " " + p[0]; }
function jours(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 86400000); }
function delai(iso) {
  const n = jours(DATA.meta.aujourdhui, iso);
  if (n < 0) return "en retard de " + (-n) + " j";
  if (n === 0) return "aujourd'hui";
  return "dans " + n + " j";
}

/* ------------------------------------------------------------------ moteur */

function valeur(ind, i) { const v = ind.h[i]; return v === undefined ? null : v; }

function statut(ind, i) {
  const v = valeur(ind, i);
  if (v === null) return "MAN";
  if (ind.sens === "max") {
    if (ind.lim !== null && v > ind.lim + EPS) return "DEP";
    if (ind.al !== null && v >= ind.al - EPS) return "SURV";
    return "OK";
  }
  if (ind.lim !== null && v < ind.lim - EPS) return "DEP";
  if (ind.al !== null && v <= ind.al + EPS) return "SURV";
  return "OK";
}

/* marge restante avant la limite, dans l'unité de l'indicateur (négative = dépassement) */
function marge(ind, i) {
  const v = valeur(ind, i);
  if (v === null || ind.lim === null) return null;
  return ind.sens === "max" ? ind.lim - v : v - ind.lim;
}

function tendance(ind, i) {
  if (i === 0) return { code: "nd", lib: "n.d." };
  const a = valeur(ind, i - 1), b = valeur(ind, i);
  if (a === null || b === null) return { code: "nd", lib: "n.d." };
  if (Math.abs(b - a) < EPS) return { code: "stable", lib: "Stable" };
  const hausse = b > a;
  const degrade = ind.sens === "max" ? hausse : !hausse;
  return degrade ? { code: "deg", lib: "Dégradation" } : { code: "ame", lib: "Amélioration" };
}

function pire(codes) {
  let w = "NA";
  codes.forEach(c => { if (STATUTS[c].rang > STATUTS[w].rang) w = c; });
  return w;
}

function indicateursDe(f, d) {
  return DATA.indicateurs.filter(x => (!f || x.f === f) && (!d || x.d === d));
}

function statutDomaine(f, d, i) {
  const l = indicateursDe(f, d);
  if (!l.length) return "NA";
  return pire(l.map(x => statut(x, i)));
}

function statutFonds(f, i) { return pire(indicateursDe(f).map(x => statut(x, i))); }

function compte(i, perimetre) {
  const c = { DEP: 0, MAN: 0, SURV: 0, OK: 0, total: 0 };
  DATA.indicateurs.forEach(x => {
    if (perimetre && perimetre.indexOf(x.f) < 0) return;
    c[statut(x, i)] += 1; c.total += 1;
  });
  return c;
}

/* indicateurs hors statut conforme, du plus grave au moins grave */
function pointsAttention(i, perimetre) {
  return DATA.indicateurs
    .filter(x => (!perimetre || perimetre.indexOf(x.f) >= 0) && statut(x, i) !== "OK")
    .sort((a, b) => STATUTS[statut(b, i)].rang - STATUTS[statut(a, i)].rang
      || DATA.fonds.indexOf(fondsById[a.f]) - DATA.fonds.indexOf(fondsById[b.f]));
}

function actionsOuvertes(perimetre) {
  return DATA.actions
    .filter(a => !perimetre || !a.f || perimetre.indexOf(a.f) >= 0)
    .slice().sort((a, b) => a.ech < b.ech ? -1 : a.ech > b.ech ? 1 : 0);
}

function echeancesDans(nbJours, perimetre) {
  return DATA.echeances.filter(e => {
    const n = jours(DATA.meta.aujourdhui, e.date);
    if (n < 0 || n > nbJours) return false;
    return !perimetre || e.f.some(x => perimetre.indexOf(x) >= 0);
  });
}

/* ------------------------------------------------------------------ liquidité */

const SCENARIOS = [
  { k: "base", lib: "Base" },
  { k: "adverse", lib: "Adverse" },
  { k: "extreme", lib: "Extrême" }
];

function copieLiquidite(id) { return JSON.parse(JSON.stringify(DATA.liquidite[id])); }

function paramsMap(cfg) { const m = {}; cfg.params.forEach(p => { m[p.k] = Number(p.v) || 0; }); return m; }
function scenMap(cfg, s) { const m = {}; cfg.scen.forEach(p => { m[p.k] = Number(p.v[s]) || 0; }); return m; }

function liqOuvert(cfg, s) {
  const p = paramsMap(cfg), q = scenMap(cfg, s);
  const ressources = p.vl * (p.poche + p.obtenable) / 100 + p.credit;
  const demande = p.vl * q.rachats / 100;
  const servis = p.vl * Math.min(q.rachats, p.gate) / 100;
  const appels = p.vl * q.appels / 100;
  const besoins = servis + appels;
  return { ressources, demande, servis, differes: demande - servis, appels, besoins,
           solde: ressources - besoins, couverture: besoins > 0 ? ressources / besoins : null };
}

function liqFerme(cfg, s) {
  const p = paramsMap(cfg), q = scenMap(cfg, s);
  const encaisses = p.appels * (1 - q.defaut / 100);
  const ligne = p.ligne * q.maintien / 100;
  const ressources = p.treso + encaisses + ligne;
  return { appels: p.appels, defautMontant: p.appels - encaisses, encaisses, treso: p.treso, ligne,
           ressources, besoins: p.decaiss, solde: ressources - p.decaiss,
           couverture: p.decaiss > 0 ? ressources / p.decaiss : null };
}

function liquidite(cfg, s) { return cfg.type === "ouvert" ? liqOuvert(cfg, s) : liqFerme(cfg, s); }

/* reverse stress test : le niveau de choc qui ramène la couverture à 1,00x */
function reverse(cfg) {
  const p = paramsMap(cfg);
  if (cfg.type === "ouvert") {
    const q = scenMap(cfg, "base");
    const ressources = p.vl * (p.poche + p.obtenable) / 100 + p.credit;
    const rmax = p.vl > 0 ? (ressources - p.vl * q.appels / 100) / p.vl * 100 : null;
    const gateProtege = rmax !== null && rmax >= p.gate;
    return {
      type: "ouvert", rmax, gate: p.gate, gateProtege,
      texte: rmax === null ? "Données insuffisantes." :
        rmax < 0 ? "Les ressources ne couvrent pas les seuls appels des fonds cibles du scénario de base, avant tout rachat." :
        "Les ressources couvrent au plus " + val(rmax, "% VL") + " de rachats par fenêtre, après les appels du scénario de base. " +
        (gateProtege ? "La gate (" + val(p.gate, "% VL") + ") s'applique avant l'épuisement des ressources." :
                       "Ce niveau est inférieur à la gate (" + val(p.gate, "% VL") + ") : le Fonds ne peut pas servir une fenêtre pleine sans cession ni financement.")
    };
  }
  const dmax = p.appels > 0 ? (1 - (p.decaiss - p.treso - p.ligne) / p.appels) * 100 : null;
  const dmaxSans = p.appels > 0 ? (1 - (p.decaiss - p.treso) / p.appels) * 100 : null;
  const borne = x => x === null ? "n.d." : x <= 0 ? "aucun défaut" : x >= 100 ? "plus de 100 %" : val(x, "%");
  return {
    type: "ferme", dmax, dmaxSans,
    texte: "Défaut des investisseurs supportable avant que la couverture ne tombe sous 1,00x : " + borne(dmax) +
      " avec la ligne relais, " + borne(dmaxSans) + " sans la ligne relais."
  };
}

/* ------------------------------------------------------------------ thème */

function theme(btn) {
  const libs = { auto: "automatique", light: "clair", dark: "sombre" };
  let cur = "auto";
  try { const s = localStorage.getItem("riskstudio_theme"); if (s === "light" || s === "dark") cur = s; } catch (e) {}
  function apply() {
    if (cur === "auto") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", cur);
    if (btn) btn.textContent = "Thème : " + libs[cur];
  }
  apply();
  if (btn) btn.addEventListener("click", () => {
    cur = cur === "auto" ? "light" : cur === "light" ? "dark" : "auto";
    try { localStorage.setItem("riskstudio_theme", cur); } catch (e) {}
    apply();
  });
}

window.RS = {
  DATA, STATUTS, ORDRE_STATUTS, SCENARIOS, fondsById, domById, indById, actById, pipeById,
  esc, fr, t, nombre, val, signe, date, dateCourte, dateLongue, moisAnnee, jours, delai,
  valeur, statut, marge, tendance, pire, indicateursDe, statutDomaine, statutFonds, compte,
  pointsAttention, actionsOuvertes, echeancesDans,
  copieLiquidite, liquidite, reverse, theme
};
})();
