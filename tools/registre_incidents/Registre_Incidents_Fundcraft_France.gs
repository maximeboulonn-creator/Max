/**
 * Registre des incidents et des escalades - Fundcraft France SAS
 * Google Apps Script lié au classeur Google Sheets issu de KPI_Registre_Incidents_Fundcraft_France.xlsx.
 *
 * Menu « Registre » :
 *  - Nouvel incident : prochaine référence INC-AAAA-NNN, date de déclaration du jour, statut Ouvert
 *    (verrou pour deux saisies simultanées).
 *  - Contrôler le registre : champs obligatoires, référentiels, références, chronologie, cohérence DORA,
 *    clôture, cause racine, actions correctives.
 *  - Échéances et alertes : notifications aux autorités (DORA 4 h / 24 h / 72 h / 1 mois, RGPD 72 h),
 *    escalades, réclamations, dépassements, actions correctives, ancienneté.
 *  - Figer l'arrêté trimestriel : copie en valeurs, protégée, dans un sous-dossier Drive.
 *  - Administration : diagnostic après import, deux éditeurs, droits et protections, journal des
 *    modifications, alertes par courriel, charte (Calibri, sans volets figés ni quadrillage).
 *
 * Installation : ouvrir le classeur Google Sheets > Extensions > Apps Script > coller ce fichier >
 * Enregistrer > recharger le classeur. Fuseau horaire Europe/Paris pour le classeur
 * (Fichier > Paramètres) et pour le projet (Paramètres du projet).
 * Les colonnes sont repérées par leur libellé en ligne 6 et les paramètres par les plages nommées :
 * l'ordre des colonnes peut évoluer sans modifier le script.
 */

const CFG = {
  registre: 'Registre', donnees: 'Données', tableau: 'Dashboard', journal: 'Journal',
  ligneEntete: 6, premiere: 7, derniere: 306,
  fuseau: 'Europe/Paris', prefixe: 'INC',
  police: 'Calibri', taille: 9,
  couleurs: {
    bandeau: '#2f4858', entete: '#218e8e', creme: '#fff8f2', cle: '#d9f785',
    texte: '#1a2027', gris: '#5a6b72', filet: '#e0e0e0',
  },
  preavisHeures: 24, preavisJours: 7, heureRecap: 8, maxJournal: 50,
  dossierArretes: 'Arrêtés du registre des incidents',
  nomArrete: 'Arrete_Registre_Incidents_Fundcraft_France',
  protection: 'Fundcraft - ', cleArrete: 'fundcraft_arrete',
  expediteur: 'Registre des incidents Fundcraft',
};

// Libellés des colonnes utilisées par le script (ligne d'en-tête du registre).
const CH = {
  ref: 'Réf.', fonds: 'Fonds ou véhicule', sc: 'Sous-catégorie', cat: 'Catégorie',
  surv: 'Date de survenance', det: 'Détection (date et heure)', decl: 'Déclaration au registre',
  pec: 'Prise en compte (date et heure)', orig: 'Origine de la détection', cause: 'Cause racine',
  typeDep: 'Dépassement actif ou passif', notif: 'Notification à une autorité',
  notifReal: 'Notification réalisée (date et heure)', clas: 'Classification DORA (date et heure)',
  inter: 'Rapport intermédiaire DORA (date et heure)', final: 'Rapport final DORA (date)',
  ar: 'Accusé de réception (réclamation)', nivAtt: "Niveau d'escalade atteint", dEsc: "Date d'escalade",
  statut: 'Statut', res: 'Résolution (date et heure)', clot: 'Date de clôture', reo: 'Réouvert',
  action: 'Action corrective', resp: 'Responsable', ech: "Échéance de l'action", real: 'Action réalisée le',
  ctrl: 'Contrôle de saisie', gravRang: 'Gravité retenue (rang)', nivReq: "Niveau d'escalade requis",
  reqRang: 'Niveau requis (rang)', attRang: 'Niveau atteint (rang)', escOK: 'Escalade conforme',
  delaiEsc: "Délai cible d'escalade (j.o.)", echNotif: 'Échéance de notification',
  doraEch: 'DORA - échéance de notification initiale',
};

// Colonnes de saisie (fond crème) : une ligne est vide si toutes ces cellules sont vides.
const SAISIES = ['Réf.', 'Fonds ou véhicule', 'Sous-catégorie', 'Description', 'Date de survenance',
  'Détection (date et heure)', 'Déclaration au registre', 'Prise en compte (date et heure)',
  'Origine de la détection', 'Cause racine', 'Impact financier brut (EUR)', 'Recouvrement (EUR)',
  'Écart de VL (bps)', 'Investisseurs lésés', 'Indemnisation investisseurs (EUR)', 'Impact réglementaire',
  'Impact réputationnel', 'Gravité forcée', 'Dépassement actif ou passif', 'Notification à une autorité',
  'Notification réalisée (date et heure)', 'Classification DORA (date et heure)',
  'Rapport intermédiaire DORA (date et heure)', 'Rapport final DORA (date)',
  'Accusé de réception (réclamation)', "Niveau d'escalade atteint", "Date d'escalade",
  'Nombre de réaffectations', 'Statut', 'Résolution (date et heure)', 'Date de clôture', 'Réouvert',
  'Action corrective', 'Responsable', "Échéance de l'action", 'Action réalisée le'];

const DATES = ['surv', 'det', 'decl', 'pec', 'notifReal', 'clas', 'inter', 'final', 'ar', 'dEsc', 'res',
  'clot', 'ech', 'real', 'echNotif', 'doraEch'];

// Plages nommées lues dans l'onglet Données.
const PARAMS = ['DateArrete', 'LibT', 'SeuilAge', 'DORA_H_INT', 'DORA_M_FIN', 'Recl_AR', 'Recl_Mois',
  'SC_Recl', 'Cat_Lim', 'Dep_Actif', 'Dep_Passif', 'TypeDep_Actif', 'Statut_Clos', 'Notif_DORA',
  'Notif_Aucune', 'Ref_Statut'];

const RANG_ETAT = { 'En retard': 0, "Aujourd'hui": 1, 'Sans échéance': 3, 'À venir': 4 };

// ------------------------------------------------------------------------------------------ menu
function onOpen() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  if (estArrete_(ss)) {
    ui.createMenu('Registre (arrêté figé)').addItem("Informations sur l'arrêté", 'infosArrete').addToUi();
    return;
  }
  ui.createMenu('Registre')
    .addItem('Nouvel incident', 'nouvelIncident')
    .addItem('Contrôler le registre', 'controlerRegistre')
    .addItem('Échéances et alertes', 'afficherEcheances')
    .addSeparator()
    .addItem("Figer l'arrêté trimestriel", 'figerArrete')
    .addSeparator()
    .addSubMenu(ui.createMenu('Administration')
      .addItem('Diagnostic après import', 'diagnostic')
      .addItem('Configurer les deux éditeurs', 'configurerEditeurs')
      .addItem('Appliquer les droits et protections', 'appliquerDroits')
      .addItem('Activer le journal et les alertes', 'activerAutomatisations')
      .addItem('Désactiver le journal et les alertes', 'desactiverAutomatisations')
      .addItem('Afficher ou masquer le journal', 'basculerJournal')
      .addItem('Appliquer la charte', 'appliquerCharte'))
    .addToUi();
}

function nouvelIncident() { avecInterface_(nouvelIncident_); }
function controlerRegistre() { avecInterface_(controlerRegistre_); }
function afficherEcheances() { avecInterface_(afficherEcheances_); }
function figerArrete() { avecInterface_(figerArrete_); }
function diagnostic() { avecInterface_(diagnostic_); }
function configurerEditeurs() { avecInterface_(configurerEditeurs_); }
function appliquerDroits() { avecInterface_(appliquerDroits_); }
function activerAutomatisations() { avecInterface_(activerAutomatisations_); }
function desactiverAutomatisations() { avecInterface_(desactiverAutomatisations_); }
function basculerJournal() { avecInterface_(basculerJournal_); }
function appliquerCharte() { avecInterface_(appliquerCharte_); }
function infosArrete() { avecInterface_(infosArrete_); }

function avecInterface_(fn) {
  try {
    fn();
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Registre des incidents', String((e && e.message) || e), ui.ButtonSet.OK);
  }
}

/** Sélectionne la ligne d'un incident (appelée depuis les fenêtres de contrôle et d'échéances). */
function allerALigne(ligne) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.registre);
  const ent = sh.getRange(CFG.ligneEntete, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(norm_);
  ss.setActiveSheet(sh);
  sh.getRange(Number(ligne), Math.max(1, ent.indexOf(norm_(CH.ref)) + 1)).activate();
}

// ------------------------------------------------------------------------------- nouvel incident
function nouvelIncident_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Une autre saisie est en cours : réessayer dans quelques secondes.');
  try {
    const ctx = contexte_();
    const lignes = lignes_(ctx.valeurs, ctx.idx);
    const libre = lignes.filter(o => o.vide)[0];
    if (!libre) {
      throw new Error('Le registre est plein (lignes ' + CFG.premiere + ' à ' + CFG.derniere + ') : prolonger les formules ' +
        'des colonnes calculées et les plages nommées R_* avant de saisir un nouvel incident.');
    }
    const maintenant = new Date();
    const annee = Number(Utilities.formatDate(maintenant, CFG.fuseau, 'yyyy'));
    const ref = prochaineRef_(lignes.map(o => o.ref), annee);
    const col = k => ctx.idx[norm_(CH[k])] + 1;
    ctx.sh.getRange(libre.ligne, col('ref')).setValue(ref);
    ctx.sh.getRange(libre.ligne, col('decl')).setValue(jour_(maintenant));
    ctx.sh.getRange(libre.ligne, col('statut')).setValue(ctx.p.Ref_Statut[0]);
    SpreadsheetApp.flush();  // visible par l'autre éditeur avant la libération du verrou
    try {
      ecrireJournal_(ctx.ss, [[maintenant, utilisateur_(null), CFG.registre, a1_(libre.ligne, col('ref')), ref,
        'Création', '', 'Réf., déclaration au registre et statut initialisés par le script']]);
    } catch (e) { /* journal protégé pour cet utilisateur : l'historique des versions Google fait foi */ }
    ctx.ss.setActiveSheet(ctx.sh);
    ctx.sh.getRange(libre.ligne, col('fonds')).activate();
    ctx.ss.toast(ref + ' créé ligne ' + libre.ligne + '. Saisir le fonds, la sous-catégorie, la description, ' +
      'puis la date et l\'heure de détection.', 'Nouvel incident', 8);
  } finally {
    lock.releaseLock();
  }
}

function prochaineRef_(refs, annee) {
  const re = new RegExp('^' + CFG.prefixe + '-' + annee + '-(\\d+)$');
  let max = 0;
  refs.forEach(r => {
    const m = re.exec(String(r === undefined || r === null ? '' : r).trim());
    if (m) max = Math.max(max, Number(m[1]));
  });
  return CFG.prefixe + '-' + annee + '-' + String(max + 1).padStart(3, '0');
}

// ------------------------------------------------------------------------------------- contrôles
function controlerRegistre_() {
  const ctx = contexte_();
  const lignes = lignes_(ctx.valeurs, ctx.idx);
  const anos = anomalies_(lignes, ctx.p, new Date());
  const n = lignes.filter(o => !o.vide).length;
  const html = page_('Contrôle du registre',
    n + ' incident(s) saisi(s) - ' + anos.length + ' anomalie(s) sur ' + compterLignes_(anos) + ' ligne(s). ' +
    'Cliquer sur une référence pour aller à la ligne.',
    ['Réf.', 'Ligne', 'Fonds ou véhicule', 'Anomalie'],
    anos.map(a => ({ ligne: a.ligne, cellules: [a.ref || '(sans réf.)', a.ligne, a.fonds, a.anomalie] })),
    'Aucune anomalie détectée.');
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(560),
    'Registre des incidents');
}

function anomalies_(lignes, p, maintenant) {
  const out = [];
  const motif = new RegExp('^' + CFG.prefixe + '-\\d{4}-\\d{3,}$');
  const refs = {};
  lignes.forEach(o => {
    if (!o.vide && !vide_(o.ref)) (refs[o.ref] = refs[o.ref] || []).push(o.ligne);
  });
  const auj = jour_(maintenant);
  lignes.forEach(o => {
    if (o.vide) return;
    const add = texte => out.push({ ligne: o.ligne, ref: o.ref || '', fonds: o.fonds || '', anomalie: texte });
    const det = estDate_(o.det);
    // 1. Règles de la colonne « Contrôle de saisie » (le détail du motif en plus).
    if (!det) {
      add('Détection (date et heure) manquante : la ligne est exclue des indicateurs');
    } else {
      if (vide_(o.fonds)) add('Fonds ou véhicule manquant');
      if (vide_(o.sc)) add('Sous-catégorie manquante');
      if (vide_(o.orig)) add('Origine de la détection manquante');
      if (vide_(o.statut)) add('Statut manquant');
      if (o.cat === 'À vérifier') add('Sous-catégorie hors référentiel (catégorie « À vérifier »)');
      if (o.statut === p.Statut_Clos && vide_(o.clot)) add('Statut Clôturé sans date de clôture');
      if (!vide_(o.clot) && o.statut !== p.Statut_Clos) add('Date de clôture saisie mais statut différent de Clôturé');
      if (!vide_(o.dEsc) && vide_(o.nivAtt)) add("Date d'escalade sans niveau d'escalade atteint");
    }
    // 2. Référence.
    if (vide_(o.ref)) {
      add('Réf. manquante');
    } else {
      if (!motif.test(o.ref)) add('Réf. au format inattendu (attendu ' + CFG.prefixe + '-AAAA-NNN)');
      if (refs[o.ref].length > 1) add('Réf. en double (lignes ' + refs[o.ref].join(', ') + ')');
    }
    // 3. Dates futures.
    [['det', 'Détection'], ['pec', 'Prise en compte'], ['res', 'Résolution'], ['notifReal', 'Notification réalisée'],
      ['clas', 'Classification DORA'], ['inter', 'Rapport intermédiaire DORA']].forEach(([k, lib]) => {
      if (estDate_(o[k]) && o[k] > maintenant) add('Date future : ' + lib);
    });
    [['surv', 'Date de survenance'], ['decl', 'Déclaration au registre'], ['dEsc', "Date d'escalade"],
      ['clot', 'Date de clôture'], ['ar', 'Accusé de réception'], ['final', 'Rapport final DORA'],
      ['real', 'Action réalisée le']].forEach(([k, lib]) => {
      if (estDate_(o[k]) && jour_(o[k]) > auj) add('Date future : ' + lib);
    });
    if (!det) return;
    // 4. Chronologie.
    const jDet = jour_(o.det);
    if (estDate_(o.surv) && jour_(o.surv) > jDet) add('Survenance postérieure à la détection');
    [['decl', 'Déclaration au registre antérieure'], ['dEsc', 'Escalade antérieure'],
      ['clot', 'Clôture antérieure'], ['ar', 'Accusé de réception antérieur']].forEach(([k, lib]) => {
      if (estDate_(o[k]) && jour_(o[k]) < jDet) add(lib + ' à la détection');
    });
    [['pec', 'Prise en compte antérieure'], ['res', 'Résolution antérieure'],
      ['notifReal', 'Notification antérieure'], ['clas', 'Classification DORA antérieure']].forEach(([k, lib]) => {
      if (estDate_(o[k]) && o[k] < o.det) add(lib + ' à la détection');
    });
    if (estDate_(o.clot) && estDate_(o.res) && jour_(o.clot) < jour_(o.res)) add('Clôture antérieure à la résolution');
    // 5. Notifications et DORA.
    const autorite = !vide_(o.notif) && o.notif !== p.Notif_Aucune;
    if (estDate_(o.notifReal) && !autorite) add('Notification réalisée sans autorité renseignée');
    if ((estDate_(o.clas) || estDate_(o.inter) || estDate_(o.final)) && o.notif !== p.Notif_DORA) {
      add('Dates DORA saisies alors que la notification n\'est pas « ' + p.Notif_DORA + ' »');
    }
    if (estDate_(o.inter) && !estDate_(o.notifReal)) add('Rapport intermédiaire DORA sans notification initiale');
    if (estDate_(o.final) && !estDate_(o.inter)) add('Rapport final DORA sans rapport intermédiaire');
    if (estDate_(o.inter) && estDate_(o.notifReal) && o.inter < o.notifReal) {
      add('Rapport intermédiaire DORA antérieur à la notification initiale');
    }
    if (estDate_(o.final) && estDate_(o.inter) && jour_(o.final) < jour_(o.inter)) {
      add('Rapport final DORA antérieur au rapport intermédiaire');
    }
    // 6. Clôture, escalade et actions correctives.
    if (o.statut === p.Statut_Clos && !estDate_(o.res)) add('Clôturé sans date de résolution (indicateurs de résolution incomplets)');
    if (o.reo === 'Oui' && !estDate_(o.res)) add('Réouvert sans date de résolution');
    if (Number(o.gravRang) >= 3 && o.statut === p.Statut_Clos && vide_(o.cause)) {
      add('Cause racine manquante (incident significatif ou majeur clôturé)');
    }
    if (o.escOK === 0) add("Escalade non conforme (niveau ou délai) à la date d'arrêté");
    if (!vide_(o.action) && vide_(o.resp)) add('Action corrective sans responsable');
    if (!vide_(o.action) && !estDate_(o.ech)) add("Action corrective sans échéance");
    if (estDate_(o.real) && vide_(o.action)) add('Date de réalisation sans action corrective');
  });
  return out;
}

// ------------------------------------------------------------------------------------- échéances
function afficherEcheances_() {
  const ctx = contexte_();
  const items = echeances_(lignes_(ctx.valeurs, ctx.idx), ctx.p, new Date());
  const urgents = items.filter(x => x.etat !== 'À venir').length;
  const html = page_('Échéances et alertes',
    'Au ' + Utilities.formatDate(new Date(), CFG.fuseau, 'dd/MM/yyyy HH:mm') + ' - ' + urgents +
    ' échéance(s) en retard, imminente(s) ou sans date, ' + (items.length - urgents) + ' à venir. ' +
    'Jours fériés non exclus, comme dans le classeur.',
    ['Réf.', 'Fonds ou véhicule', 'Échéance', 'Date limite', 'État', 'Précision'],
    items.map(x => ({ ligne: x.ligne, etat: x.etat,
      cellules: [x.ref || 'ligne ' + x.ligne, x.fonds, x.nature, x.echeance ? fmt_(x.echeance, x.horaire) : '', x.etat, x.detail] })),
    'Aucune échéance ouverte.');
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(html).setWidth(980).setHeight(560),
    'Registre des incidents');
}

function echeances_(lignes, p, maintenant) {
  const out = [];
  lignes.forEach(o => {
    if (o.vide || !estDate_(o.det)) return;
    const push = (nature, echeance, horaire, detail, siProche) => {
      const etat = etat_(echeance, horaire, maintenant);
      if (siProche && etat === 'À venir') return;
      out.push({ ligne: o.ligne, ref: o.ref || '', fonds: o.fonds || '', nature: nature, echeance: echeance,
        horaire: horaire, etat: etat, detail: detail || '',
        cle: [o.ref || o.ligne, nature, echeance ? echeance.getTime() : ''].join('|') });
    };
    // Notifications aux autorités : l'échéance court même après la clôture.
    if (!vide_(o.notif) && o.notif !== p.Notif_Aucune) {
      if (o.notif === p.Notif_DORA) {
        if (!estDate_(o.notifReal)) {
          push('DORA - notification initiale', estDate_(o.doraEch) ? o.doraEch : null, true,
            estDate_(o.clas) ? '4 h après la classification, au plus tard 24 h après la détection'
              : 'classification non saisie : 24 h après la détection');
        } else if (!estDate_(o.inter)) {
          push('DORA - rapport intermédiaire', ajouterHeures_(o.notifReal, Number(p.DORA_H_INT)), true,
            p.DORA_H_INT + ' h après la notification initiale');
        } else if (!estDate_(o.final)) {
          push('DORA - rapport final', ajouterMois_(o.inter, Number(p.DORA_M_FIN)), false,
            p.DORA_M_FIN + ' mois après le rapport intermédiaire');
        }
      } else if (!estDate_(o.notifReal)) {
        push('Notification ' + o.notif, estDate_(o.echNotif) ? o.echNotif : null, true,
          estDate_(o.echNotif) ? 'depuis la détection' : 'délai non paramétré dans Données : à suivre');
      }
    }
    // Actions correctives : suivies même après la clôture de l'incident.
    if (estDate_(o.ech) && !estDate_(o.real)) {
      push('Action corrective', o.ech, false, vide_(o.resp) ? 'responsable non renseigné' : 'responsable : ' + o.resp);
    }
    if (estDate_(o.clot)) return;
    const req = Number(o.reqRang) || 0;
    const att = Number(o.attRang) || 0;
    if (req > 0 && (!estDate_(o.dEsc) || att < req)) {
      push('Escalade : ' + o.nivReq, ajouterJoursOuvres_(o.det, Number(o.delaiEsc) || 0), false,
        estDate_(o.dEsc) ? 'niveau atteint inférieur au niveau requis' : 'escalade non saisie');
    }
    if (o.sc === p.SC_Recl) {
      if (!estDate_(o.ar)) push('Réclamation : accusé de réception', ajouterJoursOuvres_(o.det, Number(p.Recl_AR)), false);
      push('Réclamation : réponse et clôture', ajouterMois_(o.det, Number(p.Recl_Mois)), false);
    }
    if (o.cat === p.Cat_Lim) {
      const actif = o.typeDep === p.TypeDep_Actif;
      push('Dépassement : régularisation', ajouterJours_(o.det, Number(actif ? p.Dep_Actif : p.Dep_Passif)), false,
        vide_(o.typeDep) ? 'type non renseigné : délai passif retenu' : (actif ? 'dépassement actif' : 'dépassement passif'));
    }
    push('Ancienneté : ouvert depuis plus de ' + p.SeuilAge + ' jours', ajouterJours_(o.det, Number(p.SeuilAge)), false,
      'seuil de l\'onglet Données', true);
  });
  return out.sort((a, b) => (rangEtat_(a.etat) - rangEtat_(b.etat)) ||
    ((a.echeance ? a.echeance.getTime() : Infinity) - (b.echeance ? b.echeance.getTime() : Infinity)) || (a.ligne - b.ligne));
}

function etat_(echeance, horaire, maintenant) {
  if (!echeance) return 'Sans échéance';
  if (horaire) {
    const h = (echeance.getTime() - maintenant.getTime()) / 36e5;
    return h < 0 ? 'En retard' : h <= CFG.preavisHeures ? 'Sous ' + CFG.preavisHeures + ' h' : 'À venir';
  }
  const j = Math.round((jour_(echeance).getTime() - jour_(maintenant).getTime()) / 864e5);
  return j < 0 ? 'En retard' : j === 0 ? "Aujourd'hui" : j <= CFG.preavisJours ? 'Sous ' + CFG.preavisJours + ' jours' : 'À venir';
}

function rangEtat_(etat) {
  return etat in RANG_ETAT ? RANG_ETAT[etat] : 2;  // « Sous N h » et « Sous N jours »
}

// --------------------------------------------------------------------------- alertes par courriel
/** Déclencheur quotidien : récapitulatif des échéances en retard, imminentes ou sans date. */
function envoyerRecapitulatif() {
  const ctx = contexte_();
  const items = echeances_(lignes_(ctx.valeurs, ctx.idx), ctx.p, new Date()).filter(x => x.etat !== 'À venir');
  if (!items.length) return;
  const retard = items.filter(x => x.etat === 'En retard').length;
  envoyer_(ctx.ss, 'Registre des incidents - ' + retard + ' échéance(s) en retard, ' + (items.length - retard) +
    ' à suivre', 'Récapitulatif quotidien des échéances', items);
}

/** Déclencheur horaire : notifications aux autorités en retard ou à moins de 24 h, envoyées une seule fois par état. */
function surveillerNotifications() {
  const ctx = contexte_();
  const urgents = echeances_(lignes_(ctx.valeurs, ctx.idx), ctx.p, new Date())
    .filter(x => x.horaire && x.echeance && x.etat !== 'À venir');
  const props = proprietes_();
  let deja = [];
  try { deja = JSON.parse(props.getProperty('ALERTES') || '[]'); } catch (e) { deja = []; }
  const cles = urgents.map(x => x.cle + '|' + x.etat);
  const nouveaux = urgents.filter((x, i) => deja.indexOf(cles[i]) < 0);
  props.setProperty('ALERTES', JSON.stringify(cles));
  if (nouveaux.length) {
    envoyer_(ctx.ss, 'Registre des incidents - notification à une autorité : ' + nouveaux.map(x => x.ref).join(', '),
      'Notifications aux autorités en retard ou à moins de ' + CFG.preavisHeures + ' h', nouveaux);
  }
}

function envoyer_(ss, sujet, titre, items) {
  const dest = destinataires_();
  if (!dest.length) return;
  const c = CFG.couleurs;
  const police = 'font-family:Calibri,Carlito,Arial,sans-serif;font-size:9pt;';
  const th = police + 'background:' + c.entete + ';color:#ffffff;text-align:left;padding:4px 6px;font-weight:bold';
  const td = police + 'color:' + c.texte + ';padding:4px 6px;border-bottom:1px solid ' + c.filet + ';vertical-align:top';
  const lignes = items.map(x => '<tr>' + [x.ref || 'ligne ' + x.ligne, x.fonds, x.nature,
    x.echeance ? fmt_(x.echeance, x.horaire) : '', x.etat, x.detail].map((v, i) =>
    '<td style="' + td + ([0, 3, 4].indexOf(i) >= 0 ? ';white-space:nowrap' : '') +
    (i === 4 && x.etat === 'En retard' ? ';background:' + c.cle + ';font-weight:bold' : '') + '">' +
    esc_(v) + '</td>').join('') + '</tr>').join('');
  const html = '<div style="' + police + 'color:' + c.texte + '">' +
    '<div style="font-size:10pt;font-weight:bold;color:' + c.bandeau + ';margin-bottom:6px">' + esc_(titre) + '</div>' +
    '<table style="border-collapse:collapse"><tr>' + ['Réf.', 'Fonds ou véhicule', 'Échéance', 'Date limite', 'État',
    'Précision'].map(h => '<th style="' + th + '">' + h + '</th>').join('') + '</tr>' + lignes + '</table>' +
    '<p style="' + police + 'color:' + c.gris + '">Registre : <a style="color:' + c.entete + '" href="' + ss.getUrl() + '">' +
    esc_(ss.getName()) + '</a>. Message automatique : ni description ni montant, consulter le registre.</p></div>';
  const texte = titre + '\n\n' + items.map(x => [x.ref, x.fonds, x.nature,
    x.echeance ? fmt_(x.echeance, x.horaire) : 'sans date', x.etat].join(' - ')).join('\n') + '\n\n' + ss.getUrl();
  MailApp.sendEmail({ to: dest.join(','), subject: sujet, htmlBody: html, body: texte, name: CFG.expediteur });
}

function destinataires_() {
  const e = editeurs_();
  if (e.length) return e;
  const moi = Session.getEffectiveUser().getEmail();
  return moi ? [moi] : [];
}

// ------------------------------------------------------------------------- journal des modifications
/** Déclencheur installable « à la modification » : trace chaque saisie du registre et des données. */
function journaliser(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const nom = sh.getName();
  if (nom !== CFG.registre && nom !== CFG.donnees) return;
  const ss = e.source || sh.getParent();
  const r0 = e.range.getRow(), c0 = e.range.getColumn(), nr = e.range.getNumRows(), nc = e.range.getNumColumns();
  const horo = new Date();
  const util = utilisateur_(e);
  const registre = nom === CFG.registre;
  const entetes = registre ? sh.getRange(CFG.ligneEntete, 1, 1, sh.getLastColumn()).getDisplayValues()[0] : [];
  const cRef = registre ? entetes.map(norm_).indexOf(norm_(CH.ref)) : -1;
  const refs = cRef >= 0 ? sh.getRange(r0, cRef + 1, nr, 1).getDisplayValues().map(x => x[0]) : [];
  const libelles = registre ? [] : sh.getRange(r0, 2, nr, 1).getDisplayValues().map(x => x[0]);
  const affiche = e.range.getDisplayValues();
  const unique = nr * nc === 1;
  const lignes = [];
  for (let i = 0; i < nr && lignes.length < CFG.maxJournal; i++) {
    for (let j = 0; j < nc && lignes.length < CFG.maxJournal; j++) {
      lignes.push([horo, util, nom, a1_(r0 + i, c0 + j), refs[i] || '',
        registre ? (entetes[c0 + j - 1] || '') : (libelles[i] || ''),
        unique ? ancienneValeur_(e) : '(modification de plusieurs cellules)', affiche[i][j]]);
    }
  }
  if (nr * nc > CFG.maxJournal) {
    lignes.push([horo, util, nom, e.range.getA1Notation(), '', '', '',
      (nr * nc - CFG.maxJournal) + ' cellule(s) supplémentaire(s) non détaillée(s)']);
  }
  ecrireJournal_(ss, lignes);
}

function ecrireJournal_(ss, lignes) {
  if (!lignes.length) return;
  const j = feuilleJournal_(ss);
  j.getRange(j.getLastRow() + 1, 1, lignes.length, 8).setValues(lignes);
}

function feuilleJournal_(ss) {
  let j = ss.getSheetByName(CFG.journal);
  if (j) return j;
  const c = CFG.couleurs;
  j = ss.insertSheet(CFG.journal, ss.getSheets().length);
  j.getRange(1, 1, j.getMaxRows(), 8).setFontFamily(CFG.police).setFontSize(CFG.taille).setFontColor(c.texte);
  j.getRange(1, 1, 1, 8).setValues([['Horodatage', 'Utilisateur', 'Onglet', 'Cellule', 'Réf.', 'Champ',
    'Ancienne valeur', 'Nouvelle valeur']]).setBackground(c.entete).setFontColor('#ffffff').setFontWeight('bold');
  j.getRange(2, 1, j.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  [130, 190, 80, 70, 100, 200, 180, 180].forEach((w, i) => j.setColumnWidth(i + 1, w));
  j.setFrozenRows(0);
  j.setHiddenGridlines(true);
  proteger_(j, 'journal des modifications en lecture seule', []);
  j.hideSheet();
  return j;
}

function ancienneValeur_(e) {
  if (e.oldValue === undefined || e.oldValue === null) return '';
  const brut = String(e.oldValue);
  const fmt = String(e.range.getNumberFormat() || '');
  const n = Number(brut);
  if (brut.trim() !== '' && !isNaN(n) && /[dy]/i.test(fmt) && fmt.indexOf('#') < 0) return dateSerie_(n);
  return brut;
}

function dateSerie_(n) {
  const d = new Date(Math.round((n - 25569) * 864e5));
  return Utilities.formatDate(d, 'UTC', n % 1 ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy');
}

function utilisateur_(e) {
  let m = '';
  try {
    if (e && e.user) m = typeof e.user.getEmail === 'function' ? e.user.getEmail() : String(e.user);
    else if (!e) m = Session.getActiveUser().getEmail();
  } catch (x) { m = ''; }
  return m || 'non communiqué par Google';
}

function basculerJournal_() {
  const ss = SpreadsheetApp.getActive();
  const j = feuilleJournal_(ss);
  if (j.isSheetHidden()) {
    j.showSheet();
    ss.setActiveSheet(j);
  } else {
    ss.setActiveSheet(ss.getSheetByName(CFG.registre));
    j.hideSheet();
  }
}

// ------------------------------------------------------------------------------ arrêté figé
function figerArrete_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const ctx = contexte_();
  const lib = String(ss.getRangeByName('LibT').getDisplayValue()).trim();
  const dateArr = fmt_(ctx.p.DateArrete, false);
  const anos = anomalies_(lignes_(ctx.valeurs, ctx.idx), ctx.p, new Date());
  const msg = 'Créer une copie figée en valeurs du classeur, arrêté au ' + dateArr + ' (' + lib + ') ?' +
    (anos.length ? '\n\nAttention : ' + anos.length + ' anomalie(s) de saisie détectée(s) (menu Contrôler le registre).' : '');
  if (ui.alert("Figer l'arrêté", msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  SpreadsheetApp.flush();
  const fichier = DriveApp.getFileById(ss.getId());
  const parents = fichier.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const sous = parent.getFoldersByName(CFG.dossierArretes);
  const dossier = sous.hasNext() ? sous.next() : parent.createFolder(CFG.dossierArretes);
  let nom = CFG.nomArrete + '_' + lib.replace(/\s+/g, '_');
  if (dossier.getFilesByName(nom).hasNext()) {
    if (ui.alert("Figer l'arrêté", 'Un arrêté ' + lib + ' existe déjà. En créer un nouveau, horodaté ?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    nom += '_' + Utilities.formatDate(new Date(), CFG.fuseau, 'yyyy-MM-dd_HHmm');
  }
  const copie = fichier.makeCopy(nom, dossier);
  const cs = SpreadsheetApp.openById(copie.getId());
  cs.getSheets().forEach(s => {
    const r = s.getDataRange();
    r.copyTo(r, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  });
  const info = { arrete: dateArr, periode: lib, fige: Utilities.formatDate(new Date(), CFG.fuseau, 'dd/MM/yyyy HH:mm'),
    par: Session.getEffectiveUser().getEmail(), source: ss.getUrl(), anomalies: anos.length };
  cs.addDeveloperMetadata(CFG.cleArrete, JSON.stringify(info));
  const tab = cs.getSheetByName(CFG.tableau);
  if (tab) tab.getRange('B2').setNote('Copie figée en valeurs le ' + info.fige + ' par ' + info.par + '. Source : ' + info.source);
  cs.getSheets().forEach(s => proteger_(s, 'arrêté figé en lecture seule', []));
  const html = page_("Arrêté figé", 'Copie créée dans le dossier « ' + CFG.dossierArretes + ' » : formules remplacées ' +
    'par leurs valeurs, onglets protégés.', ['Élément', 'Valeur'],
    [['Fichier', nom], ['Arrêté', dateArr + ' (' + lib + ')'], ['Figé le', info.fige + ' par ' + info.par],
      ['Anomalies de saisie au moment du figement', String(anos.length)]].map(x => ({ cellules: x })), '',
    { lien: copie.getUrl(), libelleLien: "Ouvrir l'arrêté figé" });
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(300), 'Registre des incidents');
}

function estArrete_(ss) {
  try {
    return ss.getDeveloperMetadata().some(m => m.getKey() === CFG.cleArrete);
  } catch (e) {
    return false;
  }
}

function infosArrete_() {
  const ss = SpreadsheetApp.getActive();
  const m = ss.getDeveloperMetadata().filter(x => x.getKey() === CFG.cleArrete)[0];
  const i = JSON.parse(m.getValue());
  const ui = SpreadsheetApp.getUi();
  ui.alert("Arrêté figé", 'Arrêté au ' + i.arrete + ' (' + i.periode + ')\nFigé le ' + i.fige + ' par ' + i.par +
    '\nAnomalies de saisie au figement : ' + i.anomalies + '\nSource : ' + i.source, ui.ButtonSet.OK);
}

// ---------------------------------------------------------------------------------- administration
function diagnostic_() {
  const ss = SpreadsheetApp.getActive();
  const res = [];
  const ok = (point, conforme, detail) => res.push({ etat: conforme ? 'Conforme' : 'En retard',
    cellules: [point, conforme ? 'Conforme' : 'À corriger', detail || ''] });
  [CFG.tableau, CFG.registre, CFG.donnees].forEach(n => ok('Onglet ' + n, !!ss.getSheetByName(n)));
  const sh = ss.getSheetByName(CFG.registre);
  let colonnesOK = false;
  if (sh) {
    const ix = indexColonnes_(sh.getRange(CFG.ligneEntete, 1, 1, sh.getLastColumn()).getDisplayValues()[0]);
    colonnesOK = !ix.manquants.length;
    ok('Libellés des colonnes (ligne ' + CFG.ligneEntete + ')', colonnesOK,
      colonnesOK ? 'tous reconnus' : 'Introuvables : ' + ix.manquants.join(', '));
  }
  const absents = PARAMS.filter(n => !ss.getRangeByName(n));
  ok('Plages nommées', !absents.length, absents.length ? 'Absentes : ' + absents.join(', ') : PARAMS.length + ' vérifiées');
  const tzc = ss.getSpreadsheetTimeZone(), tzs = Session.getScriptTimeZone();
  ok('Fuseau horaire du classeur', tzc === CFG.fuseau, tzc + (tzc === CFG.fuseau ? '' : ' : Fichier > Paramètres > ' + CFG.fuseau));
  ok('Fuseau horaire du script', tzs === CFG.fuseau, tzs + (tzs === CFG.fuseau ? '' : ' : Apps Script > Paramètres du projet'));
  const loc = ss.getSpreadsheetLocale();
  ok('Paramètres régionaux', String(loc).indexOf('fr') === 0, loc + ' (format des dates jj/mm/aaaa)');
  ss.getSheets().forEach(s => {
    if (s.getName() === CFG.journal) return;
    const v = s.getDataRange().getDisplayValues();
    const err = [];
    v.forEach((r, i) => r.forEach((x, j) => {
      if (/^#(N\/A|REF!|NAME\?|VALUE!|DIV\/0!|NUM!|NULL!|ERROR!)/.test(x)) err.push(a1_(i + 1, j + 1));
    }));
    ok('Erreurs de formule - ' + s.getName(), !err.length,
      err.length ? err.length + ' cellule(s) : ' + err.slice(0, 12).join(', ') + (err.length > 12 ? '...' : '') : 'aucune');
    const charte = !s.getFrozenRows() && !s.getFrozenColumns() && s.hasHiddenGridlines();
    ok('Volets figés et quadrillage - ' + s.getName(), charte, charte ? 'aucun' : 'Administration > Appliquer la charte');
  });
  if (colonnesOK && !absents.length) {
    const ctx = contexte_();
    const lignes = lignes_(ctx.valeurs, ctx.idx);
    const aCompleter = lignes.filter(o => o.ctrl === 'À compléter').length;
    const regles = compterLignes_(anomalies_(lignes, ctx.p, new Date()).filter(a => REGLES_CONTROLE_.test(a.anomalie)));
    ok('Colonne « Contrôle de saisie » et contrôles du script', aCompleter === regles,
      aCompleter + ' ligne(s) « À compléter » dans le classeur, ' + regles + ' selon le script');
  }
  const ed = editeurs_();
  ok('Deux éditeurs configurés', ed.length === 2, ed.join(', ') || 'Administration > Configurer les deux éditeurs');
  const decl = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  ok('Journal et alertes actifs', ['journaliser', 'envoyerRecapitulatif', 'surveillerNotifications'].every(f => decl.indexOf(f) >= 0),
    decl.length ? decl.join(', ') : 'Administration > Activer le journal et les alertes');
  const html = page_('Diagnostic après import', 'Contrôles de conversion Excel vers Google Sheets et de configuration.',
    ['Point de contrôle', 'Résultat', 'Détail'], res, '');
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(860).setHeight(560), 'Registre des incidents');
}

// Motifs des anomalies équivalentes à la colonne « Contrôle de saisie » du classeur.
const REGLES_CONTROLE_ = /^(Détection \(date et heure\) manquante|Fonds ou véhicule manquant|Sous-catégorie manquante|Origine de la détection manquante|Statut manquant|Sous-catégorie hors référentiel|Statut Clôturé sans date de clôture|Date de clôture saisie mais|Date d'escalade sans niveau)/;

function configurerEditeurs_() {
  const ui = SpreadsheetApp.getUi();
  const actuels = editeurs_();
  const r = ui.prompt('Deux éditeurs du registre', 'Adresses des deux personnes autorisées à modifier le registre, ' +
    'séparées par une virgule.' + (actuels.length ? '\nActuellement : ' + actuels.join(', ') : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const a = analyserAdresses_(r.getResponseText());
  if (a.erreur) throw new Error(a.erreur);
  proprietes_().setProperty('EDITEURS', JSON.stringify(a.adresses));
  ui.alert('Deux éditeurs du registre', 'Enregistrés :\n' + a.adresses.join('\n') +
    '\n\nÉtape suivante : Administration > Appliquer les droits et protections.', ui.ButtonSet.OK);
}

function analyserAdresses_(texte) {
  const brut = String(texte || '').split(/[,;\s]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
  const uniques = brut.filter((x, i) => brut.indexOf(x) === i);
  const invalides = uniques.filter(x => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
  if (invalides.length) return { erreur: 'Adresse invalide : ' + invalides.join(', ') };
  if (uniques.length !== 2) return { erreur: 'Saisir exactement deux adresses (' + uniques.length + ' reçue(s)).' };
  return { adresses: uniques };
}

function editeurs_() {
  try {
    const v = JSON.parse(proprietes_().getProperty('EDITEURS') || '[]');
    return Array.isArray(v) ? v : [];
  } catch (e) {
    return [];
  }
}

function appliquerDroits_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const editeurs = editeurs_();
  if (editeurs.length !== 2) throw new Error('Configurer d\'abord les deux éditeurs (Administration > Configurer les deux éditeurs).');
  const moi = String(Session.getEffectiveUser().getEmail()).toLowerCase();
  const prop = ss.getOwner() ? String(ss.getOwner().getEmail()).toLowerCase() : '';
  const actuels = ss.getEditors().map(u => String(u.getEmail()).toLowerCase()).filter(Boolean);
  const aAjouter = editeurs.filter(e => actuels.indexOf(e) < 0);
  const enTrop = actuels.filter(e => editeurs.indexOf(e) < 0 && e !== prop && e !== moi);
  const fichier = DriveApp.getFileById(ss.getId());
  let lienModif = false;
  try {
    lienModif = fichier.getSharingAccess() !== DriveApp.Access.PRIVATE && fichier.getSharingPermission() === DriveApp.Permission.EDIT;
  } catch (e) { lienModif = false; }
  const plan = ['Éditeurs autorisés : ' + editeurs.join(', ')];
  if (aAjouter.length) plan.push('- ajout en éditeur : ' + aAjouter.join(', '));
  if (enTrop.length) plan.push('- passage en lecture seule : ' + enTrop.join(', '));
  if (lienModif) plan.push('- lien de partage en modification ramené en lecture');
  plan.push('- les éditeurs ne peuvent plus partager le fichier');
  plan.push('- Registre et Données : seules les cellules de saisie (fond crème) restent modifiables');
  plan.push('- Dashboard et Journal : lecture seule');
  if (editeurs.indexOf(moi) < 0 && moi !== prop) plan.push('Vous (' + moi + ') restez éditeur en tant qu\'administrateur.');
  if (ui.alert('Droits et protections', plan.join('\n') + '\n\nConfirmer ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  aAjouter.forEach(e => ss.addEditor(e));
  enTrop.forEach(e => { ss.removeEditor(e); ss.addViewer(e); });
  if (lienModif) fichier.setSharing(fichier.getSharingAccess(), DriveApp.Permission.VIEW);
  try { fichier.setShareableByEditors(false); } catch (e) { plan.push('Restriction de partage non appliquée : ' + e.message); }
  protections_(ss);
  ui.alert('Droits et protections', 'Appliqués.\n\n' + plan.join('\n'), ui.ButtonSet.OK);
}

function protections_(ss) {
  const reg = ss.getSheetByName(CFG.registre), don = ss.getSheetByName(CFG.donnees), tab = ss.getSheetByName(CFG.tableau);
  proteger_(reg, 'registre : seules les cellules de saisie (fond crème) sont modifiables', plagesCreme_(reg));
  proteger_(don, 'données : seules les cellules de saisie (fond crème) sont modifiables', plagesCreme_(don));
  proteger_(tab, 'tableau de bord en lecture seule', []);
  proteger_(feuilleJournal_(ss), 'journal des modifications en lecture seule', []);
}

function proteger_(sh, description, exceptions) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(pr => {
    if (String(pr.getDescription()).indexOf(CFG.protection) === 0) pr.remove();
  });
  const pr = sh.protect().setDescription(CFG.protection + description);
  if (exceptions.length) pr.setUnprotectedRanges(exceptions);
  const moi = Session.getEffectiveUser();
  pr.addEditor(moi);
  pr.removeEditors(pr.getEditors().filter(u => u.getEmail() !== moi.getEmail()));
  if (pr.canDomainEdit()) pr.setDomainEdit(false);
  return pr;
}

function plagesCreme_(sh) {
  const bg = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).getBackgrounds();
  return segmentsCreme_(bg).map(s => sh.getRange(s[0], s[1], s[2], 1));
}

/** Segments verticaux de cellules crème : [ligne, colonne, nombre de lignes] (base 1). */
function segmentsCreme_(bg) {
  const out = [];
  const nl = bg.length, nc = nl ? bg[0].length : 0;
  for (let c = 0; c < nc; c++) {
    let debut = -1;
    for (let r = 0; r <= nl; r++) {
      const creme = r < nl && String(bg[r][c]).toLowerCase() === CFG.couleurs.creme;
      if (creme && debut < 0) debut = r;
      if (!creme && debut >= 0) {
        out.push([debut + 1, c + 1, r - debut]);
        debut = -1;
      }
    }
  }
  return out;
}

function activerAutomatisations_() {
  const ss = SpreadsheetApp.getActive();
  proprietes_().setProperty('ID_CLASSEUR', ss.getId());
  supprimerDeclencheurs_();
  ScriptApp.newTrigger('journaliser').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('envoyerRecapitulatif').timeBased().everyDays(1).atHour(CFG.heureRecap).inTimezone(CFG.fuseau).create();
  ScriptApp.newTrigger('surveillerNotifications').timeBased().everyHours(1).create();
  feuilleJournal_(ss);
  const ui = SpreadsheetApp.getUi();
  ui.alert('Journal et alertes', 'Activés :\n- journal de chaque modification du registre et des données (onglet masqué Journal)\n' +
    '- récapitulatif des échéances chaque jour vers ' + CFG.heureRecap + ' h\n' +
    '- surveillance horaire des notifications aux autorités\nDestinataires : ' + destinataires_().join(', '), ui.ButtonSet.OK);
}

function desactiverAutomatisations_() {
  const n = supprimerDeclencheurs_();
  SpreadsheetApp.getActive().toast(n + ' déclencheur(s) supprimé(s).', 'Journal et alertes', 5);
}

function supprimerDeclencheurs_() {
  const noms = ['journaliser', 'envoyerRecapitulatif', 'surveillerNotifications'];
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (noms.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

function appliquerCharte_() {
  const ss = SpreadsheetApp.getActive();
  const echecs = [];
  ss.getSheets().forEach(s => {
    try {
      s.setFrozenRows(0);
      s.setFrozenColumns(0);
      s.setHiddenGridlines(true);
      s.getDataRange().setFontFamily(CFG.police);
    } catch (e) {
      echecs.push(s.getName());
    }
  });
  ss.toast(echecs.length ? 'Onglets protégés non traités : ' + echecs.join(', ') : 'Calibri, sans volets figés ni quadrillage.',
    'Charte', 6);
}

// ------------------------------------------------------------------------------------- lecture
function classeur_() {
  const ss = SpreadsheetApp.getActive();
  if (ss) return ss;
  const id = proprietes_().getProperty('ID_CLASSEUR');
  if (!id) throw new Error('Classeur introuvable : relancer Administration > Activer le journal et les alertes.');
  return SpreadsheetApp.openById(id);
}

function proprietes_() {
  return PropertiesService.getScriptProperties();
}

function contexte_() {
  const ss = classeur_();
  const sh = ss.getSheetByName(CFG.registre);
  if (!sh) throw new Error('Onglet « ' + CFG.registre + ' » introuvable.');
  const nc = sh.getLastColumn();
  const ix = indexColonnes_(sh.getRange(CFG.ligneEntete, 1, 1, nc).getDisplayValues()[0]);
  if (ix.manquants.length) throw new Error('Colonnes introuvables en ligne ' + CFG.ligneEntete + ' : ' + ix.manquants.join(', '));
  const valeurs = sh.getRange(CFG.premiere, 1, CFG.derniere - CFG.premiere + 1, nc).getValues();
  return { ss: ss, sh: sh, idx: ix.idx, valeurs: valeurs, p: parametres_(ss) };
}

function parametres_(ss) {
  const p = {};
  PARAMS.forEach(n => {
    const r = ss.getRangeByName(n);
    if (!r) throw new Error('Plage nommée « ' + n + ' » introuvable (onglet Données).');
    p[n] = r.getNumRows() * r.getNumColumns() > 1 ? r.getValues().map(x => x[0]) : r.getValue();
  });
  return p;
}

function indexColonnes_(entetes) {
  const idx = {};
  entetes.forEach((h, i) => {
    const k = norm_(h);
    if (k && !(k in idx)) idx[k] = i;
  });
  const requis = Object.keys(CH).map(k => CH[k]).concat(SAISIES);
  const manquants = requis.filter((l, i) => requis.indexOf(l) === i && !(norm_(l) in idx));
  return { idx: idx, manquants: manquants };
}

function lignes_(valeurs, idx) {
  const cles = Object.keys(CH);
  const pos = SAISIES.map(l => idx[norm_(l)]);
  return valeurs.map((row, i) => {
    const o = { ligne: CFG.premiere + i };
    cles.forEach(k => {
      let v = row[idx[norm_(CH[k])]];
      if (typeof v === 'string') v = v.trim();
      if (DATES.indexOf(k) >= 0 && typeof v === 'number' && v > 0) v = serieVersDate_(v);
      o[k] = v;
    });
    o.vide = pos.every(c => vide_(row[c]));
    return o;
  });
}

// ------------------------------------------------------------------------------------- outils
function norm_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
}

function vide_(v) {
  return v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

function estDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function jour_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function ajouterHeures_(d, h) {
  return new Date(d.getTime() + h * 36e5);
}

function ajouterJours_(d, n) {
  const x = jour_(d);
  x.setDate(x.getDate() + n);
  return x;
}

/** Équivalent de SERIE.JOUR.OUVRE / WORKDAY(ENT(d); n) : week-ends exclus, jours fériés non exclus. */
function ajouterJoursOuvres_(d, n) {
  const x = jour_(d);
  let k = 0;
  while (k < n) {
    x.setDate(x.getDate() + 1);
    if (x.getDay() !== 0 && x.getDay() !== 6) k++;
  }
  return x;
}

/** Équivalent de MOIS.DECALER / EDATE(ENT(d); n). */
function ajouterMois_(d, n) {
  const x = jour_(d);
  const j = x.getDate();
  x.setDate(1);
  x.setMonth(x.getMonth() + n);
  x.setDate(Math.min(j, new Date(x.getFullYear(), x.getMonth() + 1, 0).getDate()));
  return x;
}

function serieVersDate_(n) {
  const u = new Date(Math.round((n - 25569) * 864e5));
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
}

function fmt_(d, horaire) {
  return Utilities.formatDate(d, CFG.fuseau, horaire ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy');
}

function a1_(r, c) {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s + r;
}

function compterLignes_(anos) {
  return anos.map(a => a.ligne).filter((l, i, t) => t.indexOf(l) === i).length;
}

function esc_(v) {
  return String(v === undefined || v === null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Page HTML sobre (Calibri 9, en-têtes #218E8E, filets #E0E0E0, sans quadrillage) pour les fenêtres. */
function page_(titre, sousTitre, colonnes, lignes, siVide, options) {
  const c = CFG.couleurs;
  const css = 'body{font-family:Calibri,Carlito,Arial,sans-serif;font-size:9pt;color:' + c.texte + ';margin:12px}' +
    'h1{font-size:10pt;color:' + c.bandeau + ';margin:0 0 4px}p{color:' + c.gris + ';margin:0 0 10px}' +
    'table{border-collapse:collapse;width:100%}th{background:' + c.entete + ';color:#fff;text-align:left;padding:4px 6px}' +
    'td{padding:4px 6px;border-bottom:1px solid ' + c.filet + ';vertical-align:top}' +
    'a{color:' + c.entete + ';cursor:pointer;text-decoration:none;font-weight:bold}' +
    '.retard td{font-weight:bold}.retard td.etat{background:' + c.cle + '}.avenir td{color:' + c.gris + '}';
  const corps = lignes.map(l => {
    const classe = l.etat === 'En retard' ? 'retard' : l.etat === 'À venir' ? 'avenir' : '';
    return '<tr class="' + classe + '">' + l.cellules.map((v, i) => {
      const cl = colonnes[i] === 'État' ? ' class="etat"' : '';
      if (i === 0 && l.ligne) {
        return '<td' + cl + '><a onclick="google.script.run.allerALigne(' + Number(l.ligne) + ')">' + esc_(v) + '</a></td>';
      }
      return '<td' + cl + '>' + esc_(v) + '</td>';
    }).join('') + '</tr>';
  }).join('');
  const lien = options && options.lien ? '<p><a href="' + esc_(options.lien) + '" target="_blank">' +
    esc_(options.libelleLien) + '</a></p>' : '';
  return '<!doctype html><html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' +
    '<h1>' + esc_(titre) + '</h1><p>' + esc_(sousTitre) + '</p>' + lien +
    (lignes.length ? '<table><tr>' + colonnes.map(h => '<th>' + esc_(h) + '</th>').join('') + '</tr>' + corps + '</table>'
      : '<p>' + esc_(siVide) + '</p>') + '</body></html>';
}
