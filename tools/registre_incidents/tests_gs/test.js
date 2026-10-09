'use strict';
// Banc de test du script Apps Script : maquette stricte des services Google et données du classeur.
// Usage (dans un dossier de travail) : python export_classeur.py classeur.xlsx classeur.json
//   puis TZ=Europe/Paris node test.js Registre_Incidents_Fundcraft_France.gs classeur.json
// Les fenêtres et le courriel générés sont écrits dans le dossier courant (controle.html, echeances.html...).
const vm = require('vm'), fs = require('fs');
const { creerEnv } = require('./mock.js');
let ok = 0, ko = 0;
const t = (nom, cond, info) => { if (cond) ok++; else { ko++; console.log('ÉCHEC', nom, info === undefined ? '' : JSON.stringify(info)); } };
const SCRIPT = fs.readFileSync(process.argv[2], 'utf8');
function charger(opts) {
  const { env, g } = creerEnv(process.argv[3] || 'classeur.json', Object.assign({ owner: 'risk@fundcraft.fr' }, opts || {}));
  const ctx = vm.createContext(g); vm.runInContext(SCRIPT, ctx);
  const run = code => vm.runInContext(code, ctx);
  return { env, ctx, run };
}
const H = 'Registre';
// ------------------------------------------------------------------ fonctions pures sur les données réelles
let { env, ctx, run } = charger();
const reg = env.ss.sheets.find(s => s.name === H);
const ent = reg.values[5];
const col = lib => ent.indexOf(lib) + 1;
const ix = run('indexColonnes_(SpreadsheetApp.getActive().getSheetByName("Registre").getRange(6,1,1,88).getDisplayValues()[0])');
t('libellés reconnus', ix.manquants.length === 0, ix.manquants);
const c = run('contexte_()');
const lignes = ctx.lignes_(c.valeurs, c.idx);
t('27 incidents lus', lignes.filter(o => !o.vide).length === 27, lignes.filter(o => !o.vide).length);
t('dates lues comme dates', Object.prototype.toString.call(lignes[0].det) === '[object Date]');
t('paramètres lus', c.p.Statut_Clos === 'Clôturé' && c.p.Notif_DORA === 'DORA - incident majeur' && c.p.Ref_Statut[0] === 'Ouvert' && c.p.DORA_H_INT === 72, c.p);
t('prochaine réf. 2026', ctx.prochaineRef_(lignes.map(o => o.ref), 2026) === 'INC-2026-028');
t('prochaine réf. 2027', ctx.prochaineRef_(lignes.map(o => o.ref), 2027) === 'INC-2027-001');
t('prochaine réf. trous et formats', ctx.prochaineRef_(['INC-2026-009', ' INC-2026-120 ', 'X-2026-500', 'INC-2025-900', '', null], 2026) === 'INC-2026-121');
const maintenant = new Date(2026, 9, 9, 10, 0);
const anos = ctx.anomalies_(lignes, c.p, maintenant);
const REG = run('REGLES_CONTROLE_');
const lignesRegle = new Set(anos.filter(a => REG.test(a.anomalie)).map(a => a.ligne));
const lignesAT = new Set(lignes.filter(o => o.ctrl === 'À compléter').map(o => o.ligne));
t('contrôle de saisie identique au classeur (données réelles)', [...lignesAT].every(l => lignesRegle.has(l)) && lignesAT.size === lignesRegle.size, [[...lignesAT], [...lignesRegle]]);
console.log('anomalies sur les 27 incidents :', anos.length, 'sur', new Set(anos.map(a => a.ligne)).size, 'lignes');
const parType = {}; anos.forEach(a => { const k = a.anomalie.replace(/\(lignes.*\)/, ''); parType[k] = (parType[k] || 0) + 1; });
console.log(parType);
// cohérence escalade : AS = 1 <=> date d'escalade <= WORKDAY(INT(J), délai) et niveau atteint >= requis
let incoh = [];
lignes.filter(o => !o.vide && o.dEsc instanceof Object && typeof o.escOK === 'number').forEach(o => {
  const due = ctx.ajouterJoursOuvres_(o.det, Number(o.delaiEsc));
  const conf = ctx.jour_(o.dEsc) <= due && Number(o.attRang) >= Number(o.reqRang) ? 1 : 0;
  if (conf !== o.escOK) incoh.push([o.ligne, conf, o.escOK]);
});
t('délai d\'escalade (jours ouvrés) identique au classeur', incoh.length === 0, incoh);
t('EDATE fin de mois', ctx.ajouterMois_(new Date(2026, 0, 31, 15), 1).getTime() === new Date(2026, 1, 28).getTime());
t('WORKDAY vendredi + 1', ctx.ajouterJoursOuvres_(new Date(2026, 9, 9, 18), 1).getTime() === new Date(2026, 9, 12).getTime());
t('WORKDAY 0', ctx.ajouterJoursOuvres_(new Date(2026, 9, 10, 8), 0).getTime() === new Date(2026, 9, 10).getTime());
t('série Excel -> date', ctx.serieVersDate_(46304.5).getTime() === new Date(2026, 9, 9, 12, 0).getTime(), ctx.serieVersDate_(46304.5));
// segments crème
const seg = ctx.segmentsCreme_(reg.bgs);
t('36 plages de saisie dans le registre, lignes 7 à 306', seg.length === 36 && seg.every(s => s[0] === 7 && s[2] === 300), seg.slice(0, 3));
const don = env.ss.sheets.find(s => s.name === 'Données');
const segD = ctx.segmentsCreme_(don.bgs);
t('cellules crème de Données toutes couvertes', segD.reduce((n, s) => n + s[2], 0) === don.bgs.flat().filter(x => x === '#fff8f2').length, segD.length);
// ------------------------------------------------------------------ échéances sur lignes synthétiques
function ligne(o) { return Object.assign({ vide: false, ligne: 99, ref: 'INC-T', fonds: 'F', statut: 'Ouvert' }, o); }
const p = c.p;
const J = new Date(2026, 9, 8, 10, 0);
let e = ctx.echeances_([ligne({ det: J, notif: p.Notif_DORA, clas: new Date(2026, 9, 8, 15, 0), doraEch: new Date(2026, 9, 8, 19, 0) })], p, new Date(2026, 9, 8, 17, 0));
t('DORA initiale sous 24 h', e[0].nature === 'DORA - notification initiale' && e[0].etat === 'Sous 24 h', e);
e = ctx.echeances_([ligne({ det: J, notif: p.Notif_DORA, clas: new Date(2026, 9, 8, 15, 0), doraEch: new Date(2026, 9, 8, 19, 0) })], p, new Date(2026, 9, 8, 19, 1));
t('DORA initiale en retard', e[0].etat === 'En retard');
e = ctx.echeances_([ligne({ det: J, notif: p.Notif_DORA, notifReal: new Date(2026, 9, 8, 18, 30), doraEch: new Date(2026, 9, 8, 19, 0) })], p, new Date(2026, 9, 9, 9, 0));
t('DORA intermédiaire 72 h', e[0].nature === 'DORA - rapport intermédiaire' && e[0].echeance.getTime() === new Date(2026, 9, 11, 18, 30).getTime() && e[0].etat === 'À venir', e);
e = ctx.echeances_([ligne({ det: J, notif: p.Notif_DORA, notifReal: new Date(2026, 9, 8, 18, 30), inter: new Date(2026, 9, 10, 9, 0), doraEch: new Date(2026, 9, 8, 19, 0) })], p, new Date(2026, 10, 9, 9, 0));
t('DORA final 1 mois, sous 7 jours', e[0].nature === 'DORA - rapport final' && e[0].echeance.getTime() === new Date(2026, 10, 10).getTime() && e[0].etat === 'Sous 7 jours', e);
e = ctx.echeances_([ligne({ det: J, notif: 'AMF', echNotif: '' })], p, maintenant);
t('notification sans délai paramétré', e[0].etat === 'Sans échéance');
e = ctx.echeances_([ligne({ det: J, notif: 'RGPD - CNIL', echNotif: new Date(2026, 9, 11, 10, 0), notifReal: new Date(2026, 9, 9, 9, 0) })], p, maintenant);
t('notification réalisée : plus d\'échéance', e.filter(x => /Notification/.test(x.nature)).length === 0, e);
e = ctx.echeances_([ligne({ det: J, reqRang: 2, attRang: 1, dEsc: new Date(2026, 9, 8), delaiEsc: 2, nivReq: 'Risques et conformité' })], p, maintenant);
t('escalade insuffisante', e.some(x => x.nature === 'Escalade : Risques et conformité' && x.echeance.getTime() === new Date(2026, 9, 12).getTime()), e);
e = ctx.echeances_([ligne({ det: J, reqRang: 2, attRang: 2, dEsc: new Date(2026, 9, 8), delaiEsc: 2 })], p, maintenant);
t('escalade réalisée : plus d\'échéance', !e.some(x => /Escalade/.test(x.nature)));
e = ctx.echeances_([ligne({ det: new Date(2026, 5, 1, 9), clot: '' })], p, maintenant);
t('ancienneté > 90 jours signalée', e.some(x => /Ancienneté/.test(x.nature) && x.etat === 'En retard'), e);
e = ctx.echeances_([ligne({ det: new Date(2026, 9, 1, 9) })], p, maintenant);
t('ancienneté lointaine non signalée', !e.some(x => /Ancienneté/.test(x.nature)));
e = ctx.echeances_([ligne({ det: J, sc: p.SC_Recl })], p, maintenant);
t('réclamation : AR 10 j.o. et réponse 2 mois', e.some(x => /accusé/.test(x.nature) && x.echeance.getTime() === new Date(2026, 9, 22).getTime()) && e.some(x => /réponse/.test(x.nature) && x.echeance.getTime() === new Date(2026, 11, 8).getTime()), e);
e = ctx.echeances_([ligne({ det: J, cat: p.Cat_Lim, typeDep: p.TypeDep_Actif })], p, maintenant);
t('dépassement actif 5 jours', e.some(x => /Dépassement/.test(x.nature) && x.echeance.getTime() === new Date(2026, 9, 13).getTime()), e);
e = ctx.echeances_([ligne({ det: J, ech: new Date(2026, 9, 1), resp: 'Middle office', clot: new Date(2026, 9, 8) })], p, maintenant);
t('action corrective en retard même après clôture', e.length === 1 && e[0].nature === 'Action corrective' && e[0].etat === 'En retard', e);
t('tri : en retard d\'abord', ctx.echeances_([ligne({ det: J, ech: new Date(2026, 11, 1), ligne: 1 }), ligne({ det: J, ech: new Date(2026, 9, 1), ligne: 2 })], p, maintenant)[0].ligne === 2);
// anomalies synthétiques
const a = (o) => ctx.anomalies_([ligne(o)], p, maintenant).map(x => x.anomalie);
t('futur et chronologie', (() => { const r = a({ det: J, fonds: 'F', sc: 'S', orig: 'O', surv: new Date(2026, 9, 9), pec: new Date(2026, 9, 8, 9), res: new Date(2026, 9, 10, 9) }); return r.includes('Survenance postérieure à la détection') && r.includes('Prise en compte antérieure à la détection') && r.includes('Date future : Résolution'); })(), a({ det: J, surv: new Date(2026, 9, 9), pec: new Date(2026, 9, 8, 9), res: new Date(2026, 9, 10, 9) }));
t('DORA incohérent', a({ det: J, fonds: 'F', sc: 'S', orig: 'O', notif: 'AMF', inter: new Date(2026, 9, 8, 12) }).filter(x => /DORA/.test(x)).length === 2);
t('doublon de réf.', ctx.anomalies_([ligne({ det: J, ligne: 7 }), ligne({ det: J, ligne: 8 })], p, maintenant).filter(x => /double/.test(x.anomalie)).length === 2);
t('cause racine des significatifs clôturés', a({ det: J, fonds: 'F', sc: 'S', orig: 'O', statut: 'Clôturé', clot: new Date(2026, 9, 9), res: new Date(2026, 9, 8, 12), gravRang: 3, cause: '' }).includes('Cause racine manquante (incident significatif ou majeur clôturé)'));
// ------------------------------------------------------------------ fonctions avec interface
run('onOpen()');
const m = env.ui.menus[0];
const fns = []; const parcours = mm => mm.items.forEach(i => { if (i[0] === '>') parcours(i[1]); else if (i[1]) fns.push(i[1]); });
parcours(m);
t('menu Registre', m.nom === 'Registre' && fns.length === 11, fns);
t('fonctions du menu définies', fns.every(f => typeof ctx[f] === 'function'), fns.filter(f => typeof ctx[f] !== 'function'));
run('nouvelIncident()');
t('nouvel incident sans erreur', !env.ui.alerts.length, env.ui.alerts);
t('nouvel incident ligne 34', reg._get(34, col('Réf.')) === 'INC-2026-028' && reg._get(34, col('Statut')) === 'Ouvert', [reg._get(34, col('Réf.')), reg._get(34, col('Statut'))]);
const decl = reg._get(34, col('Déclaration au registre'));
t('déclaration = aujourd\'hui à minuit', decl.getHours() === 0 && decl.getDate() === new Date().getDate(), decl);
t('cellule fonds sélectionnée', env.ss.actif && env.ss.actif[1] === 'C34', env.ss.actif);
t('verrou libéré', env.lockReleased === true);
const jr = env.ss.sheets.find(s => s.name === 'Journal');
t('journal créé, masqué, protégé', jr && jr.hidden && jr.protections.length === 1 && jr.values[1][4] === 'INC-2026-028', jr && jr.values[1]);
run('nouvelIncident()');
t('second incident INC-2026-029 ligne 35', reg._get(35, col('Réf.')) === 'INC-2026-029');
run('controlerRegistre()');
let d = env.ui.dialogs.pop();
t('fenêtre de contrôle', d[0] === 'modeless' && /Contrôle du registre/.test(d[2]) && /allerALigne\(34\)/.test(d[2]) && /29 incident/.test(d[2]), d[2].slice(0, 300));
fs.writeFileSync('controle.html', d[2]);
run('afficherEcheances()');
d = env.ui.dialogs.pop();
t('fenêtre des échéances', /Échéances et alertes/.test(d[2]) && /<table>/.test(d[2]));
fs.writeFileSync('echeances.html', d[2]);
run('allerALigne(12)');
t('aller à la ligne', env.ss.actif[1] === 'B12', env.ss.actif);
// journal : modification d'une cellule
const avant = jr.values.length;
const cel = env.ss.proxy.getSheetByName(H).getRange(12, col('Statut'));
run('journaliser')({ range: cel, source: env.ss.proxy, oldValue: 'Ouvert', value: 'Clôturé', user: { getEmail: () => 'delegue@fundcraft.fr' } });
t('journal : modification simple', jr.values.length === avant + 1 && jr.values[avant][1] === 'delegue@fundcraft.fr' && jr.values[avant][5] === 'Statut' && jr.values[avant][4] === 'INC-2026-006' && jr.values[avant][6] === 'Ouvert', jr.values[avant]);
const celD = env.ss.proxy.getSheetByName(H).getRange(12, col('Date de clôture'));
run('journaliser')({ range: celD, source: env.ss.proxy, oldValue: '46304', user: { getEmail: () => '' } });
t('journal : ancienne date convertie, utilisateur non communiqué', jr.values[avant + 1][6] === '09/10/2026' && jr.values[avant + 1][1] === 'non communiqué par Google', jr.values[avant + 1]);
run('journaliser')({ range: env.ss.proxy.getSheetByName(H).getRange(7, 2, 10, 8), source: env.ss.proxy });
t('journal : collage plafonné à 50 + 1 ligne de synthèse', jr.values.length === avant + 2 + 51, jr.values.length - avant);
const nJ = jr.values.length;
run('journaliser')({ range: env.ss.proxy.getSheetByName('Dashboard').getRange(5, 5), source: env.ss.proxy });
t('journal : Dashboard ignoré', jr.values.length === nJ);
run('journaliser')({ range: env.ss.proxy.getSheetByName('Données').getRange('C6'), source: env.ss.proxy, oldValue: '46295' });
t('journal : Données avec libellé de ligne', jr.values[nJ][5] === "Date d'arrêté" && jr.values[nJ][6] === '30/09/2026', jr.values[nJ]);
// configuration des éditeurs
env.ui.prompts.push(['OK', 'risk@fundcraft.fr ; delegue@fundcraft.fr']);
run('configurerEditeurs()');
t('éditeurs enregistrés', env.props.EDITEURS === '["risk@fundcraft.fr","delegue@fundcraft.fr"]', env.props);
env.ui.prompts.push(['OK', 'a@b.fr']); env.ui.alerts = [];
run('configurerEditeurs()');
t('refus si une seule adresse', /exactement deux/.test(JSON.stringify(env.ui.alerts)) && env.props.EDITEURS.includes('delegue'), env.ui.alerts);
// droits et protections
env.ss.editors.push('ancien@fundcraft.fr'); env.ui.reponses.push('YES'); env.ui.alerts = [];
run('appliquerDroits()');
t('droits : éditeur ajouté, ancien en lecture', env.ss.editors.includes('delegue@fundcraft.fr') && !env.ss.editors.includes('ancien@fundcraft.fr') && env.ss.viewers.includes('ancien@fundcraft.fr'), [env.ss.editors, env.ss.viewers]);
t('droits : lien en lecture, partage par les éditeurs bloqué', env.partage[1] === 'VIEW' && env.partageEditeurs === false, env.partage);
const prot = n => env.ss.sheets.find(s => s.name === n).protections;
t('protections : registre avec 36 plages de saisie', prot(H).length === 1 && prot(H)[0].getDescription().startsWith('Fundcraft - ') && prot(H)[0].unprot.length === 36, prot(H).map(p => p.unprot.length));
t('protections : Données, Dashboard, Journal', prot('Données')[0].unprot.length === segD.length && prot('Dashboard')[0].unprot.length === 0 && prot('Journal').length === 1);
t('protections : seul le propriétaire édite les formules', prot(H)[0].getEditors().map(u => u.getEmail()).join() === 'risk@fundcraft.fr' && !prot(H)[0].canDomainEdit());
env.ui.reponses.push('YES'); run('appliquerDroits()');
t('protections idempotentes', prot(H).length === 1 && prot('Journal').length === 1);
// automatisations
run('activerAutomatisations()'); run('activerAutomatisations()');
t('3 déclencheurs, sans doublon', env.triggers.length === 3 && env.triggers.some(x => x[0] === 'jour' && x[2] === 8 && x[3] === 'Europe/Paris'), env.triggers);
// alertes : ligne DORA urgente relative à maintenant
const now = new Date();
const r40 = 40;
const setv = (lib, v) => reg._set(r40, col(lib), v);
setv('Réf.', 'INC-2026-099'); setv('Fonds ou véhicule', 'Openstone Infraworld'); setv('Sous-catégorie', 'Cyberattaque ou intrusion');
setv('Détection (date et heure)', new Date(now.getTime() - 2 * 36e5)); setv('Notification à une autorité', 'DORA - incident majeur');
setv('DORA - échéance de notification initiale', new Date(now.getTime() + 2 * 36e5));
env.mails = [];
run('surveillerNotifications()');
t('alerte DORA envoyée aux deux éditeurs', env.mails.length === 1 && env.mails[0].to === 'risk@fundcraft.fr,delegue@fundcraft.fr' && /INC-2026-099/.test(env.mails[0].subject) && !/Cyberattaque/.test(env.mails[0].htmlBody), env.mails.map(x => x.subject));
fs.writeFileSync('alerte.html', env.mails[0].htmlBody);
run('surveillerNotifications()');
t('pas de doublon d\'alerte à l\'heure suivante', env.mails.length === 1);
setv('DORA - échéance de notification initiale', new Date(now.getTime() - 60e3));
run('surveillerNotifications()');
t('nouvelle alerte au passage en retard', env.mails.length === 2 && /En retard/.test(env.mails[1].htmlBody));
run('envoyerRecapitulatif()');
t('récapitulatif quotidien', env.mails.length === 3 && /échéance\(s\) en retard/.test(env.mails[2].subject), env.mails[2] && env.mails[2].subject);
console.log('récapitulatif :', env.mails[2].subject);
// diagnostic
run('diagnostic()');
d = env.ui.dialogs.pop();
fs.writeFileSync('diagnostic.html', d[2]);
const nonConf = (d[2].match(/À corriger/g) || []).length;
t('diagnostic : tout conforme sauf contrôle de saisie (lignes de test)', /Diagnostic après import/.test(d[2]), nonConf);
console.log('diagnostic, points à corriger :', nonConf, [...d[2].matchAll(/<tr class="retard"><td>([^<]*)<\/td><td>[^<]*<\/td><td>([^<]*)/g)].map(x => x[1] + ' : ' + x[2]));
// charte, journal, désactivation
env.ss.sheets[0].frozenR = 2; run('appliquerCharte()');
t('charte : volets libérés', env.ss.sheets.every(s => s.frozenR === 0 && s.frozenC === 0 && s.gridHidden));
run('basculerJournal()'); t('journal affiché', !jr.hidden); run('basculerJournal()'); t('journal masqué', jr.hidden);
// arrêté figé
env.ui.reponses.push('YES'); env.ui.alerts = [];
run('figerArrete()');
const copie = Object.values(env.classeurs).find(x => x.id !== 'ID1');
t('arrêté créé dans le sous-dossier', copie && env.racine.sous['Arrêtés du registre des incidents'].fichiers['Arrete_Registre_Incidents_Fundcraft_France_T3_2026'], Object.keys(env.racine.sous));
t('arrêté : valeurs collées sur chaque onglet', copie.log.filter(x => x[0] === 'copyTo').length === copie.sheets.length);
t('arrêté : métadonnée, note, protections', copie.meta.length === 1 && JSON.parse(copie.meta[0][1]).periode === 'T3 2026' && copie.sheets.every(s => s.protections.length === 1) && /Copie figée/.test(copie.sheets.find(s => s.name === 'Dashboard').notes.B2));
d = env.ui.dialogs.pop(); t('arrêté : fenêtre avec lien', /Ouvrir l&#39;arrêté figé/.test(d[2]) && /ID\d/.test(d[2]));
env.ui.reponses.push('YES', 'YES'); run('figerArrete()');
t('second arrêté horodaté', Object.keys(env.racine.sous['Arrêtés du registre des incidents'].fichiers).length === 2, Object.keys(env.racine.sous['Arrêtés du registre des incidents'].fichiers));
env.actif = copie; env.ui.menus = []; run('onOpen()');
t('menu réduit dans l\'arrêté', env.ui.menus[0].nom === 'Registre (arrêté figé)');
env.ui.alerts = []; run('infosArrete()'); t('informations de l\'arrêté', /30\/09\/2026/.test(env.ui.alerts[0][1]), env.ui.alerts);
env.actif = env.ss;
run('desactiverAutomatisations()'); t('déclencheurs supprimés', env.triggers.length === 0);
// erreurs lisibles
reg.values[5][col('Statut') - 1] = 'Etat'; env.ui.alerts = [];
run('nouvelIncident()');
t('colonne renommée : message clair', /Colonnes introuvables en ligne 6 : Statut/.test(env.ui.alerts[0][1]), env.ui.alerts);
// fuseau horaire du script différent
({ env, ctx, run } = charger({ tzScript: 'America/New_York' }));
run('diagnostic()'); d = env.ui.dialogs.pop();
t('diagnostic : fuseau du script signalé', /America\/New_York : Apps Script/.test(d[2]));
console.log(`\n${ok} tests réussis, ${ko} échec(s)`);
process.exit(ko ? 1 : 0);
