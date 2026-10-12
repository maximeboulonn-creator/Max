/* Test du moteur JS contre les valeurs calculées du classeur (openpyxl data_only). */
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..');
const fees = JSON.parse(fs.readFileSync(path.join(dir, 'fees_model.json'), 'utf8'));
const expected = JSON.parse(fs.readFileSync(path.join(__dirname, 'expected.json'), 'utf8'));
const helpers = fs.readFileSync(path.join(__dirname, 'helpers_stub.js'), 'utf8');
const mod = fs.readFileSync(path.join(dir, 'fees_module.js'), 'utf8');

const stubs = `
  const DATA = { fees: __FEES__ };
  const ID = { id: 'airfund-cvc', export_prefix: 'AirFund_CVC' };
  const SK = 'om:test:';
  const EDITS = {};
  const state = { tab: 'fees-tool' };
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const render = () => {};
  const offerFile = () => {};
  const localStorage = { getItem: () => null, setItem: () => {} };
  const document = { addEventListener: () => {}, querySelector: () => null };
  const window = { scrollY: 0, scrollTo: () => {} };
`;
const src = '(function(){\n' + stubs.replace('__FEES__', JSON.stringify(fees)) + helpers + mod + '\n  return { model, recapOf, tresoOf, detailOf, breakevenOf, lineAnnual, refsOf, renderFees, viewSaisie, viewRecap, viewDetail, viewProvision, viewTreso, viewBreakeven, feesCSV, getFS: () => FS, setOv: (p, v) => { FS.ov[p] = v; }, resetOv: () => { FS.ov = {}; FS.invoices = null; } };\n})()';
const E = eval(src);

let fails = 0, checks = 0;
const near = (a, b, tol) => Math.abs(a - b) <= (tol == null ? 0.011 : tol);
function check(label, got, exp, tol) { checks++; if (!near(got, exp, tol)) { fails++; console.log('FAIL', label, 'got', got, 'expected', exp); } }

const m = E.model();
const r = E.recapOf(m);
// 04_Recap : ladder + each line + totals
check('ladder length', r.ladder.length, expected.ladder.length);
expected.ladder.forEach((a, i) => check('ladder[' + i + ']', r.ladder[i], a));
expected.recap.forEach((row, li) => row.forEach((v, i) => check('recap ' + m.lines[li].id + ' @' + expected.ladder[i], r.rows[li].vals[i], v)));
expected.total.forEach((v, i) => check('total @' + expected.ladder[i], r.tot[i], v));
expected.pct.forEach((v, i) => check('pct @' + expected.ladder[i], r.pct[i], v, 1e-9));
// 03_Detail : écarts à zéro
const d = E.detailOf(m);
d.D.rows.forEach(x => check('detail dep ecart @' + x.a, x.ecart, 0));
d.V.rows.forEach(x => check('detail valo ecart @' + x.a, x.ecart, 0));
// 05_Tresorerie
const T = E.tresoOf(m);
check('prorata0', T.prorata0, expected.treso.prorata, 1e-6);
check('nb mois ex1', T.ex[0].nbm, expected.treso.ex1_nbm);
check('ex2 end', isoOfLocal(T.ex[1].end) === '2028-12-31' ? 1 : 0, 1);
function isoOfLocal(dd) { return dd.toISOString().slice(0, 10); }
check('anniv', isoOfLocal(T.anniv) === '2027-08-13' ? 1 : 0, 1);
Object.keys(expected.treso.annual).forEach((id, k) => { const row = T.rows.find(x => x.l.id === id); expected.treso.annual[id].forEach((v, i) => check('annual ' + id + ' m' + i, row.annual[i], v)); });
Object.keys(expected.treso.prov).forEach(id => { const row = T.rows.find(x => x.l.id === id); expected.treso.prov[id].forEach((v, i) => check('prov ' + id + ' m' + i, row.prov[i], v)); });
expected.treso.provC.forEach((v, i) => check('provC m' + i, T.provC[i], v));
expected.treso.cashC.forEach((v, i) => check('cashC m' + i, T.cashC[i], v));
// facturation : mensuel = provision du mois ; semestriel en décembre = somme août..décembre (le classeur ne somme qu'août+septembre faute d'AUM saisi ensuite)
const g = T.rows.find(x => x.l.id === 'gestion');
check('inv gestion m0', g.inv[0], g.prov[0]); check('inv gestion m1', g.inv[1], g.prov[1]);
const dep = T.rows.find(x => x.l.id === 'depositaire');
check('inv depositaire m0..m3 = 0', dep.inv[0] + dep.inv[1] + dep.inv[2] + dep.inv[3], 0);
check('inv depositaire dec26 = sum prov aug..dec', dep.inv[4], dep.prov.slice(0, 5).reduce((a, b) => a + b, 0));
check('inv depositaire jun27 = sum prov jan..jun', dep.inv[10], dep.prov.slice(5, 11).reduce((a, b) => a + b, 0));
check('cash depositaire jan27 = inv dec26', dep.cash[5], dep.inv[4]);
check('cash m0 = 0', T.cashT[0], 0);
const fin = T.rows.find(x => x.l.id === 'finologee');
check('finologee no prorata m0', fin.prov[0], 116);
// règle CAC par distribution : 2 distributions/an, classe distribuante -> facturé en juillet et décembre, montant = annual * 6/12
E.setOv('params.nb_distributions', 2); E.setOv('classes.0.distribuante', 'Oui');
const m2 = E.model(); const T2 = E.tresoOf(m2); const cd = T2.rows.find(x => x.l.id === 'cac_distribution');
check('cac_distribution annual', cd.annual[0], 3200);
check('cac_distribution freq semestriel', cd.freq === 'Semestriel' ? 1 : 0, 1);
check('cac_distribution inv aug26', cd.inv[0], 0);
check('cac_distribution inv dec26', cd.inv[4], 3200 * 6 / 12);
check('cac_distribution inv jul27', cd.inv[11], 3200 * 6 / 12);
check('cac_distribution inv jun27', cd.inv[10], 0);
E.resetOv();
// surcharge : plancher gestion 0 -> à 5 M€ la commission = 0,18 % x 5 M€ = 9 000
E.setOv('lines.0.plancher', 0);
check('override plancher gestion', E.recapOf(E.model()).rows[0].vals[0], 9000);
E.resetOv();
// classe B : sous-distributeurs 0
E.setOv('classe_selectionnee', 'B');
check('classe B sous-distrib', E.recapOf(E.model()).rows[3].vals[0], 0);
check('classe B distrib global', E.recapOf(E.model()).rows[2].vals[0], 5000);
E.resetOv();
// valorisation : 3 classes -> +10 % ; bi-mensuelle ; 25 lignes -> +1000
E.setOv('classes.2.nom', 'C'); E.setOv('valo.frequence_vl', 'Bi-mensuelle'); E.setOv('valo.nb_lignes', 25);
check('valo majorée', E.recapOf(E.model()).rows[8].vals[0], 17000 * 1.1 + 1000);
E.resetOv();
// rendu HTML de chaque vue (pas d'exception)
['saisie', 'recap', 'detail', 'provision', 'treso', 'breakeven'].forEach(k => { E.getFS().mode = k; const h = E.renderFees({}); checks++; if (!/<div class="card"/.test(h) || /flagbox">Le modèle de frais n’a pas pu/.test(h)) { fails++; console.log('FAIL render', k, h.slice(0, 300)); } });
const csv = E.feesCSV('recap'); checks++; if (!/TOTAL/.test(csv)) { fails++; console.log('FAIL csv'); }
const csv2 = E.feesCSV('treso'); checks++; if (!/Provision comptable/.test(csv2)) { fails++; console.log('FAIL csv treso'); }
console.log(checks + ' contrôles, ' + fails + ' échec(s)');
process.exit(fails ? 1 : 0);
