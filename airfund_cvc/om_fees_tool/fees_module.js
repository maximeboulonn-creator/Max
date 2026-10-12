  /* =====================================================================
     Modèle de frais - port fidèle du classeur Fundcraft_AirFund_Modele_Frais_3.xlsx
     (01_Input, 02_Classes_de_Parts, 03_Detail_Depositaire_Valo, 04_Recap, 05_Tresorerie, 06_Documentation).
     Les valeurs du classeur sont dans le bloc de données (DATA.fees) ; chaque saisie de l'utilisateur est une
     surcharge FS.ov[chemin] = valeur, conservée dans ce navigateur et dans le fichier (« Enregistrer le fichier »).
     Le moteur ci-dessous reproduit les formules du classeur : barème Forfait / Taux / Palier à cinq tranches,
     plancher et plafond annuels, modes « Taux Classe sélectionnée » et « Frais KYC », échelle d'AUM, prorata du
     mois de lancement en Jours Ouvrés, facturation par échéance, décaissement le mois suivant, et ré-amortissement
     des Frais de Constitution au fil des factures réelles.
     ===================================================================== */
  const FEES = DATA.fees || { lines: [], classes: [], params: {}, params_meta: {}, treso: {}, echelle: {}, valo: {}, baremes: {} };
  const M = 1000000, FINF = 999999999999;
  const FS_KEY = SK + 'fees.v3';
  const FMODES = ['saisie', 'recap', 'detail', 'provision', 'treso', 'breakeven'];
  const FS_EMPTY = () => ({ v: 3, ov: {}, mode: 'recap', aum: null, pmonth: null, invoices: null, envShift: 0 });
  let FS = FS_EMPTY();
  try { const s = JSON.parse(localStorage.getItem(FS_KEY) || 'null'); if (s && s.v === 3) FS = Object.assign(FS_EMPTY(), s, { envShift: 0 }); } catch (e) { /* ignore */ }
  try { if (EDITS.fees && EDITS.fees.v === 3 && !localStorage.getItem(FS_KEY)) FS = Object.assign(FS_EMPTY(), EDITS.fees, { envShift: 0 }); } catch (e) { /* ignore */ }
  function fsNorm() { if (!FS || typeof FS !== 'object') FS = FS_EMPTY(); if (!FS.ov || typeof FS.ov !== 'object') FS.ov = {}; if (!FMODES.includes(FS.mode)) FS.mode = 'recap'; if (FS.invoices != null && !Array.isArray(FS.invoices)) FS.invoices = null; }
  fsNorm();
  function saveFees() { try { localStorage.setItem(FS_KEY, JSON.stringify(Object.assign({}, FS, { envShift: 0 }))); } catch (e) { /* ignore */ } }

  /* ---------------- formats ---------------- */
  const DEC = ',';
  const eurF = n => n == null ? '-' : new Intl.NumberFormat('en-GB', { maximumFractionDigits: 0 }).format(Math.round(n)).replace(/,/g, ' ');
  const fEur = n => (n == null || isNaN(n)) ? '-' : (n < 0 ? '−' : '') + eurF(Math.abs(n));
  const fEurZ = n => (n == null || isNaN(n) || Math.abs(n) < 0.5) ? '<span class="muted">·</span>' : fEur(n);
  const fPct = (x, d) => (x == null || isNaN(x)) ? '-' : (x * 100).toFixed(d == null ? 2 : d).replace('.', DEC) + ' %';
  const fRate = x => (x == null || isNaN(x)) ? '-' : (+(x * 100).toFixed(4)).toString().replace('.', DEC) + ' %';
  const fBp = x => (x == null || isNaN(x)) ? '-' : (x * 10000).toFixed(1).replace('.', DEC) + ' bp';
  const fM = v => { if (v == null || isNaN(v)) return '-'; const x = v / M; let s = Math.abs(x) >= 100 ? x.toFixed(0) : Math.abs(x) >= 10 ? x.toFixed(1) : x.toFixed(2); if (s.includes('.')) s = s.replace(/\.?0+$/, ''); return s.replace('.', DEC) + ' M€'; };
  const fDate = iso => { const d = parseISO(iso); return d ? String(d.getUTCDate()).padStart(2, '0') + '/' + String(d.getUTCMonth() + 1).padStart(2, '0') + '/' + d.getUTCFullYear() : (iso || '-'); };
  const MFR = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
  const ymF = ym => MFR[+ym.slice(5, 7) - 1] + ' ' + ym.slice(2, 4);
  const muted = s => '<span class="muted">' + s + '</span>';

  /* ---------------- modèle effectif = classeur + surcharges ---------------- */
  function getPath(o, p) { return String(p).split('.').reduce((a, k) => (a == null ? undefined : a[k]), o); }
  function setPath(o, p, v) { const ks = String(p).split('.'); let a = o; for (let i = 0; i < ks.length - 1; i++) { const k = ks[i]; if (a[k] == null || typeof a[k] !== 'object') a[k] = /^\d+$/.test(ks[i + 1]) ? [] : {}; a = a[k]; } a[ks[ks.length - 1]] = v; }
  function model() { const m = JSON.parse(JSON.stringify(FEES)); const ov = FS.ov || {}; Object.keys(ov).forEach(p => setPath(m, p, ov[p])); return m; }
  const isOv = p => Object.prototype.hasOwnProperty.call(FS.ov || {}, p) && JSON.stringify(FS.ov[p]) !== JSON.stringify(getPath(FEES, p));
  const nOv = () => Object.keys(FS.ov || {}).filter(isOv).length;
  const isRef = x => !!x && typeof x === 'object' && !!x.ref;
  const refLabel = r => (FEES.ref_labels || {})[r] || r;

  /* valeurs pilotées (cellules grises du classeur) */
  function refsOf(m) {
    const p = m.params || {}, b = m.baremes || {}, v = m.valo || {};
    const cls = (m.classes || []).find(c => String(c.nom || '').trim() !== '' && String(c.nom) === String(m.classe_selectionnee || '')) || null;
    const nbCat = (m.classes || []).filter(c => String(c.nom || '').trim() !== '').length;
    let maj = 0; (v.majoration || []).forEach(x => { if (!maj && nbCat >= +x[0]) maj = +x[1]; });
    const pal = (v.paliers || {})[v.frequence_vl] || (v.paliers || {}).Mensuelle || [0, 0, 0, 0];
    const sup = v.supplement || {}; const nl = +v.nb_lignes || 0; const inc = +sup.lignes_incluses || 10, par = +sup.par_tranche || 10;
    const supp = nl <= inc ? 0 : Math.ceil((nl - inc) / par) * (+sup.montant || 0);
    const g = b.gestion || [], c = b.conseil || [];
    return {
      seuil_t1: +b.seuil_t1 || 0, seuil_t2: +b.seuil_t2 || 0,
      gestion_t1: +g[0] || 0, gestion_t2: +g[1] || 0, gestion_t3: +g[2] || 0,
      conseil_t1: +c[0] || 0, conseil_t2: +c[1] || 0, conseil_t3: +c[2] || 0,
      distrib_global_classe: cls ? (+cls.distrib_global || 0) : 0, distrib_sous_classe: cls ? (+cls.sous_distrib || 0) : 0,
      classe_distribue: cls ? (cls.distribuante || 'Non') : 'Non',
      cac_distribution: (cls && cls.distribuante === 'Oui') ? (+p.nb_distributions || 0) * (+p.cout_par_distribution_cac || 0) : 0,
      valo_p1: (+pal[0] || 0) * (1 + maj), valo_p2: (+pal[1] || 0) * (1 + maj), valo_p3: (+pal[2] || 0) * (1 + maj), valo_p4: (+pal[3] || 0) * (1 + maj),
      valo_supplement: supp,
      tenue_passif_variable: (+p.cout_marginal_investisseur || 0) * (+p.nb_investisseurs || 0),
      constitution_montant: Math.max(+p.plancher_constitution || 0, (+p.taux_constitution || 0) * (+p.actif_net_prev_an1 || 0)),
      constitution_plancher: +p.plancher_constitution || 0,
      nb_categories: nbCat, majoration: maj, classe: cls
    };
  }
  function val(x, R, dflt) { if (isRef(x)) return +R[x.ref] || 0; if (x == null || x === '') return dflt; const n = +x; return isNaN(n) ? dflt : n; }

  /* ---------------- moteur : 04_Recap ---------------- */
  function lineAnnual(l, aum, m, R) {
    if (l.actif !== 'Oui') return 0;
    const p = m.params || {};
    const plancher = val(l.plancher, R, 0), plafond = val(l.plafond, R, FINF);
    let x = 0;
    if (l.mode === 'residuel') x = 0;
    else if (l.mode === 'classe') x = (l.classe_taux === 'sous' ? R.distrib_sous_classe : R.distrib_global_classe) * aum;
    else if (l.mode === 'kyc') x = Math.max(0, (+p.nb_investisseurs || 0) - (+p.nb_kyc_gratuites || 0) * (Math.floor(aum / ((+p.taille_tranche_kyc) || FINF)) + 1)) * (+p.cout_kyc_supp || 0);
    else (l.tranches || []).forEach(t => {
      const inf = val(t.inf, R, 0), sup = val(t.sup, R, FINF), v = val(t.valeur, R, 0);
      if (t.nature === 'Forfait') x += aum >= inf ? v : 0;
      else if (t.nature === 'Palier') x += (aum >= inf && aum < sup) ? v : 0;
      else x += Math.max(Math.min(aum, sup) - inf, 0) * v;
    });
    return Math.min(plafond, Math.max(plancher, x));
  }
  function ladderOf(m) {
    const e = m.echelle || {}; const mn = +e.min || 0, mx = +e.max || 0, st = +e.pas || 0, cap = +e.max_paliers || 150;
    const out = []; if (!(mn > 0)) return out; out.push(mn);
    if (st > 0) for (let a = mn + st; a <= mx + 1e-6 && out.length < cap; a += st) out.push(a);
    return out;
  }
  const ladderBad = m => { const e = m.echelle || {}; return !(+e.min > 0) || !(+e.pas > 0) || +e.min >= +e.max; };
  function recapOf(m) {
    const R = refsOf(m); const lad = ladderOf(m);
    const rows = (m.lines || []).map(l => ({ l, vals: lad.map(a => lineAnnual(l, a, m, R)) }));
    const tot = lad.map((a, i) => rows.reduce((s, r) => s + r.vals[i], 0));
    const vat = lad.map((a, i) => rows.reduce((s, r) => s + r.vals[i] * (+r.l.tva || 0), 0));
    return { R, ladder: lad, rows, tot, vat, pct: lad.map((a, i) => a > 0 ? tot[i] / a : null) };
  }
  function totalAt(m, aum, R) { R = R || refsOf(m); let t = 0, v = 0; (m.lines || []).forEach(l => { const x = lineAnnual(l, aum, m, R); t += x; v += x * (+l.tva || 0); }); return { total: t, vat: v }; }
  const lineById = (m, id) => (m.lines || []).find(l => l.id === id) || null;

  /* ---------------- moteur : 03_Detail ---------------- */
  function detailOf(m) {
    const R = refsOf(m); const lad = ladderOf(m);
    const dep = lineById(m, 'depositaire'), valo = lineById(m, 'valorisation');
    const D = dep ? { l: dep, taux: val(((dep.tranches || [])[0] || {}).valeur, R, 0), plancher: val(dep.plancher, R, 0), plafond: val(dep.plafond, R, FINF) } : null;
    if (D) D.rows = lad.map(a => { const th = a * D.taux; const ret = Math.min(D.plafond, Math.max(D.plancher, th)); return { a, th, mec: th < D.plancher ? 'Plancher (min.)' : th > D.plafond ? 'Plafond (max.)' : 'Barème (plein)', ret, ecart: Math.round((ret - lineAnnual(dep, a, m, R)) * 100) / 100 }; });
    const V = valo ? { l: valo, pal: (valo.tranches || []).filter(t => t.nature === 'Palier'), forf: (valo.tranches || []).filter(t => t.nature === 'Forfait') } : null;
    if (V) V.rows = lad.map(a => {
      const t = V.pal.find(x => a >= val(x.inf, R, 0) && a < val(x.sup, R, FINF)) || null;
      const lab = t ? (val(t.sup, R, FINF) >= FINF ? '> ' + fM(val(t.inf, R, 0)) : fM(val(t.inf, R, 0)) + ' – ' + fM(val(t.sup, R, FINF))) : 'hors barème';
      const mt = t ? val(t.valeur, R, 0) : 0; const sp = V.forf.reduce((s, x) => s + (a >= val(x.inf, R, 0) ? val(x.valeur, R, 0) : 0), 0);
      const ret = Math.min(val(valo.plafond, R, FINF), Math.max(val(valo.plancher, R, 0), mt + sp));
      return { a, lab, mt, sp, ret, ecart: Math.round((ret - lineAnnual(valo, a, m, R)) * 100) / 100 };
    });
    return { R, ladder: lad, D, V };
  }

  /* ---------------- moteur : 05_Tresorerie ---------------- */
  const dUTC = (y, mo, d) => new Date(Date.UTC(y, mo, d));
  const eomD = d => dUTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0);
  const bomD = d => dUTC(d.getUTCFullYear(), d.getUTCMonth(), 1);
  const ymOf = d => isoOf(d).slice(0, 7);
  function feeBD(d) { const w = d.getUTCDay(); if (w === 0 || w === 6) return false; return !holidays(d.getUTCFullYear()).get(isoOf(d)); }
  function networkdays(a, b) { let n = 0; for (let d = a; d <= b; d = addDaysD(d, 1)) if (feeBD(d)) n++; return n; }
  function datedifM(a, b) { let n = (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + (b.getUTCMonth() - a.getUTCMonth()); if (b.getUTCDate() < a.getUTCDate()) n--; return n; }
  const FREQ_N = { Mensuel: 1, Trimestriel: 3, Semestriel: 6, Annuel: 12 };
  function lineFreq(l, m) {
    if (l.regle_facturation === 'par_distribution') { const n = +((m.params || {}).nb_distributions) || 0; return n === 12 ? 'Mensuel' : n === 4 ? 'Trimestriel' : n === 2 ? 'Semestriel' : 'Annuel'; }
    return FREQ_N[l.freq] ? l.freq : 'Annuel';
  }
  function invoicesOf(m) { return Array.isArray(FS.invoices) ? FS.invoices : JSON.parse(JSON.stringify(((m || model()).treso || {}).factures_constitution || [])); }
  function tresoOf(m) {
    const T = m.treso || {}; const R = refsOf(m);
    const launch = parseISO(T.lancement) || dUTC(2026, 7, 13);
    const fye1 = parseISO(T.fin_exercice_1) || eomD(addMonths(launch, 12));
    const H = Math.min(36, Math.max(1, Math.round(+T.horizon || 24)));
    const months = []; for (let i = 0; i < H; i++) { const d = addMonths(launch, i); months.push({ i, d, ym: ymOf(d), mon: d.getUTCMonth() + 1 }); }
    const nbdA = networkdays(launch, eomD(launch)), nbdM = networkdays(bomD(launch), eomD(launch));
    const prorata0 = nbdM ? nbdA / nbdM : 1;
    const aumIn = Array.isArray(T.aum_mensuel) ? T.aum_mensuel : []; let last = null;
    const aum = months.map((mo, i) => { const v = aumIn[i]; const set = v != null && v !== '' && !isNaN(+v); if (set) last = +v; return { v: last, set }; });
    const rows = (m.lines || []).map(l => {
      const annual = months.map((mo, i) => aum[i].v == null ? 0 : lineAnnual(l, aum[i].v, m, R));
      const prov = annual.map((a, i) => a / 12 * (i === 0 && l.prorata_lancement !== 'Non' ? prorata0 : 1));
      const fr = lineFreq(l, m), f = FREQ_N[fr];
      const inv = months.map((mo, i) => {
        if (l.regle_facturation === 'par_distribution') { const n = ((mo.mon >= 2 && (mo.mon - 1) % f === 0) ? 1 : 0) + (mo.mon === 12 ? 1 : 0); return n * annual[i] * f / 12; }
        if (mo.mon % f !== 0) return 0; let s = 0; for (let j = Math.max(0, i - (f - 1)); j <= i; j++) s += prov[j]; return s;
      });
      const cash = months.map((mo, i) => i > 0 ? inv[i - 1] : 0);
      return { l, freq: fr, annual, prov, inv, cash };
    });
    /* §8 - Frais de Constitution : exercices, anniversaire, factures réelles, ré-amortissement */
    const ex = []; let s0 = launch, e0 = fye1;
    for (let k = 0; k < 5; k++) { ex.push({ k: k + 1, start: s0, end: e0, nbm: datedifM(s0, e0) + 1 }); s0 = addDaysD(e0, 1); e0 = addDaysD(addMonths(s0, 12), -1); }
    const anniv = addMonths(launch, 12);
    const exA = ex.find(x => anniv >= x.start && anniv <= x.end) || null;
    const totalDu = R.constitution_montant, plancher = R.constitution_plancher;
    const reconnuEx1 = anniv <= fye1 ? totalDu : plancher, complement = anniv <= fye1 ? 0 : Math.max(0, totalDu - plancher);
    const facts = invoicesOf(m).map(x => { const di = parseISO(x.inv), dp = parseISO(x.pay); return Object.assign({}, x, { di, dp, eff: dp || (di ? addMonths(di, 1) : null) }); });
    const hasInv = months.map(mo => facts.some(x => x.di && ymOf(x.di) === mo.ym));
    const amtInv = months.map(mo => facts.reduce((s, x) => s + ((x.di && ymOf(x.di) === mo.ym) ? (+x.amount || 0) : 0), 0));
    const ex1 = ex[0]; const inEx = (mo, X) => !!X && mo.d >= X.start && mo.d <= X.end;
    const nb1 = months.filter((mo, i) => inEx(mo, ex1) && hasInv[i]).length, am1 = months.reduce((s, mo, i) => s + (inEx(mo, ex1) ? amtInv[i] : 0), 0);
    const nbA = exA ? months.filter((mo, i) => inEx(mo, exA) && hasInv[i]).length : 0, amA = exA ? months.reduce((s, mo, i) => s + (inEx(mo, exA) ? amtInv[i] : 0), 0) : 0;
    const provC = months.map((mo, i) => {
      if (inEx(mo, ex1)) return hasInv[i] ? amtInv[i] : ((ex1.nbm - nb1) <= 0 ? 0 : Math.max(0, reconnuEx1 - am1) / (ex1.nbm - nb1));
      if (exA && inEx(mo, exA)) return hasInv[i] ? amtInv[i] : ((exA.nbm - nbA) <= 0 ? 0 : Math.max(0, complement - amA) / (exA.nbm - nbA));
      return 0;
    });
    const cashC = months.map(mo => facts.reduce((s, x) => s + ((x.eff && ymOf(x.eff) === mo.ym) ? (+x.amount || 0) : 0), 0));
    const sumRows = k => months.map((_, i) => rows.reduce((s, r) => s + r[k][i], 0));
    const provT = sumRows('prov').map((v, i) => v + provC[i]), invT = sumRows('inv'), cashT = sumRows('cash').map((v, i) => v + cashC[i]);
    return { launch, fye1, H, months, nbdA, nbdM, prorata0, end: months[H - 1].d, aum, rows, ex, anniv, exA, totalDu, plancher, reconnuEx1, complement, facts, hasInv, amtInv, nb1, am1, nbA, amA, provC, cashC, provT, invT, cashT, billed: facts.reduce((s, x) => s + (+x.amount || 0), 0) };
  }

  /* ---------------- seuil de rentabilité et sensibilités ---------------- */
  const refBp = m => +((m.reference || {}).taux_bp) || 0;
  function breakevenOf(m) {
    const R = refsOf(m); const bp = refBp(m);
    const net = a => a * bp / 10000 - totalAt(m, a, R).total;
    const pts = []; for (let a = 0.25; a <= 100; a += 0.25) pts.push([a, net(a * M)]); for (let a = 101; a <= 2000; a += 1) pts.push([a, net(a * M)]);
    const roots = []; if (pts[0][1] >= 0) roots.push({ aum: 0, dir: 'up', below: true });
    for (let i = 1; i < pts.length; i++) { const [a0, m0] = pts[i - 1], [a1, m1] = pts[i]; if ((m0 < 0 && m1 >= 0) || (m0 >= 0 && m1 < 0)) roots.push({ aum: a0 + (0 - m0) * (a1 - a0) / (m1 - m0), dir: m1 >= 0 ? 'up' : 'down' }); }
    return roots;
  }
  const mRound = v => { const r = v < 10 ? Math.round(v * 10) / 10 : Math.round(v); return Number.isInteger(r) ? String(r) : r.toFixed(1).replace('.', DEC); };
  function beText(roots) { if (!roots.length) return 'non atteint sous 2 Md€'; return roots.map(r => r.below ? 'couvert dès le premier euro' : mRound(r.aum) + ' M€' + (r.dir === 'down' ? ' ↓' : '')).join(' · '); }
  function withShock(sn, sign, fn) {
    const saved = FS.ov; FS.ov = Object.assign({}, saved);
    (sn.ops || []).forEach(op => { const cur = getPath(model(), op.path); FS.ov[op.path] = (isRef(cur) ? 0 : (+cur || 0)) + sign * (+op.delta || 0); });
    let out; try { out = fn(model()); } finally { FS.ov = saved; }
    return out;
  }
  function feeAumRef(m) { if (FS.aum && +FS.aum > 0) return +FS.aum * M; const a = (m.treso || {}).aum_mensuel || []; for (let i = a.length - 1; i >= 0; i--) if (a[i] != null && +a[i] > 0) return +a[i]; const lad = ladderOf(m); return lad.length ? lad[0] : 10 * M; }

  /* ---------------- saisie : champs ---------------- */
  function inp(path, kind, o) {
    o = o || {}; const m = o.m || model(); const v = getPath(m, path);
    if (isRef(v)) return calc(o.R ? fmtKind(kind === 'auto' ? 'eur' : kind, +o.R[v.ref] || 0) : '…', refLabel(v.ref));
    const cls = 'fm-in' + (kind === 'txt' ? ' txt' : '') + (kind === 'sel' ? ' sel' : '') + (isOv(path) ? ' ov' : '') + (o.cls ? ' ' + o.cls : '') + (o.bad ? ' fm-bad' : '');
    const attrs = ' data-fv="' + esc(path) + '" data-kind="' + kind + '"' + (o.title ? ' title="' + esc(o.title) + '"' : '') + (o.w ? ' style="width:' + o.w + 'px"' : '');
    if (kind === 'sel') return '<select class="' + cls + '"' + attrs + '>' + (o.opts || []).map(x => { const ov = Array.isArray(x) ? x[0] : x, lab = Array.isArray(x) ? x[1] : x; return '<option value="' + esc(ov) + '"' + (String(v == null ? '' : v) === String(ov) ? ' selected' : '') + '>' + esc(lab) + '</option>'; }).join('') + '</select>';
    if (kind === 'txt') return '<input type="text" class="' + cls + '"' + attrs + ' value="' + esc(v == null ? '' : v) + '"' + (o.ph ? ' placeholder="' + esc(o.ph) + '"' : '') + '>';
    if (kind === 'date') return '<input type="date" class="' + cls + '"' + attrs + ' value="' + esc(v || '') + '">';
    if (kind === 'pct') return '<input type="number" step="any" class="' + cls + '"' + attrs + ' value="' + (v == null || v === '' ? '' : +(+v * 100).toFixed(6)) + '"' + (o.ph ? ' placeholder="' + esc(o.ph) + '"' : '') + '>';
    return '<input type="number" step="' + (kind === 'int' ? '1' : 'any') + '" class="' + cls + '"' + attrs + ' value="' + (v == null || v === '' ? '' : v) + '"' + (o.ph ? ' placeholder="' + esc(o.ph) + '"' : '') + '>';
  }
  const calc = (text, title) => '<span class="fm-calc"' + (title ? ' title="' + esc(title) + '"' : '') + '>' + text + '</span>';
  function fmtKind(kind, v) { return kind === 'pct' ? fRate(v) : kind === 'int' ? String(Math.round(v)) : fEur(v); }
  const NAT = ['Forfait', 'Taux', 'Palier'];
  const MODE_OPTS = [['bareme', 'Barème (Forfait / Taux / Palier)'], ['classe', 'Taux Classe sélectionnée'], ['kyc', 'Frais KYC (paliers de 100 M€)'], ['residuel', 'Résiduel vs Gestion (non paramétré : 0)']];
  const ON = ['Oui', 'Non'];
  function trancheCell(path, t, k, R, m) {
    const kindV = t.nature === 'Taux' ? 'pct' : 'eur';
    return '<span class="fm-tr"><b>T' + (k + 1) + '</b>' + inp(path + '.inf', 'eur', { m, R, w: 92, title: 'borne inférieure (€ d’Actif Net)' }) + '<i>→</i>' + inp(path + '.sup', 'eur', { m, R, w: 92, ph: '∞', title: 'borne supérieure (€ d’Actif Net) - vide : sans limite' }) + inp(path + '.nature', 'sel', { m, R, opts: NAT, w: 72 }) + inp(path + '.valeur', kindV, { m, R, w: 86, title: t.nature === 'Taux' ? 'taux annuel en % de la tranche' : 'montant annuel en €' }) + '</span>';
  }
  function tranchesShown(l, R) { const tr = l.tranches || []; let last = -1; tr.forEach((t, k) => { const used = isRef(t.valeur) || isRef(t.sup) || t.sup == null || val(t.sup, R, 0) !== 0 || val(t.valeur, R, 0) !== 0; if (used) last = k; }); return Math.max(1, Math.min(tr.length, last + 2)); }

  /* ---------------- vues ---------------- */
  function feeLegend() { return '<div class="fm-legend"><span class="lime">saisie - propre à ce fonds</span><span class="grey">calculé ou barème standard Fundcraft - piloté par un autre paramètre</span><span class="ov">modifié par rapport au classeur</span><span class="bad">incohérence à corriger</span></div>'; }
  function feeToolbar(m) {
    const n = nOv();
    return '<div class="savebar' + (n ? ' local' : ' readonly') + '"><span class="mode">' + (n ? n + (n === 1 ? ' saisie modifiée' : ' saisies modifiées') + ' par rapport au classeur' : 'Valeurs du classeur') + '</span><span class="hint2">Les saisies sont conservées dans ce navigateur et, après « Enregistrer le fichier », dans ce fichier HTML. Classeur de référence : ' + esc(FEES.source_file || '') + '.</span><span class="grow"></span>' +
      '<button type="button" data-fee-export-json>Exporter les saisies (JSON)</button><button type="button" data-fee-import-json>Importer des saisies (JSON)</button><input type="file" accept=".json,application/json" data-fee-import-file hidden>' +
      (n || FS.invoices ? '<button type="button" class="danger" data-fee-reset>Revenir au classeur</button>' : '') + '</div>';
  }
  function viewSaisie(m) {
    const R = refsOf(m); const lad = ladderOf(m); const bad = ladderBad(m);
    let h = '<div class="lead-p derived-lead"><span class="pill reading">Outil d’interface</span> Zone de saisie du modèle de frais, reprise de l’onglet 01_Input (et 02_Classes_de_Parts) du classeur Fundcraft : nom du fonds, échelle d’Actif Net, paramètres auxiliaires, classe de parts, valorisation, barèmes de commissions et tableau des vingt postes de frais à cinq tranches. Les vues Récap, Détail, Provision et Trésorerie se recalculent à chaque saisie.</div>';
    h += feeLegend() + feeToolbar(m);
    /* §1 + §2 */
    h += '<div class="card"><h2>1. Fonds testé · 2. Échelle d’Actif Net</h2><div class="sub">L’échelle est l’axe du Récap (04) : un palier par pas, du minimum au maximum, 150 paliers au plus.</div><div class="pform">' +
      '<label class="pfield"><span class="eyebrow">Nom du fonds</span>' + inp('nom_fonds', 'txt', { m, R }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">AUM minimum testé (€)</span>' + inp('echelle.min', 'eur', { m, R, bad }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">AUM maximum testé (€)</span>' + inp('echelle.max', 'eur', { m, R, bad }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Pas / incrément (€)</span>' + inp('echelle.pas', 'eur', { m, R, bad }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Paliers utilisés / disponibles</span>' + calc(lad.length + ' / ' + (+(m.echelle || {}).max_paliers || 150)) + (bad ? '<small class="flagbox">AUM minimum ≥ maximum ou pas nul : échelle incohérente.</small>' : '<small>' + lad.map(fM).slice(0, 6).join(' · ') + (lad.length > 6 ? ' · …' : '') + '</small>') + '</label></div></div>';
    /* §3 */
    h += '<div class="card"><h2>3. Paramètres auxiliaires</h2><div class="sub">Utilisés par certaines lignes du tableau de frais : Tenue du Passif, Frais KYC, forfait CAC par distribution, Frais de Constitution.</div><div class="pform">';
    Object.keys(m.params || {}).forEach(k => { const meta = (m.params_meta || {})[k] || {}; h += '<label class="pfield"><span class="eyebrow">' + esc(meta.label || k) + '</span>' + inp('params.' + k, meta.kind || 'eur', { m, R }) + (meta.note ? '<small>' + esc(meta.note) + '</small>' : '') + '</label>'; });
    h += '</div></div>';
    /* §3B */
    const classOpts = (m.classes || []).filter(c => String(c.nom || '').trim() !== '').map(c => String(c.nom));
    h += '<div class="card"><h2>3B. Classe de parts sélectionnée</h2><div class="sub">Le Récap et la Trésorerie calculent le détail pour une classe à la fois (lignes en mode « Taux Classe sélectionnée », forfait CAC par distribution). Pour lire une autre classe, changez la sélection : le classeur ne fait pas de mix pondéré multi-classes.</div><div class="pform">' +
      '<label class="pfield"><span class="eyebrow">Classe sélectionnée</span>' + inp('classe_selectionnee', 'sel', { m, R, opts: classOpts.length ? classOpts : [''] }) + (R.classe ? '' : '<small class="flagbox">classe introuvable dans 02_Classes_de_Parts</small>') + '</label>' +
      '<label class="pfield"><span class="eyebrow">Taux Distributeur Global (calculé)</span>' + calc(fRate(R.distrib_global_classe), 'repris de 02_Classes_de_Parts') + '</label>' +
      '<label class="pfield"><span class="eyebrow">Taux Sous-distributeurs (calculé)</span>' + calc(fRate(R.distrib_sous_classe), 'repris de 02_Classes_de_Parts') + '</label>' +
      '<label class="pfield"><span class="eyebrow">Cette classe distribue-t-elle ? (calculé)</span>' + calc(esc(R.classe_distribue)) + '</label></div></div>';
    /* §3C */
    const V = m.valo || {}; const pals = V.paliers || {}; const bornes = (lineById(m, 'valorisation') || { tranches: [] }).tranches.filter(t => t.nature === 'Palier');
    h += '<div class="card"><h2>3C. Paramètres de valorisation (CIC)</h2><div class="sub">Annexe tarifaire du contrat de délégation de valorisation Fundcraft × CIC du 24 juin 2026 : quatre paliers exclusifs d’Actif Net selon la fréquence de VL, majorés de ' + fPct(((V.majoration || [])[1] || [0, 0])[1], 0) + ' à partir de ' + (((V.majoration || [])[1] || [3])[0]) + ' catégories de parts et de ' + fPct(((V.majoration || [])[0] || [0, 0])[1], 0) + ' à partir de ' + (((V.majoration || [])[0] || [5])[0]) + ', plus ' + fEur((V.supplement || {}).montant) + ' € par tranche de ' + ((V.supplement || {}).par_tranche || 10) + ' lignes de portefeuille au-delà de ' + ((V.supplement || {}).lignes_incluses || 10) + '.</div><div class="pform">' +
      '<label class="pfield"><span class="eyebrow">Fréquence de la VL</span>' + inp('valo.frequence_vl', 'sel', { m, R, opts: Object.keys(pals).length ? Object.keys(pals) : ['Mensuelle', 'Bi-mensuelle'] }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Nombre de lignes du portefeuille</span>' + inp('valo.nb_lignes', 'int', { m, R }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Catégories de parts (déduit de 02)</span>' + calc(String(R.nb_categories)) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Majoration applicable (calculée)</span>' + calc(fPct(R.majoration, 0)) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Supplément lignes (calculé)</span>' + calc(fEur(R.valo_supplement) + ' €') + '</label></div>';
    h += '<h3 class="sub">Barème standard CIC - verrouillé dans le classeur, modifiable ici en connaissance de cause</h3><div class="tbl-wrap"><table class="tbl fm"><thead><tr><th>Tranche d’Actif Net</th>' + Object.keys(pals).map(f => '<th class="num">VL ' + esc(f.toLowerCase()) + ' (€/an)</th>').join('') + '<th class="num">Montant retenu (€/an)</th></tr></thead><tbody>';
    (pals.Mensuelle || []).forEach((x, k) => { const b = bornes[k] || {}; h += '<tr><td>' + (k === 0 ? '0 – ' + fM(val(b.sup, R, FINF)) : val(b.sup, R, FINF) >= FINF ? '> ' + fM(val(b.inf, R, 0)) : fM(val(b.inf, R, 0)) + ' – ' + fM(val(b.sup, R, FINF))) + ' <small class="muted">bornes modifiables dans la ligne « Valorisation (CIC) » du tableau 4</small></td>' + Object.keys(pals).map(f => '<td class="num">' + inp('valo.paliers.' + f + '.' + k, 'eur', { m, R, cls: 'grey' }) + '</td>').join('') + '<td class="num">' + calc(fEur(R['valo_p' + (k + 1)]), 'palier × (1 + majoration)') + '</td></tr>'; });
    h += '</tbody></table></div><div class="pform" style="margin-top:8px">' +
      '<label class="pfield"><span class="eyebrow">Lignes incluses</span>' + inp('valo.supplement.lignes_incluses', 'int', { m, R, cls: 'grey' }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Par tranche de (lignes)</span>' + inp('valo.supplement.par_tranche', 'int', { m, R, cls: 'grey' }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Supplément par tranche (€)</span>' + inp('valo.supplement.montant', 'eur', { m, R, cls: 'grey' }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Majoration dès (catégories) → taux</span><span>' + inp('valo.majoration.1.0', 'int', { m, R, cls: 'grey', w: 60 }) + ' ' + inp('valo.majoration.1.1', 'pct', { m, R, cls: 'grey', w: 70 }) + '</span></label>' +
      '<label class="pfield"><span class="eyebrow">Majoration dès (catégories) → taux</span><span>' + inp('valo.majoration.0.0', 'int', { m, R, cls: 'grey', w: 60 }) + ' ' + inp('valo.majoration.0.1', 'pct', { m, R, cls: 'grey', w: 70 }) + '</span></label></div></div>';
    /* §3D */
    const B = m.baremes || {};
    h += '<div class="card"><h2>3D. Barèmes de commissions - Commission de Gestion et Commission de Conseil</h2><div class="sub">Un seul tableau pour les deux commissions calculées par tranches progressives d’Actif Net (Prospectus Art. 9.2(a) et (b)) : les deux seuils et les six taux se répercutent dans les lignes 1 et 2 du tableau des frais (cellules grises).</div><div class="pform">' +
      '<label class="pfield"><span class="eyebrow">Seuil tranche 1 → 2 (€ d’Actif Net)</span>' + inp('baremes.seuil_t1', 'eur', { m, R }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Seuil tranche 2 → 3 (€ d’Actif Net)</span>' + inp('baremes.seuil_t2', 'eur', { m, R }) + '</label></div>' +
      '<div class="tbl-wrap"><table class="tbl fm"><thead><tr><th>Tranche (Actif Net du fonds)</th><th class="num">Commission de Gestion (taux)</th><th class="num">Commission de Conseil (taux)</th></tr></thead><tbody>' +
      [['Tranche 1 : 0 € → ' + fEur(B.seuil_t1) + ' €', 0], ['Tranche 2 : ' + fEur(B.seuil_t1) + ' € → ' + fEur(B.seuil_t2) + ' €', 1], ['Tranche 3 : au-delà de ' + fEur(B.seuil_t2) + ' €', 2]].map(([lab, k]) => '<tr><td>' + esc(lab) + '</td><td class="num">' + inp('baremes.gestion.' + k, 'pct', { m, R }) + '</td><td class="num">' + inp('baremes.conseil.' + k, 'pct', { m, R }) + '</td></tr>').join('') +
      '</tbody></table></div></div>';
    /* 02 classes */
    h += '<div class="card"><h2>02. Classes de parts</h2><div class="sub">Chaque classe porte ses deux étages de Commission de Distribution (Distributeur Global, Sous-distributeurs) et son attribut capitalisante / distribuante. Une ligne au nom vide est ignorée ; le nombre de lignes renseignées donne le nombre de catégories de parts (§3C).</div><div class="tbl-wrap"><table class="tbl fm"><thead><tr><th>#</th><th>Nom de la classe</th><th class="num">Taux Distributeur Global</th><th class="num">Taux Sous-distributeurs</th><th>Distribuante ?</th><th>Commentaire / source</th></tr></thead><tbody>';
    (m.classes || []).forEach((c, i) => { const empty = String(c.nom || '').trim() === ''; h += '<tr' + (empty ? ' class="off"' : '') + '><td>' + (i + 1) + '</td><td>' + inp('classes.' + i + '.nom', 'txt', { m, R, w: 120, ph: 'libre' }) + '</td><td class="num">' + inp('classes.' + i + '.distrib_global', 'pct', { m, R }) + '</td><td class="num">' + inp('classes.' + i + '.sous_distrib', 'pct', { m, R }) + '</td><td>' + inp('classes.' + i + '.distribuante', 'sel', { m, R, opts: ON }) + '</td><td>' + inp('classes.' + i + '.commentaire', 'txt', { m, R, w: 360 }) + '</td></tr>'; });
    h += '</tbody></table></div></div>';
    /* §4 lignes */
    h += '<div class="card"><h2>4. Frais et provisions du Fonds - vingt postes</h2><div class="sub">Une ligne par frais ou provision, jusqu’à cinq tranches Forfait / Taux / Palier. <b>Forfait</b> : montant fixe dès que l’Actif Net atteint la borne inférieure (plusieurs forfaits se cumulent). <b>Taux</b> : portion de l’Actif Net dans la tranche × taux, par tranches progressives. <b>Palier</b> : montant fixe exclusif de la tranche atteinte. Le résultat est borné par le plancher et le plafond annuels. Désactivez une ligne (Actif = Non) plutôt que de l’effacer ; les lignes 18 à 20 sont libres. Une tranche vide apparaît après la dernière utilisée.</div>';
    h += '<div class="tbl-wrap"><table class="tbl fm lines"><thead><tr><th>#</th><th>Actif ?</th><th>Nom du poste</th><th>Bénéficiaire</th><th>Récurrent ?</th><th>Mode de calcul</th><th class="num">Plancher annuel (€)</th><th class="num">Plafond annuel (€)</th><th>TVA</th><th>Fact.</th><th>Prorata lancement</th></tr></thead><tbody>';
    (m.lines || []).forEach((l, i) => {
      const P = 'lines.' + i; const off = l.actif !== 'Oui';
      h += '<tr class="main' + (off ? ' off' : '') + '"><td>' + l.n + '</td><td>' + inp(P + '.actif', 'sel', { m, R, opts: ON }) + '</td><td>' + inp(P + '.nom', 'txt', { m, R, w: 230, ph: 'poste libre' }) + '</td><td>' + inp(P + '.beneficiaire', 'txt', { m, R, w: 170 }) + '</td><td>' + inp(P + '.recurrent', 'sel', { m, R, opts: ON }) + '</td><td>' + inp(P + '.mode', 'sel', { m, R, opts: MODE_OPTS }) + '</td><td class="num">' + inp(P + '.plancher', 'eur', { m, R, ph: '0' }) + '</td><td class="num">' + inp(P + '.plafond', 'eur', { m, R, ph: '∞' }) + '</td><td>' + inp(P + '.tva', 'sel', { m, R, opts: [['0', 'non'], ['0.2', '20 %']], w: 64 }) + '</td><td>' + (l.regle_facturation === 'par_distribution' ? calc(esc(lineFreq(l, m)), 'fréquence dérivée du nombre de distributions par an (§3)') : inp(P + '.freq', 'sel', { m, R, opts: Object.keys(FREQ_N) })) + '</td><td>' + inp(P + '.prorata_lancement', 'sel', { m, R, opts: ON, title: 'Oui : le premier mois est provisionné au prorata des Jours Ouvrés ; Non : mois entamé dû en totalité (Finologee)' }) + '</td></tr>';
      let sub = '';
      if (l.mode === 'bareme') { const n = tranchesShown(l, R); for (let k = 0; k < n; k++) sub += trancheCell(P + '.tranches.' + k, (l.tranches || [])[k] || {}, k, R, m); }
      else if (l.mode === 'classe') sub = '<span class="fm-tr"><b>Taux de la classe sélectionnée</b>' + inp(P + '.classe_taux', 'sel', { m, R, opts: [['global', 'Distributeur Global'], ['sous', 'Sous-distributeurs']] }) + calc(fRate(l.classe_taux === 'sous' ? R.distrib_sous_classe : R.distrib_global_classe) + ' × Actif Net', 'classe ' + esc(m.classe_selectionnee || '')) + '</span>';
      else if (l.mode === 'kyc') sub = '<span class="fm-tr"><b>Frais KYC</b>' + calc('MAX(0 ; ' + (+(m.params || {}).nb_investisseurs || 0) + ' investisseurs − ' + (+(m.params || {}).nb_kyc_gratuites || 0) + ' × (ENT(Actif Net ÷ ' + fM((m.params || {}).taille_tranche_kyc) + ') + 1)) × ' + fEur((m.params || {}).cout_kyc_supp) + ' €', 'paramètres du §3') + '</span>';
      else sub = '<span class="fm-tr"><b>Résiduel vs Gestion</b>' + calc('0 - mode non paramétré dans cette version du classeur') + '</span>';
      const src = l.source ? '<details class="fm-src"><summary>Commentaire / source</summary>' + inp(P + '.source', 'txt', { m, R, w: 720 }) + '<div class="fm-srctext">' + esc(l.source) + '</div></details>' : '<details class="fm-src"><summary>Commentaire / source</summary>' + inp(P + '.source', 'txt', { m, R, w: 720, ph: 'source, convention, décision…' }) + '</details>';
      h += '<tr class="tr-row' + (off ? ' off' : '') + '"><td></td><td colspan="10">' + sub + src + '</td></tr>';
    });
    h += '</tbody></table></div></div>';
    return h + feeDocs(m);
  }
  function feeDocs(m) {
    const docs = m.documentation || []; if (!docs.length) return '';
    let h = '<div class="card"><h2>Documentation du classeur (06)</h2><div class="sub">Méthode, hypothèses et sources, reprises de l’onglet 06_Documentation - sans effet sur le calcul.</div>';
    docs.forEach(d => { h += '<h3 class="sub">' + esc(d.title) + '</h3><div class="kv small">' + (d.items || []).map(it => '<div class="k">' + esc(it.k || '') + '</div><div class="v">' + esc(it.v || '') + '</div>').join('') + '</div>'; });
    return h + '</div>';
  }
  function kpi(k, v, s) { return '<div class="kpi"><div class="k">' + k + '</div><div class="v">' + v + '</div>' + (s ? '<div class="s">' + s + '</div>' : '') + '</div>'; }
  function aumRefKpi(m) { const a = feeAumRef(m); return '<div class="kpi"><div class="k">Actif Net de référence</div><div class="v"><input type="number" min="0.1" step="0.05" data-fee-aum value="' + (Math.round(a / M * 1000) / 1000) + '" style="width:96px"> <span class="u">M€</span></div><div class="s">' + (FS.aum ? 'saisi - videz pour revenir au dernier Actif Net mensuel' : 'dernier Actif Net mensuel saisi (05 §2)') + '</div></div>'; }
  function viewRecap(m) {
    const r = recapOf(m); const a0 = feeAumRef(m); const t0 = totalAt(m, a0, r.R); const be = breakevenOf(m);
    let h = '<div class="lead-p derived-lead"><span class="pill reading">Outil d’interface</span> Zéro saisie : une ligne par poste, une colonne par palier d’Actif Net de l’échelle (01 §2), comme l’onglet 04_Recap du classeur - pour lire comment chaque poste évolue avec la taille du Fonds. Montants annuels HT ; la classe testée est <b>' + esc(m.classe_selectionnee || '-') + '</b>. Les Frais de Constitution sont traités en Trésorerie (§8) et n’apparaissent ici que si leur ligne est activée.</div>';
    h += '<div class="card"><div class="kpis">' + aumRefKpi(m) +
      kpi('Frais directs annuels à cet Actif Net', fEur(t0.total) + ' <span class="u">EUR</span>', fPct(a0 > 0 ? t0.total / a0 : null) + ' de l’Actif Net · ' + fBp(a0 > 0 ? t0.total / a0 : null)) +
      kpi('TVA sur les lignes assujetties', fEur(t0.vat) + ' <span class="u">EUR</span>', 'total TTC ' + fEur(t0.total + t0.vat) + ' € · ' + fPct(a0 > 0 ? (t0.total + t0.vat) / a0 : null)) +
      '<div class="kpi"><div class="k">Actif Net au taux de référence</div><div class="v">' + esc(beText(be)) + '</div><div class="s">taux de référence <input type="number" step="1" data-fv="reference.taux_bp" data-kind="int" class="fm-in' + (isOv('reference.taux_bp') ? ' ov' : '') + '" style="width:64px" value="' + refBp(m) + '"> bp - voir « Seuil de rentabilité »</div></div>' +
      '</div>' + (ladderBad(m) ? '<div class="lead-p flagbox">Échelle d’Actif Net incohérente (01 §2) : corrigez le minimum, le maximum ou le pas dans la Saisie.</div>' : '') + '</div>';
    h += '<div class="card"><h2>Récap par palier d’Actif Net (04)</h2><div class="sub">Fonds testé : <b>' + esc(m.nom_fonds || '') + '</b> · classe testée : <b>' + esc(m.classe_selectionnee || '-') + '</b> · ' + r.ladder.length + ' paliers de ' + fM((m.echelle || {}).min) + ' à ' + fM(r.ladder[r.ladder.length - 1]) + ' par pas de ' + fM((m.echelle || {}).pas) + '. € par an, et en petit le poids en % de l’Actif Net. <button type="button" class="mini" data-fee-export-csv="recap">Exporter en CSV</button></div>';
    h += '<div class="tbl-wrap"><table class="tbl fees fm wide"><thead><tr><th>Poste de frais</th><th>Bénéficiaire</th><th>Actif ?</th>' + r.ladder.map(a => '<th class="num">' + fM(a) + '</th>').join('') + '</tr></thead><tbody>';
    r.rows.forEach(x => { const l = x.l; const off = l.actif !== 'Oui'; h += '<tr' + (off ? ' class="off"' : '') + '><td>' + (l.nom ? esc(l.nom) : muted('(inactif)')) + (+l.tva ? ' <small class="muted">+TVA</small>' : '') + '</td><td><small>' + esc(l.beneficiaire || '') + '</small></td><td>' + esc(l.actif) + '</td>' + x.vals.map((v, i) => '<td class="num">' + (off ? muted('0') : fEur(v) + '<br><small>' + fPct(r.ladder[i] > 0 ? v / r.ladder[i] : null, 3) + '</small>') + '</td>').join('') + '</tr>'; });
    h += '<tr class="total"><td>TOTAL (€)</td><td></td><td></td>' + r.tot.map(v => '<td class="num">' + fEur(v) + '</td>').join('') + '</tr>';
    h += '<tr class="total"><td>TOTAL (% AUM)</td><td></td><td></td>' + r.pct.map(v => '<td class="num">' + fPct(v, 3) + '</td>').join('') + '</tr>';
    h += '<tr class="indirect"><td>TVA sur les lignes assujetties (€)</td><td></td><td></td>' + r.vat.map(v => '<td class="num">' + fEur(v) + '</td>').join('') + '</tr>';
    h += '<tr class="indirect"><td>TOTAL TTC (€)</td><td></td><td></td>' + r.tot.map((v, i) => '<td class="num">' + fEur(v + r.vat[i]) + '</td>').join('') + '</tr>';
    h += '</tbody></table></div></div>';
    h += '<div class="card"><h2>Frais directs en fonction de l’Actif Net</h2><div class="sub">Total annuel des postes actifs (€, HT, à gauche) et taux all-in (bp, à droite), sur l’échelle testée. Le plancher de la Commission de Gestion et les forfaits des prestataires rendent la courbe plate aux faibles encours.</div><div class="swim-wrap">' + feeCurve(m, 'fund') + '</div></div>';
    h += '<div class="card"><h2>Sensibilités à ' + fM(a0) + '</h2><div class="sub">Chaque choc est appliqué seul (+ / −) sur le modèle tel que saisi ; le reste reste inchangé.</div><div class="tbl-wrap"><table class="tbl"><thead><tr><th>Choc</th><th class="num">Δ frais directs (€ / an)</th><th class="num">Δ taux all-in</th><th>Actif Net au taux de référence après choc (+ / −)</th></tr></thead><tbody>';
    (m.sensitivities || []).forEach(sn => {
      const up = withShock(sn, 1, mm => ({ t: totalAt(mm, a0).total, b: breakevenOf(mm) })), dn = withShock(sn, -1, mm => ({ t: totalAt(mm, a0).total, b: breakevenOf(mm) }));
      const sg = v => (v > 0 ? '+' : '') + fEur(v);
      h += '<tr><td>' + esc(sn.label) + '</td><td class="num">' + sg(up.t - t0.total) + ' / ' + sg(dn.t - t0.total) + '</td><td class="num">' + (a0 > 0 ? ((up.t - t0.total) / a0 * 10000 > 0 ? '+' : '') + fBp((up.t - t0.total) / a0) + ' / ' + fBp((dn.t - t0.total) / a0) : '-') + '</td><td>' + esc(beText(up.b)) + ' / ' + esc(beText(dn.b)) + '</td></tr>';
    });
    h += '</tbody></table></div></div>';
    return h;
  }
  function viewDetail(m) {
    const d = detailOf(m); const lad = d.ladder;
    let h = '<div class="lead-p derived-lead"><span class="pill reading">Outil d’interface</span> Zéro saisie : décompose pas à pas, sur toute l’échelle d’Actif Net, les deux postes dont le calcul n’est pas un simple taux - la Commission Dépositaire (taux, plancher et plafond annuels) et la Valorisation CIC (palier exclusif par tranche d’actif, fréquence de VL, majoration par nombre de catégories de parts, supplément de lignes) - comme l’onglet 03 du classeur. Les lignes « Écart vs Récap » doivent rester à 0.</div>';
    if (!d.D && !d.V) return h + '<div class="card"><div class="lead-p">Les lignes « Commission Dépositaire (BFCM) » et « Valorisation (CIC) » sont absentes du modèle.</div></div>';
    h += '<div class="card"><h2>Paramètres actifs - repris de la Saisie</h2><div class="kv small">' +
      (d.D ? '<div class="k">Dépositaire - taux appliqué</div><div class="v">' + fRate(d.D.taux) + '</div><div class="k">Dépositaire - plancher annuel</div><div class="v">' + fEur(d.D.plancher) + ' €</div><div class="k">Dépositaire - plafond annuel</div><div class="v">' + (d.D.plafond >= FINF ? 'sans plafond' : fEur(d.D.plafond) + ' €') + '</div>' : '') +
      (d.V ? '<div class="k">Valorisation - fréquence de VL</div><div class="v">' + esc((m.valo || {}).frequence_vl || '') + '</div><div class="k">Valorisation - lignes du portefeuille</div><div class="v">' + ((m.valo || {}).nb_lignes || 0) + '</div><div class="k">Valorisation - catégories de parts</div><div class="v">' + d.R.nb_categories + '</div><div class="k">Valorisation - majoration applicable</div><div class="v">' + fPct(d.R.majoration, 0) + '</div>' : '') + '</div></div>';
    const head = '<thead><tr><th>Étape de calcul</th>' + lad.map(a => '<th class="num">' + fM(a) + '</th>').join('') + '</tr></thead>';
    if (d.D) {
      h += '<div class="card"><h2>Commission Dépositaire (BFCM)</h2><div class="sub">Montant théorique = taux × Actif Net ; retenu = MIN(plafond ; MAX(plancher ; théorique)).</div><div class="tbl-wrap"><table class="tbl fees fm wide">' + head + '<tbody>';
      h += '<tr><td>Taux appliqué</td>' + lad.map(() => '<td class="num">' + fRate(d.D.taux) + '</td>').join('') + '</tr>';
      h += '<tr><td>Montant théorique = taux × AUM (€)</td>' + d.D.rows.map(x => '<td class="num">' + fEur(x.th) + '</td>').join('') + '</tr>';
      h += '<tr><td>Plancher annuel (€)</td>' + lad.map(() => '<td class="num">' + fEur(d.D.plancher) + '</td>').join('') + '</tr>';
      h += '<tr><td>Plafond annuel (€)</td>' + lad.map(() => '<td class="num">' + (d.D.plafond >= FINF ? '∞' : fEur(d.D.plafond)) + '</td>').join('') + '</tr>';
      h += '<tr><td>Mécanisme retenu</td>' + d.D.rows.map(x => '<td class="num"><small>' + x.mec + '</small></td>').join('') + '</tr>';
      h += '<tr class="total"><td>→ Commission Dépositaire retenue (€)</td>' + d.D.rows.map(x => '<td class="num">' + fEur(x.ret) + '</td>').join('') + '</tr>';
      h += '<tr class="indirect"><td>Écart vs Récap (doit être 0)</td>' + d.D.rows.map(x => '<td class="num' + (x.ecart ? ' neg' : '') + '">' + fEur(x.ecart) + '</td>').join('') + '</tr></tbody></table></div></div>';
    }
    if (d.V) {
      h += '<div class="card"><h2>Valorisation (CIC)</h2><div class="sub">Palier exclusif de la tranche d’actif atteinte, selon la fréquence de VL et majoré du nombre de catégories de parts, plus le supplément de lignes de portefeuille.</div><div class="tbl-wrap"><table class="tbl fees fm wide">' + head + '<tbody>';
      h += '<tr><td>Fréquence de VL retenue</td>' + lad.map(() => '<td class="num">' + esc((m.valo || {}).frequence_vl || '') + '</td>').join('') + '</tr>';
      h += '<tr><td>Tranche d’actif atteinte</td>' + d.V.rows.map(x => '<td class="num"><small>' + esc(x.lab) + '</small></td>').join('') + '</tr>';
      h += '<tr><td>Montant de la tranche, majoration incluse (€)</td>' + d.V.rows.map(x => '<td class="num">' + fEur(x.mt) + '</td>').join('') + '</tr>';
      h += '<tr><td>Majoration appliquée</td>' + lad.map(() => '<td class="num">' + fPct(d.R.majoration, 0) + '</td>').join('') + '</tr>';
      h += '<tr><td>Nombre de lignes du portefeuille</td>' + lad.map(() => '<td class="num">' + ((m.valo || {}).nb_lignes || 0) + '</td>').join('') + '</tr>';
      h += '<tr><td>Supplément lignes (€)</td>' + d.V.rows.map(x => '<td class="num">' + fEur(x.sp) + '</td>').join('') + '</tr>';
      h += '<tr class="total"><td>→ Valorisation (CIC) retenue (€)</td>' + d.V.rows.map(x => '<td class="num">' + fEur(x.ret) + '</td>').join('') + '</tr>';
      h += '<tr class="indirect"><td>Écart vs Récap (doit être 0)</td>' + d.V.rows.map(x => '<td class="num' + (x.ecart ? ' neg' : '') + '">' + fEur(x.ecart) + '</td>').join('') + '</tr></tbody></table></div></div>';
    }
    return h;
  }
  function viewProvision(m) {
    const R = refsOf(m); const T = tresoOf(m); const a0 = feeAumRef(m);
    const ymLaunch = ymOf(T.launch); const today = ymOf(new Date()); const inH = ym => T.months.some(x => x.ym === ym);
    const ym = (FS.pmonth && /^\d{4}-\d{2}$/.test(FS.pmonth)) ? FS.pmonth : (inH(today) ? today : ymLaunch);
    const mo = T.months.find(x => x.ym === ym) || null; const prorata = ym === ymLaunch ? T.prorata0 : 1;
    const rows = (m.lines || []).filter(l => l.actif === 'Oui').map(l => { const ann = lineAnnual(l, a0, m, R); const p = ann / 12 * (l.prorata_lancement !== 'Non' ? prorata : 1); const v = p * (+l.tva || 0); return { l, ann, p, v }; });
    const provC = mo ? T.provC[mo.i] : 0;
    const sumP = rows.reduce((s, r) => s + r.p, 0), sumV = rows.reduce((s, r) => s + r.v, 0), sumA = rows.reduce((s, r) => s + r.ann, 0);
    let h = '<div class="lead-p derived-lead"><span class="pill reading">Outil d’interface</span> Provision de frais à comptabiliser à la Valeur Liquidative d’un mois donné, à l’Actif Net constaté : pour chaque poste actif, charge du mois = barème à cet Actif Net ÷ 12, au prorata des Jours Ouvrés si c’est le mois de lancement (sauf Finologee, mois entamé dû en totalité). Les Frais de Constitution suivent le calendrier de la Trésorerie (§8) : montant réel facturé ce mois-ci, sinon solde non facturé réparti sur les mois restants du 1er exercice.</div>';
    h += '<div class="card"><div class="kpis">' + aumRefKpi(m) +
      '<div class="kpi"><div class="k">Mois de la VL</div><div class="v"><input type="month" data-fee-pmonth value="' + esc(ym) + '"></div><div class="s">' + (ym === ymLaunch ? 'mois de lancement : prorata ' + fPct(T.prorata0, 1) + ' (' + T.nbdA + ' / ' + T.nbdM + ' Jours Ouvrés)' : mo ? 'mois ' + (mo.i + 1) + ' de l’horizon' : 'hors de l’horizon du prévisionnel : Frais de Constitution non calculés') + '</div></div>' +
      kpi('Provision du mois - postes récurrents (HT)', fEur(sumP) + ' <span class="u">EUR</span>', '+ TVA ' + fEur(sumV) + ' · TTC ' + fEur(sumP + sumV)) +
      kpi('Frais de Constitution du mois', fEur(provC) + ' <span class="u">EUR</span>', mo ? (T.hasInv[mo.i] ? 'facture réelle datée ce mois-ci' : 'solde réparti sur les mois restants du 1er exercice') : '-') +
      kpi('Total à provisionner (HT)', fEur(sumP + provC) + ' <span class="u">EUR</span>', (a0 > 0 ? fBp((sumP + provC) / a0) + ' du mois · ' : '') + fEur(sumP + provC + sumV) + ' € TTC') + '</div></div>';
    h += '<div class="card"><h2>Provision du mois, ligne par ligne</h2><div class="sub">À ' + fM(a0) + ' d’Actif Net, ' + esc(ymF(ym)) + '. Montants HT ; la colonne TVA n’est servie que pour les lignes assujetties.</div><div class="tbl-wrap"><table class="tbl fees fm"><thead><tr><th>Poste</th><th>Bénéficiaire</th><th class="num">Charge annuelle (€)</th><th class="num">Provision du mois (€)</th><th class="num">TVA du mois (€)</th><th class="num">Mois TTC (€)</th></tr></thead><tbody>';
    rows.forEach(r => { h += '<tr><td>' + esc(r.l.nom) + '</td><td><small>' + esc(r.l.beneficiaire || '') + '</small></td><td class="num">' + fEur(r.ann) + '</td><td class="num">' + fEur(r.p) + '</td><td class="num">' + (+r.l.tva ? fEur(r.v) : muted('-')) + '</td><td class="num">' + fEur(r.p + r.v) + '</td></tr>'; });
    h += '<tr class="total"><td>Total récurrent</td><td></td><td class="num">' + fEur(sumA) + '</td><td class="num">' + fEur(sumP) + '</td><td class="num">' + fEur(sumV) + '</td><td class="num">' + fEur(sumP + sumV) + '</td></tr>';
    h += '<tr class="indirect"><td>Frais de Constitution (05 §8)</td><td><small>divers</small></td><td class="num">' + fEur(T.reconnuEx1) + ' <small>sur le 1er exercice</small></td><td class="num">' + fEur(provC) + '</td><td class="num">' + muted('-') + '</td><td class="num">' + fEur(provC) + '</td></tr>';
    h += '<tr class="total"><td>Total à provisionner</td><td></td><td class="num"></td><td class="num">' + fEur(sumP + provC) + '</td><td class="num">' + fEur(sumV) + '</td><td class="num">' + fEur(sumP + provC + sumV) + '</td></tr></tbody></table></div></div>';
    const lad = ladderOf(m);
    h += '<div class="card"><h2>Provision mensuelle par palier d’Actif Net</h2><div class="sub">Lecture de sensibilité : ce que coûte un mois plein (hors prorata de lancement et hors Frais de Constitution) selon l’Actif Net constaté à la VL, sur l’échelle de la Saisie.</div><div class="tbl-wrap"><table class="tbl fees fm wide"><thead><tr><th>Poste</th>' + lad.map(a => '<th class="num">' + fM(a) + '</th>').join('') + '</tr></thead><tbody>';
    const cols = lad.map(a => { let p = 0, v = 0; const per = {}; (m.lines || []).forEach(l => { if (l.actif !== 'Oui') return; const x = lineAnnual(l, a, m, R) / 12; per[l.id] = x; p += x; v += x * (+l.tva || 0); }); return { a, p, v, per }; });
    (m.lines || []).filter(l => l.actif === 'Oui').forEach(l => { h += '<tr><td>' + esc(l.nom) + '</td>' + cols.map(c => '<td class="num">' + fEur(c.per[l.id]) + '</td>').join('') + '</tr>'; });
    h += '<tr class="total"><td>Provision du mois (HT)</td>' + cols.map(c => '<td class="num">' + fEur(c.p) + '</td>').join('') + '</tr>';
    h += '<tr class="total"><td>Provision du mois (TTC)</td>' + cols.map(c => '<td class="num">' + fEur(c.p + c.v) + '</td>').join('') + '</tr>';
    h += '<tr class="total"><td>Taux annualisé (HT)</td>' + cols.map(c => '<td class="num">' + fPct(c.a > 0 ? c.p * 12 / c.a : null, 3) + '</td>').join('') + '</tr></tbody></table></div></div>';
    return h;
  }
  function viewTreso(m) {
    const T = tresoOf(m); const sum = a => a.reduce((x, y) => x + y, 0); const n12 = Math.min(12, T.H);
    const head = '<thead><tr><th>Poste</th><th>Fact.</th>' + T.months.map((mo, i) => '<th class="num">' + esc(ymF(mo.ym)) + '<br><small>' + (T.aum[i].v == null ? '-' : fM(T.aum[i].v)) + '</small></th>').join('') + '</tr></thead>';
    const act = T.rows.filter(r => r.l.actif === 'Oui');
    const block = (title, sub, key, totals, extra) => {
      let t = '<div class="card"><h2>' + title + '</h2><div class="sub">' + sub + '</div><div class="tbl-wrap"><table class="tbl fees fm fc">' + head + '<tbody>';
      act.forEach(r => { t += '<tr><td>' + esc(r.l.nom) + '</td><td><small>' + esc(r.freq) + '</small></td>' + r[key].map(v => '<td class="num">' + fEurZ(v) + '</td>').join('') + '</tr>'; });
      if (extra) t += '<tr class="indirect"><td>' + extra.label + '</td><td></td>' + extra.vals.map(v => '<td class="num">' + fEurZ(v) + '</td>').join('') + '</tr>';
      t += '<tr class="total"><td>' + (extra ? 'TOTAL (€)' : 'TOTAL (€)') + '</td><td></td>' + totals.map(v => '<td class="num">' + fEur(v) + '</td>').join('') + '</tr>';
      return t + '</tbody></table></div></div>';
    };
    let h = '<div class="lead-p derived-lead"><span class="pill reading">Outil d’interface</span> Prévisionnel mensuel chaîné comme l’onglet 05_Tresorerie du classeur : l’Actif Net se saisit mois par mois ; chaque poste rejoue son barème à l’Actif Net du mois (§4), en <b>provisionne</b> un douzième - au prorata des Jours Ouvrés sur le mois de lancement, sauf Finologee (§5) -, <b>facture</b> à chaque échéance la somme des provisions accumulées depuis la précédente (§6) et <b>décaisse</b> le mois suivant la facture (§7). Les Frais de Constitution sont provisionnés sur le 1er exercice et ré-amortis au fil des factures réelles (§8). Montants HT.</div>';
    h += feeToolbar(m);
    h += '<div class="card"><h2>1. Paramètres du prévisionnel</h2><div class="pform">' +
      '<label class="pfield"><span class="eyebrow">Date de lancement du fonds</span>' + inp('treso.lancement', 'date', { m }) + '<small>premier closing : les charges courent à partir de ce jour</small></label>' +
      '<label class="pfield"><span class="eyebrow">Fin du 1er exercice comptable</span>' + inp('treso.fin_exercice_1', 'date', { m }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Nombre de mois du prévisionnel (36 au plus)</span>' + inp('treso.horizon', 'int', { m }) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Fin du prévisionnel (calculé)</span>' + calc(fDate(isoOf(T.end))) + '</label>' +
      '<label class="pfield"><span class="eyebrow">Prorata du mois de lancement (calculé)</span>' + calc(fPct(T.prorata0, 2), 'jours ouvrés actifs ÷ jours ouvrés du mois civil, jours fériés français exclus') + '<small>' + T.nbdA + ' Jours Ouvrés actifs sur ' + T.nbdM + ' en ' + esc(ymF(ymOf(T.launch))) + '</small></label></div></div>';
    h += '<div class="card"><h2>2. Actif Net mensuel du Fonds (€)</h2><div class="sub">' + esc((m.treso || {}).aum_note || '') + '</div><div class="tbl-wrap"><table class="tbl fm fc"><thead><tr><th>Mois</th>' + T.months.map(mo => '<th class="num">' + esc(ymF(mo.ym)) + '</th>').join('') + '</tr></thead><tbody><tr><td>Actif Net</td>' + T.months.map((mo, i) => '<td class="num">' + inp('treso.aum_mensuel.' + i, 'eur', { m, w: 96, ph: T.aum[i].v == null ? '' : eurF(T.aum[i].v), cls: T.aum[i].set ? '' : 'carried' }) + '</td>').join('') + '</tr></tbody></table></div></div>';
    h += '<div class="card"><h2>3. Postes de frais - fréquence de facturation</h2><div class="sub">Mensuel pour Commission de Gestion, Conseil, Distribution et Finologee ; semestriel pour Dépositaire, Valorisation et CAC ; annuel pour les autres. Le forfait CAC par distribution suit une règle à part : sa fréquence est dérivée du nombre de distributions par an (01 §3) et il est facturé au début de chaque période à hauteur de sa quote-part. La fréquence et le prorata se modifient aussi dans la Saisie (colonnes « Fact. » et « Prorata lancement »).</div><div class="tbl-wrap"><table class="tbl fm"><thead><tr><th>Poste</th><th>Bénéficiaire</th><th>Fréquence de facturation</th><th>Prorata du mois de lancement</th></tr></thead><tbody>';
    (m.lines || []).forEach((l, i) => { if (l.actif !== 'Oui') return; h += '<tr><td>' + esc(l.nom) + '</td><td><small>' + esc(l.beneficiaire || '') + '</small></td><td>' + (l.regle_facturation === 'par_distribution' ? calc(esc(lineFreq(l, m)), 'dérivée du nombre de distributions par an') : inp('lines.' + i + '.freq', 'sel', { m, opts: Object.keys(FREQ_N) })) + '</td><td>' + inp('lines.' + i + '.prorata_lancement', 'sel', { m, opts: ON }) + '</td></tr>'; });
    h += '</tbody></table></div></div>';
    h += '<div class="card"><div class="kpis">' + kpi('Provisions - ' + n12 + ' premiers mois', fEur(sum(T.provT.slice(0, n12))) + ' <span class="u">EUR</span>', 'facturé ' + fEur(sum(T.invT.slice(0, n12))) + ' · décaissé ' + fEur(sum(T.cashT.slice(0, n12)))) + kpi('Provisions - horizon complet (' + T.H + ' mois)', fEur(sum(T.provT)) + ' <span class="u">EUR</span>', 'facturé ' + fEur(sum(T.invT)) + ' · décaissé ' + fEur(sum(T.cashT))) + kpi('Frais de Constitution', fEur(T.reconnuEx1) + ' <span class="u">EUR</span>', 'reconnus sur le 1er exercice · ' + fEur(T.billed) + ' € réellement engagés') + '<div class="kpi"><div class="k">Export</div><div class="v"><button type="button" class="mini" data-fee-export-csv="treso">Exporter le prévisionnel en CSV</button></div><div class="s">§4 à §7, un bloc par vue</div></div></div></div>';
    h += '<details class="fm-details"><summary>4. Montant annuel équivalent au barème - calculé mois par mois, à l’Actif Net de ce mois</summary>' + block('4. Montant annuel équivalent', 'Rejoue le barème de chaque poste à l’Actif Net du mois concerné : le montant annuel équivalent change de mois en mois si l’Actif Net évolue. Sert de base à la provision comptable (§5).', 'annual', T.months.map((_, i) => act.reduce((s, r) => s + r.annual[i], 0))) + '</details>';
    h += block('5. Vue comptable - provisions mensuelles', 'Charge lissée à l’Actif Net de chaque mois (un douzième du montant annuel équivalent), proratisée en Jours Ouvrés sur le mois de lancement, sauf Finologee. C’est le montant à passer à chaque VL.', 'prov', T.provT, { label: 'Frais de Constitution (provision, voir §8)', vals: T.provC });
    h += '<details class="fm-details"><summary>6. Facturation - échéances par poste</summary>' + block('6. Facturation', 'Facture émise le mois de fin de période calendaire (mensuelle ; trimestrielle mars, juin, sept., déc. ; semestrielle juin, déc. ; annuelle déc.) pour la somme des provisions accumulées depuis la précédente échéance - y compris le prorata du premier mois et les variations d’Actif Net. Intermédiaire : sert à construire la vue cash.', 'inv', T.invT) + '</details>';
    h += block('7. Vue cash - décaissements mensuels', 'Facture émise en fin de période, payée le mois suivant, quelle que soit la fréquence. Les Frais de Constitution sont décaissés à la date de paiement effective de chaque facture réelle (§8).', 'cash', T.cashT, { label: 'Frais de Constitution (cash, voir §8)', vals: T.cashC });
    /* §8 */
    h += '<div class="card"><h2>8. Frais de Constitution - calendrier réglementaire et suivi réel</h2><div class="sub">Actifs par défaut : le plancher (' + fEur(T.plancher) + ' €) est provisionné sur le 1er exercice comptable quoi qu’il arrive. Au 1er anniversaire du fonds, le montant dû - MAX(plancher ; taux × Actif Net au 1er anniversaire) - est comparé au plancher : le complément éventuel est provisionné sur l’exercice qui contient cet anniversaire, sans lissage au-delà (indication du CAC). Dès qu’une facture réelle est datée, son montant est déduit du solde restant, ré-étalé à parts égales sur les mois non encore facturés de la fenêtre.</div>';
    h += '<div class="tbl-wrap"><table class="tbl fm"><thead><tr><th>Exercice</th><th>Début</th><th>Fin</th><th class="num">Nb mois</th></tr></thead><tbody>' + T.ex.map(x => '<tr><td>Exercice ' + x.k + (x.k === 1 ? ' (1er, partiel)' : '') + '</td><td>' + fDate(isoOf(x.start)) + '</td><td>' + fDate(isoOf(x.end)) + '</td><td class="num">' + x.nbm + '</td></tr>').join('') + '</tbody></table></div>';
    h += '<div class="kv small" style="margin-top:10px"><div class="k">1er anniversaire du fonds</div><div class="v">' + fDate(isoOf(T.anniv)) + '</div><div class="k">Exercice contenant l’anniversaire</div><div class="v">' + (T.exA ? 'Exercice ' + T.exA.k + ' - du ' + fDate(isoOf(T.exA.start)) + ' au ' + fDate(isoOf(T.exA.end)) + ' (' + T.exA.nbm + ' mois)' : 'hors des cinq premiers exercices') + '</div><div class="k">Montant total dû au Prospectus</div><div class="v">' + fEur(T.totalDu) + ' € - MAX(' + fEur(T.plancher) + ' € ; ' + fRate((m.params || {}).taux_constitution) + ' × ' + fEur((m.params || {}).actif_net_prev_an1) + ' €)</div><div class="k">Montant reconnu sur l’Exercice 1</div><div class="v">' + fEur(T.reconnuEx1) + ' €</div><div class="k">Complément sur l’exercice de l’anniversaire</div><div class="v">' + fEur(T.complement) + ' €</div><div class="k">Déjà facturé (réel) sur l’Exercice 1</div><div class="v">' + fEur(T.am1) + ' € sur ' + T.nb1 + ' mois - reste ' + fEur(Math.max(0, T.reconnuEx1 - T.am1)) + ' € réparti sur ' + Math.max(0, T.ex[0].nbm - T.nb1) + ' mois</div></div>';
    h += '<h3 class="sub">Suivi des Frais de Constitution réellement engagés</h3><div class="lead-p">Renseignez chaque facture réelle reçue (avocat, dépositaire, AMF…). La date de facturation pilote le ré-amortissement de la provision ; le décaissement suit la date de paiement - à défaut, le mois suivant la facturation.</div>';
    h += '<div class="tbl-wrap"><table class="tbl fm"><thead><tr><th>Type de frais</th><th class="num">Montant (€)</th><th>Date de facturation</th><th>Date de paiement (optionnel)</th><th>Paiement effectif (calculé)</th><th></th></tr></thead><tbody>';
    T.facts.forEach((x, i) => { h += '<tr><td><input type="text" class="fm-in txt" data-inv="' + i + '" data-f="label" value="' + esc(x.label || '') + '" style="width:220px" placeholder="fournisseur"></td><td class="num"><input type="number" step="1" class="fm-in" data-inv="' + i + '" data-f="amount" value="' + (x.amount == null ? '' : x.amount) + '"></td><td><input type="date" class="fm-in" data-inv="' + i + '" data-f="inv" value="' + esc(x.inv || '') + '"></td><td><input type="date" class="fm-in" data-inv="' + i + '" data-f="pay" value="' + esc(x.pay || '') + '"></td><td>' + (x.eff ? esc(ymF(ymOf(x.eff))) : muted('·')) + '</td><td><button type="button" class="mini danger" data-inv-del="' + i + '">Supprimer</button></td></tr>'; });
    if (!T.facts.length) h += '<tr><td colspan="6">' + muted('Aucune facture saisie : le montant reconnu est étalé à parts égales sur les mois du 1er exercice.') + '</td></tr>';
    h += '<tr class="total"><td>TOTAL réel engagé (€)</td><td class="num">' + fEur(T.billed) + '</td><td colspan="4"><small>reste dans la poche sur le forfait Année 1 : ' + fEur(T.plancher - T.billed) + ' €</small></td></tr></tbody></table></div><div class="tools" style="justify-content:flex-start"><button type="button" class="mini" data-inv-add>Ajouter une facture</button></div></div>';
    return h;
  }
  function viewBreakeven(m) {
    const r = recapOf(m); const a0 = feeAumRef(m); const t0 = totalAt(m, a0, r.R); const bp = refBp(m); const be = breakevenOf(m);
    const env = a0 * bp / 10000;
    let h = '<div class="lead-p derived-lead"><span class="pill reading">Outil d’interface</span> Lecture propre à l’Operating Memorandum (hors classeur) : l’Actif Net à partir duquel la charge directe annuelle du Fonds passe sous un taux de référence. ' + esc((m.reference || {}).note || '') + '</div>';
    h += '<div class="card"><div class="kpis">' + aumRefKpi(m) +
      '<div class="kpi"><div class="k">Taux de référence</div><div class="v"><input type="number" step="1" data-fv="reference.taux_bp" data-kind="int" class="fm-in' + (isOv('reference.taux_bp') ? ' ov' : '') + '" style="width:72px" value="' + bp + '"> <span class="u">bp</span></div><div class="s">' + fPct(bp / 10000) + ' de l’Actif Net</div></div>' +
      kpi('Charge au taux de référence', fEur(env) + ' <span class="u">EUR</span>', 'à ' + fM(a0)) + kpi('Frais directs du Fonds', fEur(t0.total) + ' <span class="u">EUR</span>', fBp(a0 > 0 ? t0.total / a0 : null)) +
      '<div class="kpi"><div class="k">Écart au taux de référence</div><div class="v' + (env - t0.total < 0 ? ' neg' : '') + '">' + fEur(env - t0.total) + ' <span class="u">EUR</span></div><div class="s">' + fBp(a0 > 0 ? (env - t0.total) / a0 : null) + '</div></div>' +
      kpi('Actif Net au taux de référence', esc(beText(be)), 'seuil à partir duquel les frais directs passent sous ' + bp + ' bp') + '</div></div>';
    h += '<div class="card"><h2>Charge au taux de référence et frais directs en fonction de l’Actif Net</h2><div class="sub">Le seuil est l’Actif Net où la droite du taux de référence croise la courbe des frais directs.</div><div class="swim-wrap">' + feeCurve(m, 'breakeven') + '</div></div>';
    h += '<div class="card"><h2>Par palier d’Actif Net</h2><div class="tbl-wrap"><table class="tbl fees fm wide"><thead><tr><th>Ligne</th>' + r.ladder.map(a => '<th class="num">' + fM(a) + '</th>').join('') + '</tr></thead><tbody>';
    h += '<tr class="total"><td>Charge au taux de référence (' + bp + ' bp × Actif Net)</td>' + r.ladder.map(a => '<td class="num">' + fEur(a * bp / 10000) + '</td>').join('') + '</tr>';
    r.rows.filter(x => x.l.actif === 'Oui').forEach(x => { h += '<tr class="costrow"><td>− ' + esc(x.l.nom) + '</td>' + x.vals.map(v => '<td class="num">' + fEur(v) + '</td>').join('') + '</tr>'; });
    h += '<tr class="total"><td>Frais directs du Fonds</td>' + r.tot.map(v => '<td class="num">' + fEur(v) + '</td>').join('') + '</tr>';
    h += '<tr class="total"><td>Écart au taux de référence</td>' + r.tot.map((v, i) => { const n = r.ladder[i] * bp / 10000 - v; return '<td class="num' + (n < 0 ? ' neg' : '') + '">' + fEur(n) + '<br><small>' + fBp(r.ladder[i] > 0 ? n / r.ladder[i] : null) + '</small></td>'; }).join('') + '</tr></tbody></table></div></div>';
    return h;
  }
  /* ---------------- courbe SVG ---------------- */
  function feeCurve(m, mode) {
    const W = 760, H = 300, L = 66, Rm = 66, T = 16, B = 40; const R = refsOf(m);
    const lad = ladderOf(m); const a0 = feeAumRef(m) / M; const bp = refBp(m); const be = breakevenOf(m);
    const beMax = Math.max(0, ...be.map(x => x.aum));
    const maxA = mode === 'breakeven' ? Math.max(40, Math.ceil(beMax * 3.5 / 20) * 20) : Math.max(lad.length ? lad[lad.length - 1] / M : 50, a0 * 1.2);
    const minA = mode === 'breakeven' ? 0 : Math.max(0.5, Math.min(5, (lad.length ? lad[0] / M : 5) / 2));
    const pts = []; const step = Math.max(0.25, (maxA - minA) / 200); for (let a = Math.max(minA, 0.25); a <= maxA + 1e-9; a += step) { const t = totalAt(m, a * M, R); pts.push([a, { total: t.total, env: a * M * bp / 10000, net: a * M * bp / 10000 - t.total, allin: t.total / (a * M) }]); }
    const x = a => L + ((a - minA) / (maxA - minA)) * (W - L - Rm);
    let s = '<svg class="curve" xmlns="http://www.w3.org/2000/svg" width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Frais en fonction de l’Actif Net">';
    [0, 0.25, 0.5, 0.75, 1].forEach(f => { const a = minA + (maxA - minA) * f; s += '<text class="tk" x="' + x(a) + '" y="' + (H - B + 16) + '" text-anchor="middle">' + mRound(a) + '</text>'; });
    s += '<text class="tk" x="' + ((L + W - Rm) / 2) + '" y="' + (H - 6) + '" text-anchor="middle">Actif Net du Fonds (M€)</text>';
    if (mode === 'breakeven') {
      const vals = pts.map(p => [p[1].env, p[1].total, p[1].net]).flat(); const ymax = Math.max(1, ...vals) * 1.1, ymin = Math.min(0, ...vals) * 1.1;
      const y = v => T + (ymax - v) / (ymax - ymin) * (H - T - B);
      for (let k = 0; k <= 4; k++) { const v = ymin + (ymax - ymin) * k / 4; s += '<line class="gl" x1="' + L + '" x2="' + (W - Rm) + '" y1="' + y(v) + '" y2="' + y(v) + '"/><text class="tk" x="' + (L - 6) + '" y="' + (y(v) + 4) + '" text-anchor="end">' + eurF(v / 1000) + 'k</text>'; }
      s += '<line class="zero" x1="' + L + '" x2="' + (W - Rm) + '" y1="' + y(0) + '" y2="' + y(0) + '"/>';
      const path = (k, cls) => '<path class="' + cls + '" d="' + pts.map((p, i) => (i ? 'L' : 'M') + x(p[0]).toFixed(1) + ',' + y(p[1][k]).toFixed(1)).join(' ') + '"/>';
      s += path('env', 'c-rev') + path('total', 'c-cost') + path('net', 'c-margin');
      be.filter(r => !r.below && r.aum >= minA && r.aum <= maxA).forEach(r => { s += '<line class="be" x1="' + x(r.aum) + '" x2="' + x(r.aum) + '" y1="' + T + '" y2="' + (H - B) + '"/><text class="tk be" x="' + (x(r.aum) + 5) + '" y="' + (T + 10) + '">seuil ' + mRound(r.aum) + '</text>'; });
      s += '<g class="legend" transform="translate(' + (L + (W - L - Rm) * 0.36) + ',' + (T + 6) + ')"><rect width="230" height="50" rx="4"/><line class="c-rev" x1="8" x2="28" y1="12" y2="12"/><text x="34" y="16">charge au taux de référence</text><line class="c-cost" x1="8" x2="28" y1="26" y2="26"/><text x="34" y="30">frais directs du Fonds</text><line class="c-margin" x1="8" x2="28" y1="40" y2="40"/><text x="34" y="44">écart au taux de référence</text></g>';
    } else {
      const ymaxE = Math.max(1, ...pts.map(p => p[1].total)) * 1.1; const ymaxB = Math.max(1e-6, ...pts.map(p => p[1].allin).filter(v => isFinite(v))) * 1.1;
      const yE = v => T + (ymaxE - v) / ymaxE * (H - T - B), yB = v => T + (ymaxB - v) / ymaxB * (H - T - B);
      for (let k = 0; k <= 4; k++) { const v = ymaxE * k / 4; s += '<line class="gl" x1="' + L + '" x2="' + (W - Rm) + '" y1="' + yE(v) + '" y2="' + yE(v) + '"/><text class="tk" x="' + (L - 6) + '" y="' + (yE(v) + 4) + '" text-anchor="end">' + eurF(v / 1000) + 'k</text>'; }
      for (let k = 0; k <= 4; k++) { const v = ymaxB * k / 4; s += '<text class="tk" x="' + (W - Rm + 6) + '" y="' + (yB(v) + 4) + '">' + fBp(v) + '</text>'; }
      s += '<path class="c-eur" d="' + pts.map((p, i) => (i ? 'L' : 'M') + x(p[0]).toFixed(1) + ',' + yE(p[1].total).toFixed(1)).join(' ') + '"/>';
      s += '<path class="c-bps" d="' + pts.map((p, i) => (i ? 'L' : 'M') + x(p[0]).toFixed(1) + ',' + yB(p[1].allin).toFixed(1)).join(' ') + '"/>';
      s += '<g class="legend" transform="translate(' + (L + 8) + ',' + (T + 6) + ')"><rect width="250" height="32" rx="4"/><line class="c-eur" x1="8" x2="28" y1="12" y2="12"/><text x="34" y="16">frais directs annuels (€, à gauche)</text><line class="c-bps" x1="8" x2="28" y1="26" y2="26"/><text x="34" y="30">taux all-in (bp, à droite)</text></g>';
    }
    if (a0 >= minA && a0 <= maxA) s += '<line class="now" x1="' + x(a0) + '" x2="' + x(a0) + '" y1="' + T + '" y2="' + (H - B) + '"/>';
    return s + '</svg>';
  }
  /* ---------------- exports ---------------- */
  const csvOf = rows => '﻿' + rows.map(r => r.map(v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"').join(';')).join('\r\n');
  function feesCSV(kind) {
    const m = model();
    if (kind === 'treso') {
      const T = tresoOf(m); const act = T.rows.filter(r => r.l.actif === 'Oui'); const rows = [['Prévisionnel de trésorerie - ' + (m.nom_fonds || ''), 'lancement ' + fDate(isoOf(T.launch)), 'fin 1er exercice ' + fDate(isoOf(T.fye1)), 'prorata mois de lancement ' + fPct(T.prorata0)]];
      rows.push([]); rows.push(['Actif Net mensuel (€)', ''].concat(T.months.map(mo => ymF(mo.ym)))); rows.push(['', ''].concat(T.aum.map(a => a.v == null ? '' : Math.round(a.v))));
      [['4. Montant annuel équivalent (€)', 'annual', null], ['5. Provision comptable (€)', 'prov', 'provC'], ['6. Facturation (€)', 'inv', null], ['7. Décaissement (€)', 'cash', 'cashC']].forEach(([title, k, extra]) => {
        rows.push([]); rows.push([title, 'Fréquence'].concat(T.months.map(mo => ymF(mo.ym))));
        act.forEach(r => rows.push([r.l.nom, r.freq].concat(r[k].map(v => Math.round(v * 100) / 100))));
        if (extra) rows.push(['Frais de Constitution', ''].concat(T[extra].map(v => Math.round(v * 100) / 100)));
        rows.push(['TOTAL', ''].concat((k === 'annual' ? T.months.map((_, i) => act.reduce((s, r) => s + r.annual[i], 0)) : k === 'prov' ? T.provT : k === 'inv' ? T.invT : T.cashT).map(v => Math.round(v * 100) / 100)));
      });
      return csvOf(rows);
    }
    const r = recapOf(m); const rows = [['Poste de frais', 'Bénéficiaire', 'Actif ?'].concat(r.ladder.flatMap(a => [fM(a) + ' (€)', fM(a) + ' (% AUM)']))];
    r.rows.forEach(x => rows.push([x.l.nom, x.l.beneficiaire || '', x.l.actif].concat(x.vals.flatMap((v, i) => [Math.round(v), r.ladder[i] > 0 ? (v / r.ladder[i] * 100).toFixed(4).replace('.', DEC) : '']))));
    rows.push(['TOTAL (€)', '', ''].concat(r.tot.flatMap((v, i) => [Math.round(v), (r.pct[i] * 100).toFixed(4).replace('.', DEC)])));
    rows.push(['TVA (€)', '', ''].concat(r.vat.flatMap(v => [Math.round(v), ''])));
    rows.push([]); rows.push(['Actif Net au taux de référence (' + refBp(m) + ' bp)', beText(breakevenOf(m))]);
    return csvOf(rows);
  }
  function feesExportJSON() { return JSON.stringify({ om: ID.id, fees_model: FEES.source_file || '', exported_at: new Date().toISOString(), ov: FS.ov || {}, invoices: FS.invoices, aum: FS.aum }, null, 1); }
  function feesImportJSON(txt) { const o = JSON.parse(txt); if (!o || typeof o !== 'object' || (!o.ov && !o.invoices)) throw new Error('format'); FS.ov = (o.ov && typeof o.ov === 'object') ? o.ov : {}; FS.invoices = Array.isArray(o.invoices) ? o.invoices : null; if (o.aum != null) FS.aum = o.aum; }

  /* ---------------- rendu de l'onglet ---------------- */
  function renderFees(tab) {
    fsNorm(); const m = model();
    let h = '<div class="mode-switch" role="group" aria-label="Vue">' + [['saisie', '01 · Saisie'], ['recap', '04 · Récap par palier'], ['detail', '03 · Détail Dépositaire / Valorisation'], ['provision', 'Provision à la VL'], ['treso', '05 · Trésorerie'], ['breakeven', 'Seuil de rentabilité']].map(([k, lab]) => '<button type="button" data-fee-mode="' + k + '" aria-pressed="' + (FS.mode === k) + '">' + lab + '</button>').join('') + '</div>';
    try {
      if (FS.mode === 'saisie') h += viewSaisie(m); else if (FS.mode === 'detail') h += viewDetail(m); else if (FS.mode === 'provision') h += viewProvision(m); else if (FS.mode === 'treso') h += viewTreso(m); else if (FS.mode === 'breakeven') h += viewBreakeven(m); else h += viewRecap(m);
    } catch (err) { h += '<div class="card"><div class="lead-p flagbox">Le modèle de frais n’a pas pu être calculé avec les saisies actuelles (' + esc(String(err && err.message || err)) + '). <button type="button" class="mini danger" data-fee-reset>Revenir au classeur</button></div></div>'; }
    return h;
  }
  function rerenderFees() { saveFees(); if (state.tab === 'fees-tool') { const y = window.scrollY; render({ keepScroll: true }); window.scrollTo(0, y); } }
  document.addEventListener('change', e => {
    const t = e.target; if (!(t instanceof Element)) return;
    if (t.matches('[data-fv]')) {
      const path = t.dataset.fv, kind = t.dataset.kind; let v;
      if (kind === 'sel' || kind === 'txt' || kind === 'date') v = String(t.value); else if (kind === 'pct') v = t.value === '' ? null : +t.value / 100; else v = t.value === '' ? null : +t.value;
      if (kind === 'sel' && /\.tva$/.test(path)) v = +v;
      if (kind === 'int' && v != null) v = Math.round(v);
      if (JSON.stringify(v) === JSON.stringify(getPath(FEES, path))) delete FS.ov[path]; else FS.ov[path] = v;
      rerenderFees(); return;
    }
    if (t.matches('[data-fee-aum]')) { FS.aum = +t.value > 0 ? +t.value : null; rerenderFees(); return; }
    if (t.matches('[data-fee-pmonth]')) { FS.pmonth = /^\d{4}-\d{2}$/.test(t.value) ? t.value : null; rerenderFees(); return; }
    if (t.matches('[data-inv]')) {
      const i = +t.dataset.inv, fld = t.dataset.f; const list = invoicesOf(model()).slice(); if (!list[i]) return;
      list[i] = Object.assign({}, list[i], { [fld]: fld === 'amount' ? (t.value === '' ? null : +t.value) : String(t.value) }); FS.invoices = list;
      const sel = '[data-inv="' + i + '"][data-f="' + fld + '"]'; rerenderFees(); setTimeout(() => { const el = document.querySelector(sel); if (el) el.focus(); }, 0); return;
    }
    if (t.matches('[data-fee-import-file]')) {
      const f = t.files && t.files[0]; if (!f) return;
      const rd = new FileReader(); rd.onload = () => { try { feesImportJSON(String(rd.result)); rerenderFees(); } catch (err) { alert('Fichier de saisies illisible : attendu un export JSON de cet outil.'); } }; rd.readAsText(f); t.value = ''; return;
    }
  });
  document.addEventListener('click', e => {
    const t = e.target; if (!(t instanceof Element)) return;
    const md = t.closest('[data-fee-mode]'); if (md) { FS.mode = md.dataset.feeMode; rerenderFees(); return; }
    if (t.closest('[data-inv-add]')) { FS.invoices = invoicesOf(model()).concat([{ id: 'n' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7), label: '', amount: null, inv: '', pay: '' }]); rerenderFees(); return; }
    const dl = t.closest('[data-inv-del]'); if (dl) { const i = +dl.dataset.invDel; FS.invoices = invoicesOf(model()).filter((x, k) => k !== i); rerenderFees(); return; }
    if (t.closest('[data-fee-reset]')) { if (!confirm('Revenir aux valeurs du classeur ? Toutes les saisies du modèle de frais (paramètres, barème, Actif Net mensuel, factures) seront effacées.')) return; FS = Object.assign(FS_EMPTY(), { mode: FS.mode }); rerenderFees(); return; }
    const ex = t.closest('[data-fee-export-csv]'); if (ex) { offerFile(ID.export_prefix + '_frais_' + ex.dataset.feeExportCsv + '_' + new Date().toISOString().slice(0, 10) + '.csv', feesCSV(ex.dataset.feeExportCsv), 'txt'); return; }
    if (t.closest('[data-fee-export-json]')) { offerFile(ID.export_prefix + '_frais_saisies_' + new Date().toISOString().slice(0, 10) + '.json', feesExportJSON(), 'txt'); return; }
    if (t.closest('[data-fee-import-json]')) { const f = document.querySelector('[data-fee-import-file]'); if (f) f.click(); return; }
  });

