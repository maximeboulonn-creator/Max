/* Risk Studio - kit commun aux consoles : cadre, fragments réutilisables,
   vues secondaires (liquidité, indicateurs, échéancier, onboarding, studios)
   et routage. Chaque console fournit seulement son écran d'accueil. */
(function () {
"use strict";
const R = window.RS, D = R.DATA;
const st = { i: D.meta.arretes.length - 1, vue: "accueil", filtre: { statut: "", fonds: "" }, focus: null };
let page = null, VUES = [];
const $ = id => document.getElementById(id);

const C = (code, long) => '<span class="c-' + code + '">' + R.esc(long ? R.STATUTS[code].lib : R.STATUTS[code].court) + '</span>';
const seuil = (ind, k) => ind[k] === null ? "-" : R.val(ind[k], ind.u);
const pl = (n, s, p) => n > 1 ? p : s;
const arrete = () => D.meta.arretes[st.i];

/* ------------------------------------------------------------ fragments */
function spark(ind, w, h) {
  w = w || 164; h = h || 44;
  const vals = ind.h.filter(v => v !== null);
  const pool = vals.concat([ind.al, ind.lim].filter(v => v !== null));
  let lo = Math.min.apply(null, pool), hi = Math.max.apply(null, pool);
  if (hi - lo < 1e-9) { lo -= 1; hi += 1; }
  const pad = (hi - lo) * 0.12; lo -= pad; hi += pad;
  const px = 5, py = 5;
  const X = k => px + k * (w - 2 * px) / (ind.h.length - 1);
  const Y = v => py + (hi - v) * (h - 2 * py) / (hi - lo);
  let s = '<svg class="spark" width="' + w + '" height="' + h + '" viewBox="0 0 ' + w + ' ' + h + '" role="img" aria-label="' +
    R.esc("Évolution sur quatre arrêtés : " + ind.h.map(v => R.val(v, ind.u)).join(", ")) + '">';
  if (ind.al !== null) s += '<line class="al" x1="0" x2="' + w + '" y1="' + Y(ind.al).toFixed(1) + '" y2="' + Y(ind.al).toFixed(1) + '"/>';
  if (ind.lim !== null) s += '<line class="lim" x1="0" x2="' + w + '" y1="' + Y(ind.lim).toFixed(1) + '" y2="' + Y(ind.lim).toFixed(1) + '"/>';
  let path = "", open = false;
  ind.h.forEach((v, k) => { if (v === null) { open = false; return; } path += (open ? " L" : " M") + X(k).toFixed(1) + " " + Y(v).toFixed(1); open = true; });
  if (path) s += '<path class="ln" d="' + path.trim() + '"/>';
  ind.h.forEach((v, k) => {
    if (v !== null) s += '<circle class="pt' + (k === st.i ? ' cur' : '') + '" cx="' + X(k).toFixed(1) + '" cy="' + Y(v).toFixed(1) + '" r="' + (k === st.i ? 3.2 : 2.4) + '"/>';
  });
  return s + '</svg>';
}

function article(ind) {
  const i = st.i, code = R.statut(ind, i), m = R.marge(ind, i), tr = R.tendance(ind, i);
  const a = ind.act ? R.actById[ind.act] : null;
  const seuils = ind.lim === null ? "Alerte " + seuil(ind, "al") + ", sans limite"
    : (ind.sens === "max" ? "Plafond " : "Plancher ") + seuil(ind, "lim") + " - alerte " + seuil(ind, "al");
  return '<article class="it" id="it-' + ind.id + '">' +
    '<div class="nm"><b>' + R.t(ind.lib) + '</b><span class="muted">' + R.t(ind.def) + '</span></div>' +
    '<div class="val"><b class="c-' + code + '">' + R.val(R.valeur(ind, i), ind.u) + '</b><span class="muted">' + R.t(seuils) + '</span></div>' +
    '<div>' + spark(ind) + '<div class="sx"><span>' + R.dateCourte(D.meta.arretes[0]) + '</span><span>' + R.dateCourte(D.meta.arretes[D.meta.arretes.length - 1]) + '</span></div></div>' +
    '<div>' + C(code, true) + '<br><span class="muted">' + (m === null ? "Marge n.d." : "Marge " + R.signe(m, ind.u)) + '<br>' + R.esc(tr.lib) + '</span></div>' +
    '<dl class="more">' +
    (ind.com ? '<dt>Commentaire</dt><dd>' + R.t(ind.com) + '</dd>' : '') +
    (a ? '<dt>Action</dt><dd>' + R.t(a.objet) + ' <span class="muted">' + R.esc(a.resp + ", " + R.date(a.ech) + ", priorité " + a.prio.toLowerCase()) + '</span></dd>' : '') +
    '<dt>Nature</dt><dd>' + R.t(ind.type + ", " + ind.nat.toLowerCase() + ". " + ind.ref) + '</dd>' +
    '<dt>Source</dt><dd class="muted">' + R.t(ind.src) + '</dd>' +
    '</dl></article>';
}

function todo(a) {
  const qui = a.f ? R.fondsById[a.f].court : R.pipeById[a.pipe].nom;
  const ind = a.ind ? R.indById[a.ind] : null;
  return '<li class="todo"><div class="when"><b>' + R.dateCourte(a.ech) + '</b><span>' + R.esc(R.delai(a.ech)) + '</span></div><div class="what">' +
    '<p><span class="prio-' + a.prio + '">' + R.esc("Priorité " + a.prio.toLowerCase()) + '</span> <span class="muted">- ' + R.esc(qui) + '</span></p>' +
    '<p>' + R.t(a.objet) + '</p>' +
    (ind ? '<button type="button" class="lk" data-open="' + ind.id + '">Voir l\'indicateur</button>'
         : '<button type="button" class="lk" data-vue-go="pipeline">Voir le dossier</button>') +
    '</div></li>';
}

function echItem(e) {
  return '<li class="todo"><div class="when"><b>' + R.dateCourte(e.date) + '</b><span>' + R.esc(R.delai(e.date)) + '</span></div><div class="what">' +
    '<p>' + R.t(e.objet) + '</p><p class="muted">' + R.esc(e.f.map(x => R.fondsById[x].court).join(", ")) + ' - <span class="' +
    (e.statut === "À risque" || e.statut === "Non lancé" ? "st-risk" : "") + '">' + R.esc(e.statut) + '</span></p></div></li>';
}

/* barres de couverture de liquidité d'un fonds, dessinées à l'échelle */
function liqBlock(id) {
  const cfg = R.copieLiquidite(id), res = {};
  R.SCENARIOS.forEach(s => { res[s.k] = R.liquidite(cfg, s.k); });
  const kri = D.indicateurs.find(x => x.f === id && x.d === "LIQ" && x.u === "x");
  const al = kri ? kri.al : null;
  const maxv = Math.max.apply(null, R.SCENARIOS.map(s => res[s.k].couverture || 0).concat([al || 0, 1]));
  const xmax = Math.max(2, Math.ceil(maxv * 1.1 * 2) / 2);
  const pc = v => (Math.max(0, Math.min(v, xmax)) / xmax * 100).toFixed(2) + "%";
  const refs = '<span class="ref" style="left:' + pc(1) + '"></span>' + (al ? '<span class="ref al" style="left:' + pc(al) + '"></span>' : '');
  const bars = R.SCENARIOS.map(s => {
    const v = res[s.k].couverture;
    return '<div class="liq-row"><span>' + s.lib + '</span><div class="track"><span class="b' + (v < 1 ? ' low' : '') + '" style="width:' + pc(v) + '"></span>' + refs + '</div>' +
      '<span class="n' + (v < 1 ? ' c-DEP' : '') + '">' + R.val(v, "x") + '</span></div>';
  }).join("");
  const ticks = [];
  for (let x = 0; x <= xmax + 1e-9; x += 0.5) ticks.push('<span style="left:' + pc(x) + '">' + R.nombre(x, 1) + 'x</span>');
  const eur = v => R.nombre(v, 1);
  const sg = v => (v > 0.05 ? "+" : v < -0.05 ? "-" : "") + R.nombre(Math.abs(v), 1);
  const tab = '<table class="list" style="margin-top:12px"><thead><tr><th>EUR m</th>' + R.SCENARIOS.map(s => '<th class="n">' + s.lib + '</th>').join("") + '</tr></thead><tbody>' +
    '<tr><td>Ressources</td>' + R.SCENARIOS.map(s => '<td class="n">' + eur(res[s.k].ressources) + '</td>').join("") + '</tr>' +
    '<tr><td>' + (cfg.type === "ouvert" ? "Rachats servis et appels" : "Décaissements prévus") + '</td>' + R.SCENARIOS.map(s => '<td class="n">' + eur(res[s.k].besoins) + '</td>').join("") + '</tr>' +
    (cfg.type === "ouvert" ? '<tr><td>Rachats différés par la gate</td>' + R.SCENARIOS.map(s => '<td class="n">' + eur(res[s.k].differes) + '</td>').join("") + '</tr>' : '') +
    '<tr><td><b>Solde</b></td>' + R.SCENARIOS.map(s => '<td class="n' + (res[s.k].solde < 0 ? ' c-DEP' : '') + '">' + sg(res[s.k].solde) + '</td>').join("") + '</tr></tbody></table>';
  return '<p class="muted" style="margin:0 0 12px">' +
    R.t(cfg.type === "ouvert" ? "Couverture des rachats servis et des appels des fonds cibles à 3 mois" : "Couverture des décaissements à 12 mois") + '</p>' +
    '<div class="liq">' + bars + '</div><div class="axis">' + ticks.join("") + '</div>' +
    '<p class="note">' + R.t("Pointillés : plancher 1,00x" + (al ? " et seuil d'alerte " + R.val(al, "x") : "") + ".") + '</p>' + tab +
    '<p style="margin:10px 0 0"><b>Reverse stress test.</b> ' + R.t(R.reverse(cfg).texte) + '</p>';
}

/* ------------------------------------------------------------ vues secondaires */
function vLiquidite() {
  const blocs = ["AIR", "OPN", "ALE"].map(id => '<div class="panel"><h2>' + R.esc(R.fondsById[id].nom) + '</h2>' + liqBlock(id) + '</div>').join("");
  return '<div class="view-h"><h1>Liquidité</h1><span class="muted">' + R.t("Scénarios calibrés au 30/09/2026") + '</span></div>' +
    '<div class="panel" style="margin-bottom:18px"><h2 class="lbl">Hypothèses clés et limites</h2><ul class="hyp">' +
    D.hypothesesLiquidite.map(x => '<li>' + R.t(x) + '</li>').join("") + '</ul></div><div class="liq-grid">' + blocs + '</div>';
}

function vIndicateurs() {
  const i = st.i, F = st.filtre;
  const base = D.indicateurs.filter(x => !F.fonds || x.f === F.fonds);
  const l = base.filter(x => !F.statut || R.statut(x, i) === F.statut)
    .sort((a, b) => R.STATUTS[R.statut(b, i)].rang - R.STATUTS[R.statut(a, i)].rang);
  const cnt = c => base.filter(x => R.statut(x, i) === c).length;
  const chipsS = '<div class="chips" role="group" aria-label="Filtrer par statut"><span class="k">Statut</span>' +
    [["", "Tous", base.length]].concat(R.ORDRE_STATUTS.map(c => [c, R.STATUTS[c].lib, cnt(c)])).map(c =>
      '<button type="button" class="chip" data-fs="' + c[0] + '" aria-pressed="' + (F.statut === c[0]) + '">' + R.esc(c[1]) + '<span class="ct">' + c[2] + '</span></button>').join("") + '</div>';
  const chipsF = '<div class="chips" role="group" aria-label="Filtrer par fonds"><span class="k">Fonds</span>' +
    [["", "Tous"]].concat(D.fonds.map(f => [f.id, f.court])).map(c =>
      '<button type="button" class="chip" data-ff="' + c[0] + '" aria-pressed="' + (F.fonds === c[0]) + '">' + R.esc(c[1]) + '</button>').join("") + '</div>';
  const rows = l.map(x => {
    const m = R.marge(x, i), code = R.statut(x, i);
    return '<tr class="go" data-open="' + x.id + '" tabindex="0"><td class="nw">' + C(code, true) + '</td><td class="nw">' + R.esc(R.fondsById[x.f].court) + '</td>' +
      '<td class="nw">' + R.esc(R.domById[x.d].nom) + '</td><td>' + R.t(x.lib) + '</td><td class="n">' + R.val(R.valeur(x, i), x.u) + '</td>' +
      '<td class="n">' + seuil(x, "al") + '</td><td class="n">' + seuil(x, "lim") + '</td><td class="n">' + (m === null ? "-" : R.signe(m, x.u)) + '</td>' +
      '<td class="nw muted">' + R.esc(R.tendance(x, i).lib) + '</td><td>' + spark(x, 84, 22) + '</td></tr>';
  }).join("");
  return '<div class="view-h"><h1>Indicateurs</h1><span class="muted">' + l.length + ' sur ' + D.indicateurs.length + ', du plus grave au moins grave</span></div>' +
    '<div class="panel">' + chipsS + chipsF + '<div class="scroll"><table class="list"><thead><tr><th>Statut</th><th>Fonds</th><th>Domaine</th><th>Indicateur</th><th class="n">Valeur</th>' +
    '<th class="n">Alerte</th><th class="n">Limite</th><th class="n">Marge</th><th>Tendance</th><th>Évolution</th></tr></thead><tbody>' +
    (rows || '<tr><td colspan="10" class="muted">Aucun indicateur ne correspond aux filtres.</td></tr>') + '</tbody></table></div>' +
    '<p class="note">' + R.t("Sélectionner une ligne pour l'ouvrir dans l'écran " + page.accueil.toLowerCase() + ".") + '</p></div>';
}

function vEcheancier() {
  let mois = "";
  const rows = D.echeances.map(e => {
    const m = R.moisAnnee(e.date), g = m !== mois ? '<tr class="grp"><td colspan="6">' + m + '</td></tr>' : '';
    mois = m;
    return g + '<tr><td class="nw"><b>' + R.date(e.date) + '</b></td><td class="nw muted">' + R.esc(R.delai(e.date)) + '</td><td>' + R.t(e.objet) + '</td>' +
      '<td>' + R.esc(e.f.map(x => R.fondsById[x].court).join(", ")) + '</td><td class="muted">' + R.t(e.ref) + '</td>' +
      '<td class="nw ' + (e.statut === "À risque" || e.statut === "Non lancé" ? "st-risk" : "") + '">' + R.esc(e.statut) + '</td></tr>';
  }).join("");
  return '<div class="view-h"><h1>Échéancier</h1><span class="muted">Réglementaire et interne</span></div>' +
    '<div class="panel"><div class="scroll"><table class="list"><thead><tr><th>Date</th><th>Délai</th><th>Objet</th><th>Fonds</th><th>Référence</th><th>Statut</th></tr></thead><tbody>' +
    rows + '</tbody></table></div></div>';
}

function etapes(p) {
  return '<div class="seg" aria-hidden="true">' + p.etapes.map((e, k) => '<i class="' + (k < p.cur ? "done" : k === p.cur ? "cur" : "") + '"></i>').join("") + '</div>' +
    R.t("Étape " + (p.cur + 1) + " sur " + p.etapes.length + " : " + p.etapes[p.cur]);
}

function vPipeline() {
  const rows = D.pipeline.map(p => '<tr><td><b>' + (p.url ? '<a href="' + p.url + '" target="_blank" rel="noopener">' + R.esc(p.nom) + '</a>' : R.esc(p.nom)) + '</b><br><span class="muted">' + R.esc(p.nature) + '</span></td>' +
    '<td>' + etapes(p) + '</td>' +
    '<td>' + R.esc(p.prochaine) + '<br><span class="muted">' + R.date(p.ech) + ', ' + R.esc(R.delai(p.ech)) + '</span></td><td>' + R.t(p.point) + '</td></tr>').join("");
  return '<div class="view-h"><h1>Onboarding et due diligence</h1><span class="muted">' + D.pipeline.length + ' dossiers</span></div>' +
    '<div class="panel"><div class="scroll"><table class="list"><thead><tr><th>Dossier</th><th>Avancement</th><th>Prochaine étape</th><th>Point d\'attention</th></tr></thead><tbody>' + rows + '</tbody></table></div></div>';
}

function vStudios() {
  return '<div class="view-h"><h1>Studios</h1><span class="muted">Outils de la fonction de gestion des risques</span></div>' +
    '<div class="panel"><div class="studios">' + D.studios.map(s => '<div class="studio"><b>' +
      (s.url ? '<a href="' + s.url + '" target="_blank" rel="noopener">' + R.esc(s.nom) + '</a>' : R.esc(s.nom)) + '</b>' +
      '<span>' + R.esc(s.objet) + '</span><br><span class="' + (s.etat === "En service" ? "muted" : "faint") + '">' + R.esc(s.etat) + '</span></div>').join("") + '</div></div>';
}

/* ------------------------------------------------------------ cadre et routage */
const RENDU = { liquidite: vLiquidite, indicateurs: vIndicateurs, echeancier: vEcheancier, pipeline: vPipeline, studios: vStudios };

function shell() {
  $("app").innerHTML = '<div class="shell"><header class="bar">' +
    '<div class="brand"><span class="name">Risk Studio</span><span class="ent">' + R.esc(page.nom) + '</span></div>' +
    '<nav class="views" role="tablist" id="views" aria-label="Vues"></nav>' +
    '<div class="tools"><label for="arrete">Arrêté</label><select id="arrete"></select>' +
    '<button type="button" class="quiet" id="theme">Thème&nbsp;: automatique</button></div></header>' +
    '<p class="demo" id="demo"></p><main id="view"></main><p class="foot" id="foot"></p></div>';
}

function cadre() {
  $("views").innerHTML = VUES.map(v => '<button type="button" role="tab" id="tab-' + v[0] + '" data-vue="' + v[0] + '" aria-selected="' + (v[0] === st.vue) + '">' + v[1] + '</button>').join("");
  $("arrete").innerHTML = D.meta.arretes.map((a, k) => '<option value="' + k + '"' + (k === st.i ? " selected" : "") + '>' + R.date(a) + '</option>').join("");
  $("demo").textContent = R.fr("Fundcraft France. Données de démonstration : valeurs fictives, non issues des fonds. Produit le " + D.meta.produit + ".");
  $("foot").textContent = R.fr("Risk Studio, " + page.nom.toLowerCase() + ". Les quatre consoles partagent les mêmes données, les mêmes calculs et les mêmes vues secondaires ; seul l'écran d'accueil change.");
}

function rendu() {
  cadre();
  const v = $("view");
  v.setAttribute("role", "tabpanel");
  v.setAttribute("aria-labelledby", "tab-" + st.vue);
  v.innerHTML = st.vue === "accueil" ? page.accueilHtml() : RENDU[st.vue]();
  if (page.apres) page.apres();
  if (st.focus) {
    const el = $("it-" + st.focus) || document.querySelector('[data-focus="' + st.focus + '"]');
    st.focus = null;
    if (el) { el.scrollIntoView({ block: "center" }); el.classList.remove("flash"); void el.offsetWidth; el.classList.add("flash"); }
  }
}

function vue(v) {
  st.vue = v;
  try { history.replaceState(null, "", "#" + v); } catch (e) {}
  rendu(); window.scrollTo(0, 0);
}

function ouvrir(id) {
  page.ouvrir(id);
  st.focus = id; st.vue = "accueil";
  try { history.replaceState(null, "", "#accueil"); } catch (e) {}
  rendu();
}

function boot(p) {
  page = p;
  VUES = [["accueil", p.accueil], ["liquidite", "Liquidité"], ["indicateurs", "Indicateurs"], ["echeancier", "Échéancier"],
          ["pipeline", "Onboarding et due diligence"], ["studios", "Studios"]];
  shell();
  $("views").addEventListener("click", e => { const b = e.target.closest("[data-vue]"); if (b) vue(b.dataset.vue); });
  $("arrete").addEventListener("change", e => { st.i = Number(e.target.value); rendu(); });
  $("view").addEventListener("click", e => {
    if (page.click && page.click(e)) return;
    const o = e.target.closest("[data-open]"); if (o) { ouvrir(o.dataset.open); return; }
    const g = e.target.closest("[data-vue-go]"); if (g) { vue(g.dataset.vueGo); return; }
    const fg = e.target.closest("[data-filtre]"); if (fg) { st.filtre = { statut: fg.dataset.filtre, fonds: "" }; vue("indicateurs"); return; }
    const fs = e.target.closest("[data-fs]"); if (fs) { st.filtre.statut = fs.dataset.fs; rendu(); return; }
    const ff = e.target.closest("[data-ff]"); if (ff) { st.filtre.fonds = ff.dataset.ff; rendu(); }
  });
  $("view").addEventListener("keydown", e => {
    if (page.key && page.key(e)) return;
    const row = e.target.closest("tr.go[data-open]");
    if (row && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); ouvrir(row.dataset.open); }
  });
  const h0 = (location.hash || "").slice(1);
  if (h0 === "accueil" || RENDU[h0]) st.vue = h0;
  R.theme($("theme"));
  rendu();
}

window.RK = { st, C, seuil, pl, arrete, spark, article, todo, echItem, liqBlock, etapes, rendu, vue, ouvrir, boot };
})();
