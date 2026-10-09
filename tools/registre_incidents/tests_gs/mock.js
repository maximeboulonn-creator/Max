// Maquette stricte des services Google Apps Script utilisés par le script (toute API non simulée lève une erreur).
'use strict';
const fs = require('fs');
function strict(obj, nom) {
  return new Proxy(obj, { get(t, p) {
    if (typeof p === 'symbol' || p in t) return t[p];
    throw new Error('API non simulée : ' + nom + '.' + String(p));
  } });
}
const user = email => strict({ getEmail: () => email }, 'User');
function decode(v) { return v && typeof v === 'object' && v.$d ? new Date(v.$d) : v; }
function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function parseA1(a1) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a1);
  const c1 = colNum(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}
function a1(r, c) { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s + r; }
const pad = n => String(n).padStart(2, '0');
function formatDate(d, tz, pattern) {
  const f = new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' });
  const p = {}; f.formatToParts(d).forEach(x => { p[x.type] = x.value; });
  const t = { yyyy: p.year, MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second };
  return pattern.replace(/yyyy|MM|dd|HH|mm|ss/g, k => t[k]);
}
function display(v, fmt) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    const avecHeure = /h/.test(fmt || '');
    return formatDate(v, 'Europe/Paris', avecHeure ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy');
  }
  return v === null || v === undefined ? '' : String(v);
}

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  _map(fn) { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(this.r + i, this.c + j)); out.push(row); } return out; }
  getValues() { return this._map((r, c) => this.sh._get(r, c)); }
  getValue() { return this.sh._get(this.r, this.c); }
  getDisplayValues() { return this._map((r, c) => display(this.sh._get(r, c), this.sh._fmt(r, c))); }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getBackgrounds() { return this._map((r, c) => this.sh._bg(r, c)); }
  getNumberFormat() { return this.sh._fmt(this.r, this.c); }
  setValue(v) { this.sh._set(this.r, this.c, v); this.sh.ss.log.push(['setValue', this.sh.name, a1(this.r, this.c), v]); return this.proxy; }
  setValues(v) {
    if (v.length !== this.nr || v.some(x => x.length !== this.nc)) throw new Error('setValues : dimensions incohérentes');
    v.forEach((row, i) => row.forEach((x, j) => this.sh._set(this.r + i, this.c + j, x))); return this.proxy;
  }
  getA1Notation() { return this.nr * this.nc === 1 ? a1(this.r, this.c) : a1(this.r, this.c) + ':' + a1(this.r + this.nr - 1, this.c + this.nc - 1); }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getSheet() { return this.sh.proxy; }
  activate() { this.sh.ss.actif = [this.sh.name, this.getA1Notation()]; return this.proxy; }
  setFontFamily(f) { this.sh.ss.log.push(['font', this.sh.name, f]); return this.proxy; }
  setFontSize() { return this.proxy; } setFontWeight() { return this.proxy; } setFontColor() { return this.proxy; }
  setBackground() { return this.proxy; } setNumberFormat() { return this.proxy; }
  setNote(n) { this.sh.notes[this.getA1Notation()] = n; return this.proxy; }
  copyTo(dest, type, transposed) {
    if (type !== 'PASTE_VALUES' || transposed !== false) throw new Error('copyTo : arguments inattendus');
    this.sh.ss.log.push(['copyTo', this.sh.name, this.getA1Notation()]);
  }
}
class Protection {
  constructor(sh, owner) { this.sh = sh; this.desc = ''; this.unprot = []; this.editors = [owner]; this.domain = true; }
  setDescription(d) { this.desc = d; return this.proxy; } getDescription() { return this.desc; }
  setUnprotectedRanges(r) { this.unprot = r; return this.proxy; }
  addEditor(u) { const e = typeof u === 'string' ? u : u.getEmail(); if (!this.editors.includes(e)) this.editors.push(e); return this.proxy; }
  removeEditors(us) { const es = us.map(u => typeof u === 'string' ? u : u.getEmail()); this.editors = this.editors.filter(e => e === this.sh.ss.owner || !es.includes(e)); return this.proxy; }
  getEditors() { return this.editors.map(user); }
  canDomainEdit() { return this.domain; } setDomainEdit(b) { this.domain = b; return this.proxy; }
  remove() { this.sh.protections = this.sh.protections.filter(p => p !== this.proxy); }
}
class Sheet {
  constructor(ss, name, data) {
    this.ss = ss; this.name = name; this.values = (data ? data.values : [[]]).map(r => r.map(decode));
    this.bgs = data ? data.bg : [[]]; this.fmts = data ? data.fmt : [[]];
    this.protections = []; this.hidden = false; this.frozenR = 0; this.frozenC = 0; this.gridHidden = true; this.notes = {}; this.widths = {};
    this.maxRows = Math.max(this.values.length, data ? this.values.length : 1000); this.maxCols = Math.max(...this.values.map(r => r.length), data ? 1 : 26);
  }
  _get(r, c) { const row = this.values[r - 1]; const v = row ? row[c - 1] : ''; return v === undefined || v === null ? '' : v; }
  _set(r, c, v) { while (this.values.length < r) this.values.push([]); const row = this.values[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; this.maxRows = Math.max(this.maxRows, r); this.maxCols = Math.max(this.maxCols, c); }
  _bg(r, c) { const row = this.bgs[r - 1]; return (row && row[c - 1]) || '#ffffff'; }
  _fmt(r, c) { const row = this.fmts[r - 1]; return (row && row[c - 1]) || 'General'; }
  getName() { return this.name; }
  getParent() { return this.ss.proxy; }
  getRange(a, b, c, d) {
    let args = typeof a === 'string' ? parseA1(a) : [a, b, c === undefined ? 1 : c, d === undefined ? 1 : d];
    if (args.some(x => !Number.isInteger(x) || x < 1 && x !== 0) || args[0] < 1 || args[1] < 1) throw new Error('getRange : arguments invalides ' + JSON.stringify(args));
    const r = new Range(this, ...args); r.proxy = strict(r, 'Range'); return r.proxy;
  }
  getLastRow() { let n = 0; this.values.forEach((row, i) => { if (row.some(v => v !== '' && v !== null && v !== undefined)) n = i + 1; }); return n; }
  getLastColumn() { let n = 0; this.values.forEach(row => row.forEach((v, j) => { if (v !== '' && v !== null && v !== undefined) n = Math.max(n, j + 1); })); return n; }
  getMaxRows() { return this.maxRows; } getMaxColumns() { return this.maxCols; }
  getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  protect() { const ex = this.protections[0]; if (ex) return ex; const p = new Protection(this, this.ss.owner); p.proxy = strict(p, 'Protection'); this.protections.push(p.proxy); return p.proxy; }
  getProtections(type) { if (type !== 'SHEET') throw new Error('type de protection inattendu'); return this.protections.slice(); }
  hideSheet() { this.hidden = true; return this.proxy; } showSheet() { this.hidden = false; return this.proxy; } isSheetHidden() { return this.hidden; }
  setFrozenRows(n) { this.frozenR = n; } setFrozenColumns(n) { this.frozenC = n; }
  getFrozenRows() { return this.frozenR; } getFrozenColumns() { return this.frozenC; }
  setHiddenGridlines(b) { this.gridHidden = b; return this.proxy; } hasHiddenGridlines() { return this.gridHidden; }
  setColumnWidth(c, w) { this.widths[c] = w; return this.proxy; }
}
class Spreadsheet {
  constructor(env, id, name, data, owner) {
    this.env = env; this.id = id; this.name = name; this.owner = owner; this.log = []; this.meta = [];
    this.editors = [owner]; this.viewers = []; this.tz = 'Europe/Paris'; this.locale = 'fr_FR'; this.actif = null; this.toasts = [];
    this.sheets = []; this.names = {};
    if (data) {
      for (const [n, d] of Object.entries(data.sheets)) this._add(n, d);
      this.names = data.names;
    }
  }
  _add(n, d, index) { const s = new Sheet(this, n, d); s.proxy = strict(s, 'Sheet'); if (index === undefined) this.sheets.push(s); else this.sheets.splice(index, 0, s); return s; }
  getSheetByName(n) { const s = this.sheets.find(x => x.name === n); return s ? s.proxy : null; }
  getSheets() { return this.sheets.map(s => s.proxy); }
  getRangeByName(n) { const d = this.names[n]; if (!d) return null; const s = this.sheets.find(x => x.name === d.sheet); return s.proxy.getRange(d.ref); }
  getSpreadsheetTimeZone() { return this.tz; } getSpreadsheetLocale() { return this.locale; }
  getId() { return this.id; } getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; } getName() { return this.name; }
  getEditors() { return this.editors.map(user); }
  addEditor(e) { if (!this.editors.includes(e)) this.editors.push(e); return this.proxy; }
  removeEditor(e) { if (e === this.owner) throw new Error('propriétaire non retirable'); this.editors = this.editors.filter(x => x !== e); return this.proxy; }
  addViewer(e) { this.viewers.push(e); return this.proxy; }
  getOwner() { return user(this.owner); }
  toast(m, t, s) { this.toasts.push([t, m, s]); }
  insertSheet(n, i) { if (this.getSheetByName(n)) throw new Error('onglet existant'); const s = this._add(n, null, i); return s.proxy; }
  setActiveSheet(s) { this.actif = [s.getName(), null]; return s; }
  getDeveloperMetadata() { return this.meta.map(([k, v]) => strict({ getKey: () => k, getValue: () => v }, 'DeveloperMetadata')); }
  addDeveloperMetadata(k, v) { this.meta.push([k, v]); return this.proxy; }
}
function creerEnv(json, opts) {
  const env = { ui: { menus: [], dialogs: [], alerts: [], reponses: [], prompts: [] }, mails: [], triggers: [], props: {}, fichiers: {}, dossiers: {} };
  const owner = opts.owner, courant = opts.courant || opts.owner;
  const data = JSON.parse(fs.readFileSync(json, 'utf8'));
  const ss = new Spreadsheet(env, 'ID1', 'KPI_Registre_Incidents_Fundcraft_France', data, owner); ss.proxy = strict(ss, 'Spreadsheet');
  env.ss = ss; env.classeurs = { ID1: ss }; env.actif = ss;
  const racine = { nom: 'Mon Drive', sous: {}, fichiers: {} }; env.racine = racine;
  const dossier = (d) => strict({
    getFoldersByName: n => iter(d.sous[n] ? [dossier(d.sous[n])] : []),
    createFolder: n => { d.sous[n] = { nom: n, sous: {}, fichiers: {} }; return dossier(d.sous[n]); },
    getFilesByName: n => iter(d.fichiers[n] ? [d.fichiers[n]] : []),
    __d: d,
  }, 'Folder');
  const iter = arr => { let i = 0; return strict({ hasNext: () => i < arr.length, next: () => arr[i++] }, 'Iterator'); };
  const fichier = (cl, d) => strict({
    getId: () => cl.id, getUrl: () => cl.proxy.getUrl(),
    getParents: () => iter([dossier(d)]),
    makeCopy: (nom, dest) => {
      const id = 'ID' + (Object.keys(env.classeurs).length + 1);
      const copie = new Spreadsheet(env, id, nom, null, courant);
      cl.sheets.forEach(s => { const n = copie._add(s.name, null); n.values = s.values.map(r => r.slice()); n.bgs = s.bgs; n.fmts = s.fmts; n.hidden = s.hidden; });
      copie.names = cl.names; copie.proxy = strict(copie, 'Spreadsheet'); env.classeurs[id] = copie;
      const f = fichier(copie, dest.__d); dest.__d.fichiers[nom] = f; return f;
    },
    getSharingAccess: () => env.partage[0], getSharingPermission: () => env.partage[1],
    setSharing: (a, p) => { env.partage = [a, p]; }, setShareableByEditors: b => { env.partageEditeurs = b; },
  }, 'File');
  env.partage = ['ANYONE_WITH_LINK', 'EDIT'];
  // makeCopy enregistre la copie dans le dossier de destination
  const g = {
    SpreadsheetApp: strict({
      getActive: () => env.actif.proxy, getActiveSpreadsheet: () => env.actif.proxy,
      getUi: () => strict({
        createMenu: nom => { const m = { nom, items: [] }; const pm = strict({
          addItem: (cap, fn) => { m.items.push([cap, fn]); return pm; }, addSeparator: () => { m.items.push(['---']); return pm; },
          addSubMenu: sm => { m.items.push(['>', sm.__m]); return pm; }, addToUi: () => { env.ui.menus.push(m); }, __m: m }, 'Menu'); return pm; },
        alert: (...a) => { env.ui.alerts.push(a); return env.ui.reponses.length ? env.ui.reponses.shift() : 'OK'; },
        prompt: (...a) => { env.ui.alerts.push(a); const r = env.ui.prompts.shift(); return strict({ getSelectedButton: () => r[0], getResponseText: () => r[1] }, 'PromptResponse'); },
        showModelessDialog: (h, t) => env.ui.dialogs.push(['modeless', t, h.getContent()]),
        showModalDialog: (h, t) => env.ui.dialogs.push(['modal', t, h.getContent()]),
        ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' }, Button: { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' },
      }, 'Ui'),
      openById: id => env.classeurs[id].proxy, flush: () => {},
      ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' }, CopyPasteType: { PASTE_VALUES: 'PASTE_VALUES' },
    }, 'SpreadsheetApp'),
    Session: strict({ getEffectiveUser: () => user(courant), getActiveUser: () => user(courant), getScriptTimeZone: () => opts.tzScript || 'Europe/Paris' }, 'Session'),
    Utilities: strict({ formatDate }, 'Utilities'),
    PropertiesService: strict({ getScriptProperties: () => strict({ getProperty: k => (k in env.props ? env.props[k] : null), setProperty: (k, v) => { env.props[k] = String(v); }, deleteProperty: k => { delete env.props[k]; } }, 'Properties') }, 'PropertiesService'),
    LockService: strict({ getDocumentLock: () => strict({ tryLock: () => true, releaseLock: () => { env.lockReleased = true; } }, 'Lock') }, 'LockService'),
    HtmlService: strict({ createHtmlOutput: html => { const o = strict({ setWidth: () => o, setHeight: () => o, getContent: () => html }, 'HtmlOutput'); return o; } }, 'HtmlService'),
    MailApp: strict({ sendEmail: m => { for (const k of Object.keys(m)) if (!['to', 'subject', 'htmlBody', 'body', 'name'].includes(k)) throw new Error('option sendEmail inconnue ' + k); env.mails.push(m); } }, 'MailApp'),
    ScriptApp: strict({
      newTrigger: fn => strict({
        forSpreadsheet: s => strict({ onEdit: () => strict({ create: () => { env.triggers.push(['onEdit', fn]); } }, 'SpreadsheetTriggerBuilder') }, 'TriggerBuilder'),
        timeBased: () => strict({
          everyDays: n => strict({ atHour: h => strict({ inTimezone: tz => strict({ create: () => { env.triggers.push(['jour', fn, h, tz]); } }, 'ClockTriggerBuilder') }, 'ClockTriggerBuilder') }, 'ClockTriggerBuilder'),
          everyHours: n => strict({ create: () => { env.triggers.push(['heure', fn, n]); } }, 'ClockTriggerBuilder'),
        }, 'ClockTriggerBuilder'),
      }, 'TriggerBuilder'),
      getProjectTriggers: () => env.triggers.map(t => strict({ getHandlerFunction: () => t[1], __t: t }, 'Trigger')),
      deleteTrigger: t => { env.triggers = env.triggers.filter(x => x !== t.__t); },
    }, 'ScriptApp'),
    DriveApp: strict({
      getFileById: id => fichier(env.classeurs[id], racine), getRootFolder: () => dossier(racine),
      Access: { PRIVATE: 'PRIVATE', ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', DOMAIN: 'DOMAIN', DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK' },
      Permission: { VIEW: 'VIEW', EDIT: 'EDIT', COMMENT: 'COMMENT' },
    }, 'DriveApp'),
    console,
  };
  return { env, g };
}
module.exports = { creerEnv, strict, formatDate, a1 };
