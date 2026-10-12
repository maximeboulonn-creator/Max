  const isoOf = d => d.toISOString().slice(0, 10);
  function parseISO(s) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '').trim()); if (!m) return null; const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); return isNaN(d.getTime()) ? null : d; }
  const addDaysD = (d, n) => new Date(d.getTime() + n * 86400000);
  function easterUTC(y) { const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451), mo = Math.floor((h + l - 7 * m + 114) / 31), da = ((h + l - 7 * m + 114) % 31) + 1; return new Date(Date.UTC(y, mo - 1, da)); }
  const HOLS = {};
  function holidays(y) {
    if (HOLS[y]) return HOLS[y];
    const s = new Map(); const add = (d, n) => s.set(isoOf(d), n); const e = easterUTC(y);
    [[1, 1, 'Jour de l\'an'], [5, 1, 'Fête du Travail'], [5, 8, 'Victoire 1945'], [7, 14, 'Fête nationale'], [8, 15, 'Assomption'], [11, 1, 'Toussaint'], [11, 11, 'Armistice'], [12, 25, 'Noël']].forEach(([m, d, n]) => add(new Date(Date.UTC(y, m - 1, d)), n));
    add(addDaysD(e, 1), 'Lundi de Pâques'); add(addDaysD(e, 39), 'Ascension'); add(addDaysD(e, 50), 'Lundi de Pentecôte');
    HOLS[y] = s; return s;
  }
  function addMonths(d, n) { const y = d.getUTCFullYear(), m = d.getUTCMonth() + n, day = d.getUTCDate(); const first = new Date(Date.UTC(y, m, 1)); const last = new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth() + 1, 0)).getUTCDate(); return new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth(), Math.min(day, last))); }
