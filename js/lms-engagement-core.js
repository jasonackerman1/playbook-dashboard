/* L&D dashboard core: parsing, merging, classification, aggregation. Runs in browser, worker and Node. */
(function (root) {
  const low = s => String(s == null ? '' : s).trim().toLowerCase();
  const mi = (y, m) => y * 12 + m - 1;                       // month index
  const miYear = i => Math.floor(i / 12), miMonth = i => (i % 12) + 1;
  const fyOf = i => (miMonth(i) >= 4 ? miYear(i) : miYear(i) - 1); // FY starts April
  const fmOf = i => (miMonth(i) + 8) % 12;                  // 0 = Apr ... 11 = Mar
  const fmToMi = (fy, fm) => mi(fy, 4) + fm;

  function cellVal(c) { return c == null ? null : (typeof c === 'object' && 'v' in c ? c.v : c); }
  function cellText(c) { return c == null ? '' : (typeof c === 'object' ? (c.w != null ? c.w : String(c.v == null ? '' : c.v)) : String(c)); }

  function detect(head) {
    const h = head.map(low);
    if (h.includes('item id') && h.includes('completion date') && h.includes('user id')) return 'lms';
    if (h.includes('content name') && h.some(x => x.startsWith('hours viewed'))) return 'linkedin';
    return null;
  }

  function parseDateMi(v, text) {
    if (v == null || v === '') return null;
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return mi(d.getUTCFullYear(), d.getUTCMonth() + 1);
    }
    if (v instanceof Date) return mi(v.getFullYear(), v.getMonth() + 1);
    const s = String(text || v);
    let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:,?\s+(\d{1,2}):(\d{2})\s*(AM|PM)?)?/i);
    if (m) {
      let y = +m[3]; if (y < 100) y += 2000;
      const mo = +m[1], d = +m[2];
      if (/eastern/i.test(s) && m[4]) {
        // LMS stamps completions at midnight UTC shown in US Eastern (7/8 PM the day before); use the UTC date
        let hr = +m[4]; const pm = (m[6] || '').toUpperCase() === 'PM';
        if (hr === 12) hr = pm ? 12 : 0; else if (pm) hr += 12;
        const dt = new Date(Date.UTC(y, mo - 1, d, hr + 5));
        return mi(dt.getUTCFullYear(), dt.getUTCMonth() + 1);
      }
      return mi(y, mo);
    }
    m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return mi(+m[1], +m[2]);
    return null;
  }

  function isBCA(userId, email, groups) {
    if (/^c\d/i.test(String(userId || '').trim())) return true;
    const e = low(email);
    if (e.endsWith('.ca') || e.endsWith('@itweapons.com')) return true;
    if (groups && /bca employee/i.test(groups)) return true;
    return false;
  }

  const TYPES = { online: 0, vilt: 1, ilt: 2 };
  const OK_STATUS = new Set(['online-complete', 'virtual instructor-led-complete', 'instructor-led-complete']);

  // rows: array of arrays of cells (first row = header). Returns compact records.
  function parseLms(rows) {
    const h = rows[0].map(c => low(cellText(c)));
    const col = n => h.indexOf(n);
    const c = { id: col('item id'), type: col('item type'), title: col('description'), len: h.findIndex(x => x.startsWith('course length')),
      user: col('user id'), email: col('email address'), domain: col('domain'), status: col('completion status'), date: col('completion date') };
    const st = { read: 0, kept: 0, domain: 0, type: 0, status: 0, date: 0 };
    const recs = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const id = cellText(row[c.id]).trim(); const user = cellText(row[c.user]).trim();
      if (!id && !user) continue;
      st.read++;
      if (c.domain >= 0) { const d = low(cellText(row[c.domain])); if (d && !d.startsWith('emp_')) { st.domain++; continue; } }
      const t = TYPES[low(cellText(row[c.type]))]; if (t == null) { st.type++; continue; }
      if (c.status >= 0 && !OK_STATUS.has(low(cellText(row[c.status])))) { st.status++; continue; }
      const m = parseDateMi(cellVal(row[c.date]), cellText(row[c.date])); if (m == null) { st.date++; continue; }
      let len = c.len >= 0 ? parseFloat(cellVal(row[c.len])) : NaN; if (!(len > 0)) len = null;
      recs.push([m, isBCA(user, c.email >= 0 ? cellText(row[c.email]) : '') ? 1 : 0, t, user, id, cellText(row[c.title]).trim(), len]);
      st.kept++;
    }
    return { kind: 'lms', recs, st };
  }

  function hoursViewed(cell) {
    const txt = cellText(cell), v = cellVal(cell);
    const m = txt.match(/^(\d+):(\d{1,2})(?::(\d{1,2}))?$/);
    if (m) return +m[1] + (+m[2]) / 60 + (m[3] ? +m[3] : 0) / 3600;
    if (typeof v === 'number') return v < 1 && /:/.test(txt) ? v * 24 : v;
    const f = parseFloat(txt); return isNaN(f) ? null : f;
  }

  function parseLinkedIn(rows) {
    const h = rows[0].map(c => low(cellText(c)));
    const col = n => h.indexOf(n);
    const c = { name: col('content name'), hours: h.findIndex(x => x.startsWith('hours viewed')), user: col('unique user id'), email: col('email'),
      unit: col('bus/bca'), groups: h.findIndex(x => x.startsWith('groups (at time')), last: h.findIndex(x => x.startsWith('last viewed')),
      done: h.findIndex(x => x.startsWith('completed at')), cid: col('content id'), type: col('content type') };
    const st = { read: 0, kept: 0, date: 0, hours: 0 };
    const recs = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r]; if (!row) continue;
      const name = cellText(row[c.name]).trim(); if (!name) continue;
      st.read++;
      let m = c.last >= 0 ? parseDateMi(cellVal(row[c.last]), cellText(row[c.last])) : null;
      if (m == null && c.done >= 0) m = parseDateMi(cellVal(row[c.done]), cellText(row[c.done]));
      if (m == null) { st.date++; continue; }
      const hv = hoursViewed(row[c.hours]); if (hv == null) { st.hours++; continue; }
      const user = (c.user >= 0 ? cellText(row[c.user]) : '') || (c.email >= 0 ? cellText(row[c.email]) : '');
      let unit = c.unit >= 0 ? low(cellText(row[c.unit])) : '';
      const bca = unit ? unit === 'bca' : isBCA(user, c.email >= 0 ? cellText(row[c.email]) : '', c.groups >= 0 ? cellText(row[c.groups]) : '');
      const cid = 'LI:' + (c.cid >= 0 && cellText(row[c.cid]) ? cellText(row[c.cid]).trim() : name);
      recs.push([m, bca ? 1 : 0, 0, (/@/.test(user) ? user.trim().toLowerCase() : user.trim()), cid, name, Math.round(hv * 10000) / 10000]);
      st.kept++;
    }
    return { kind: 'linkedin', recs, st };
  }

  function parseRows(rows) {
    const kind = rows.length ? detect(rows[0].map(cellText)) : null;
    if (kind === 'lms') return parseLms(rows);
    if (kind === 'linkedin') return parseLinkedIn(rows);
    return null;
  }

  /* ---------- dataset ---------- */
  function emptyDS() { return { v: 1, updatedAt: 0, users: [], courses: [], lms: { m: [], u: [], t: [], p: [], c: [], h: [] }, li: { m: [], u: [], p: [], c: [], h: [] }, log: [] }; }

  function indexer(arr, keyFn) { const map = new Map(); arr.forEach((x, i) => map.set(keyFn(x), i)); return map; }

  // Replace every month the upload covers, keep the rest.
  function merge(ds, parsed, fileName) {
    const recs = parsed.recs; if (!recs.length) return { replaced: [0, 0] };
    let lo = Infinity, hi = -Infinity; for (const r of recs) { if (r[0] < lo) lo = r[0]; if (r[0] > hi) hi = r[0]; }
    const users = indexer(ds.users, x => x), courses = indexer(ds.courses, x => x[0]);
    const tgt = parsed.kind === 'lms' ? ds.lms : ds.li;
    const keys = Object.keys(tgt), keep = [];
    for (let i = 0; i < tgt.m.length; i++) if (tgt.m[i] < lo || tgt.m[i] > hi) keep.push(i);
    const next = {}; keys.forEach(k => next[k] = keep.map(i => tgt[k][i]));
    for (const r of recs) {
      let ui = users.get(r[3]); if (ui == null) { ui = ds.users.length; ds.users.push(r[3]); users.set(r[3], ui); }
      let ci = courses.get(r[4]); if (ci == null) { ci = ds.courses.length; ds.courses.push([r[4], r[5]]); courses.set(r[4], ci); }
      else if (r[5] && ds.courses[ci][1] !== r[5]) ds.courses[ci][1] = r[5];
      next.m.push(r[0]); next.u.push(r[1]); next.p.push(ui); next.c.push(ci); next.h.push(r[6]);
      if (next.t) next.t.push(r[2]);
    }
    if (parsed.kind === 'lms') ds.lms = next; else ds.li = next;
    ds.updatedAt = Date.now();
    ds.log.push({ kind: parsed.kind, file: fileName, from: lo, to: hi, rows: recs.length, at: ds.updatedAt });
    return { replaced: [lo, hi] };
  }

  /* ---------- classification ---------- */
  const COMPLIANCE_KW = /code of conduct|harassment|phishing|phish catcher|information security fundamentals|corporate compliance|compliance first|hipaa|fcpa|foreign corrupt|data privacy|fair competition|risk manag|workplace violence|employee handbook|charter of corporate behavior|think before you click|human firewall|social engineering|security team/i;
  function autoCompliance(id, title) {
    if (/^LINKEDINLEARNING|^LI:/i.test(id)) return false;
    const hasYear = /20\d\d/.test(id) || /\b20\d\d\b/.test(title);
    return hasYear && COMPLIANCE_KW.test(title);
  }
  const DEPTS = ['Technical Training', 'Sales Training', 'Corporate', 'Global Client Services'];
  function inferDept(id) {
    const u = String(id).toUpperCase();
    if (/^(GCS|SERVICENXT)/.test(u)) return 'Global Client Services';
    if (/^MMS/.test(u)) return 'Corporate';
    if (/^T/.test(u)) return 'Technical Training';
    return 'Sales Training';
  }

  // Build per-course lookup arrays given settings + built-in lists.
  function classify(ds, settings, builtin) {
    const n = ds.courses.length, comp = new Uint8Array(n), li = new Uint8Array(n), dept = new Array(n);
    const add = new Set(settings.compAdd || []), rem = new Set(settings.compRemove || []), deptOv = settings.dept || {};
    for (let i = 0; i < n; i++) {
      const [id, title] = ds.courses[i];
      li[i] = /^LINKEDINLEARNING|^LI:/i.test(id) ? 1 : 0;
      let c = builtin.compliance.has(id) || (settings.autoCompliance !== false && autoCompliance(id, title || ''));
      if (add.has(id)) c = true; if (rem.has(id)) c = false;
      comp[i] = c ? 1 : 0;
      dept[i] = deptOv[id] || builtin.dept[id] || inferDept(id);
    }
    return { comp, li, dept };
  }

  // Flatten into the effective record list: LinkedIn export replaces LMS LinkedIn rows for months it covers.
  // Returns columnar arrays with category codes: 0 online, 1 compliance, 2 linkedin, 3 instructor-led; delivery 0 online,1 vilt,2 ilt
  function effective(ds, cls, defaultHours) {
    const liMonths = new Set(ds.li.m);
    const N = ds.lms.m.length + ds.li.m.length;
    const out = { m: new Int32Array(N), u: new Uint8Array(N), d: new Uint8Array(N), cat: new Uint8Array(N), p: new Int32Array(N), c: new Int32Array(N), h: new Float64Array(N), n: 0, liMonths };
    let k = 0;
    const L = ds.lms;
    for (let i = 0; i < L.m.length; i++) {
      const ci = L.c[i], t = L.t[i];
      if (cls.li[ci] && liMonths.has(L.m[i])) continue;
      out.m[k] = L.m[i]; out.u[k] = L.u[i]; out.d[k] = t; out.p[k] = L.p[i]; out.c[k] = ci;
      out.h[k] = L.h[i] == null ? defaultHours : L.h[i];
      out.cat[k] = t > 0 ? 3 : cls.li[ci] ? 2 : cls.comp[ci] ? 1 : 0; k++;
    }
    const Q = ds.li;
    for (let i = 0; i < Q.m.length; i++) {
      out.m[k] = Q.m[i]; out.u[k] = Q.u[i]; out.d[k] = 0; out.p[k] = Q.p[i]; out.c[k] = Q.c[i]; out.h[k] = Q.h[i]; out.cat[k] = 2; k++;
    }
    out.n = k; return out;
  }

  function newAgg() { return { n: 0, h: 0, users: new Set() }; }
  function add(a, E, i) { a.n++; a.h += E.h[i]; a.users.add(E.p[i]); }

  // Month window for a fiscal year: fm 0..through (inclusive); limited by data availability for monthsCount
  function inWindow(m, fy, through) { return fyOf(m) === fy && fmOf(m) <= through; }

  function finish(a, months) {
    return { courses: a.n, hours: a.h, users: a.users.size, perMonth: months ? a.h / months : 0,
      coursesPer: a.users.size ? a.n / a.users.size : 0, hoursPer: a.users.size ? a.h / a.users.size : 0 };
  }

  // The FY Learning Completion breakdown (mirrors the Summary sheet)
  function reportTable(E, cls, fy, through, months) {
    const g = {}; const key = k => (g[k] = g[k] || newAgg());
    const CAT = ['LMS Online Courses', 'Compliance', 'LinkedIn', 'Instructor Led Training'];
    const U = ['BUS', 'BCA'];
    for (let i = 0; i < E.n; i++) {
      if (!inWindow(E.m[i], fy, through)) continue;
      const cat = CAT[E.cat[i]], u = U[E.u[i]];
      add(key(cat), E, i); add(key(cat + '|' + u), E, i); add(key('Total'), E, i); add(key('Total|' + u), E, i);
      if (E.cat[i] === 3) add(key('ILT|' + cls.dept[E.c[i]]), E, i);
    }
    const f = k => finish(g[k] || newAgg(), months);
    const rows = [];
    CAT.forEach((c, ci) => {
      rows.push({ label: c, level: 0, cat: ci, ...f(c) });
      U.forEach(u => rows.push({ label: u, level: 1, cat: ci, unit: u, ...f(c + '|' + u) }));
      if (ci === 3) DEPTS.forEach(d => rows.push({ label: d, level: 1, cat: ci, dept: d, ...f('ILT|' + d) }));
    });
    rows.push({ label: 'Total', level: 0, total: true, ...f('Total') });
    U.forEach(u => rows.push({ label: u, level: 1, total: true, unit: u, ...f('Total|' + u) }));
    return rows;
  }

  // Year totals with optional predicate (e.g. exclude compliance) split by unit
  function yearTotals(E, fy, through, months, pred) {
    const all = newAgg(), bu = [newAgg(), newAgg()];
    for (let i = 0; i < E.n; i++) {
      if (!inWindow(E.m[i], fy, through) || (pred && !pred(i))) continue;
      add(all, E, i); add(bu[E.u[i]], E, i);
    }
    return { all: finish(all, months), BUS: finish(bu[0], months), BCA: finish(bu[1], months) };
  }

  // Reads a workbook/CSV buffer, finds recognised sheets (first LMS sheet, first LinkedIn sheet), parses them.
  function readBook(XLSX, buf, progress) {
    progress && progress('Looking for completion data in the file');
    const head = XLSX.read(buf, { type: 'array', sheetRows: 2, dense: true });
    const pick = {};
    for (const n of head.SheetNames) {
      const ws = head.Sheets[n]; const rows = ws['!data'] || ws; if (!rows || !rows[0]) continue;
      const k = detect(rows[0].map(cellText)); if (k && !pick[k]) pick[k] = n;
    }
    const names = Object.values(pick);
    if (!names.length) return [];
    progress && progress('Reading ' + names.map(n => '“' + n + '”').join(' and '));
    const wb = XLSX.read(buf, { type: 'array', sheets: names, dense: true });
    return names.map(n => { const ws = wb.Sheets[n]; return { sheet: n, parsed: parseRows(ws['!data'] || ws) }; });
  }

  root.LDCore = { readBook, inWindow, finish, reportTable, yearTotals, mi, miYear, miMonth, fyOf, fmOf, fmToMi, detect, parseRows, parseLms, parseLinkedIn, parseDateMi, emptyDS, merge, classify, effective, autoCompliance, inferDept, DEPTS, newAgg, add, cellText };
})(typeof self !== 'undefined' ? self : globalThis);
