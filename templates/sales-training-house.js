/* House-style additions for the Sales Training & L&D Engagement dashboard (inlined into
   sales-training.html by update_sales_training_dashboard.py). Resmie's own page logic stays as
   she wrote it in the script above; this only adds the shared header behavior:
   theme (pb-theme), triple-click home, Show/Hide TLG, "?" info popovers, and Export. */

// ── Theme (shared pb-theme key; light unless the user picked dark) ──
(function () {
  if (localStorage.getItem('pb-theme') !== 'dark') document.body.classList.add('light-mode');
  setThemeLabel();
})();
function setThemeLabel() {
  document.getElementById('themeBtn').textContent = document.body.classList.contains('light-mode') ? '🌙 Dark' : '☀ Light';
}
function toggleTheme() {
  const light = document.body.classList.toggle('light-mode');
  try { localStorage.setItem('pb-theme', light ? 'light' : 'dark'); } catch (e) {}
  setThemeLabel(); render();
}

// ── Home: triple-click the title ──
(function () {
  const h1 = document.getElementById('dash-title'); let n = 0, t = null;
  h1.addEventListener('click', function () {
    n++; clearTimeout(t);
    if (n >= 3) { n = 0; window.location.href = 'index.html'; }
    else t = setTimeout(function () { n = 0; }, 1500);
  });
})();

// ── TLG: hidden by default ──
function toggleTLG() {
  state.tlg = !state.tlg;
  const b = document.getElementById('btn-tlg');
  b.classList.toggle('active', !state.tlg);
  b.textContent = state.tlg ? 'Hide TLG' : 'Show TLG';
  render();
}

// ── Info popovers ──
const INFO = {
  'export': 'Download what is on screen now: the selected tab, filters and TLG setting. PDF prints this page; Excel gives every section as its own sheet, with the full lists (all categories, all courses, every branch).',
  'tlg': 'The TLG (internal L&D team) is hidden from every number, chart and export by default. Click Show TLG to include them.',
  'filters': 'Every section follows the tab and these filters, except Upcoming live classes, which follows the tab, its first filter and Rep status only. Rep status: new hire = completion within 65 days of the hire date. Dealers have no usable hire date, so they appear only under All reps. Period: fiscal year to date (April 1 onward) or a single quarter.',
  'trend': 'Completions in each month of the period, split into instructor-led (ILT + VILT) and online. The current month is month to date (MTD).',
  'cats': 'The 10 catalog categories with the most completions. A course can belong to several categories and counts toward each, so these add up to more than total completions.',
  'ilt': 'Every instructor-led (ILT or VILT) course with a completion, ranked by completions. * = completions recorded as Complete ILT (taken in person).',
  'cmp': 'New hires (completion within 65 days of the hire date) compared with tenured reps. Reach shows on the Sales Org tab only: roster reps with a completion in view, divided by the roster. Empty on the Dealers tab, since dealer hire dates are placeholders.',
  'upcoming': 'Live classes starting on or after the report date, from the future Registration Status report. Each person counts once per class; waitlisted people count as registered and cancelled registrations are left out. Sections of the same class on the same date are combined.',
  'breakdown': 'People, completions, online, instructor-led and completions per person by group. The first view is business unit (All Groups), sales field (Sales Org) or country (Dealers); the second is branch, or dealer on the Dealers tab. Click a column to sort.',
  'top': 'The ten courses with the most completions, with each course’s main catalog category and format.'
};
function showInfo(e, key) {
  e.stopPropagation();
  const p = document.getElementById('info-popover');
  if (p.classList.contains('visible') && p.dataset.key === key) { p.classList.remove('visible'); return; }
  p.textContent = INFO[key] || ''; p.dataset.key = key; p.classList.add('visible');
  const r = e.target.getBoundingClientRect();
  p.style.left = Math.max(8, Math.min(r.left, innerWidth - p.offsetWidth - 8)) + 'px';
  p.style.top = (r.bottom + 8 + p.offsetHeight > innerHeight ? r.top - p.offsetHeight - 8 : r.bottom + 8) + 'px';
}
function toggleExportDrop(e) { e.stopPropagation(); document.getElementById('export-menu').classList.toggle('open'); }
document.addEventListener('click', function () {
  document.getElementById('info-popover').classList.remove('visible');
  document.getElementById('export-menu').classList.remove('open');
});

// ── Export ──
function scopeText() { return document.getElementById('scope').textContent.replace(/^Showing\s*/, '') + (state.tlg ? ', TLG shown' : ', TLG hidden'); }
function runExportPDF() {
  document.getElementById('export-menu').classList.remove('open');
  document.getElementById('ph-date').textContent = 'Data through ' + document.getElementById('asof').textContent + ' · ' + scopeText();
  window.print();
}
function runExportXLSX() {
  document.getElementById('export-menu').classList.remove('open');
  const rows = baseRows(), tab = TABS.find(t => t.id === state.tab), a = agg(rows), wb = XLSX.utils.book_new();
  const add = (name, aoa, widths) => { const ws = XLSX.utils.aoa_to_sheet(aoa); if (widths) ws['!cols'] = widths.map(w => ({ wch: w })); XLSX.utils.book_append_sheet(wb, ws, name); };
  const head = [['Sales Training & L&D Engagement'], ['Data through ' + document.getElementById('asof').textContent], ['Showing: ' + scopeText()], []];
  let iltN = 0; rows.forEach(i => { if (C.i[i]) iltN++; });
  const r = reachOf(rows);
  add('Summary', head.concat([['Measure', 'Value'],
    ['People who completed training', a.learners], ['Completions', a.k], ['Online completions', a.k - a.il], ['Instructor-led completions', a.il],
    ['  recorded as Complete ILT (in person)', iltN], ['Instructor-led share of completions', pct(a.il, a.k) / 100]]
    .concat(r ? [['Reach (' + r.label + ' roster)', r.pct / 100], ['Roster reps with a completion', r.done], ['Reps on the roster', r.n]] : [])), [42, 16]);
  // by month
  const per = PERIODS[state.year], months = [];
  for (let ym = per[1]; ym <= per[2]; ym++) {
    let il = 0, on = 0; rows.forEach(i => { if (YM[i] === ym) C.v[i] ? il++ : on++; });
    months.push([MON[ym % 12] + ' ' + Math.floor(ym / 12) + (ym === asofYM ? ' (MTD)' : ''), il, on, il + on]);
  }
  add('By Month', head.concat([['Month', 'Instructor-led', 'Online', 'Total']], months), [16, 14, 12, 12]);
  // categories (all)
  const cm = D.cats.map(() => [0, 0]); rows.forEach(i => D.courses[C.c[i]][2].forEach(c => cm[c][C.v[i] ? 0 : 1]++));
  add('Categories', head.concat([['Category', 'Instructor-led', 'Online', 'Total']],
    D.cats.map((c, j) => [c, cm[j][0], cm[j][1], cm[j][0] + cm[j][1]]).filter(x => x[3]).sort((x, y) => y[3] - x[3])), [44, 14, 12, 12]);
  // courses (all, ranked)
  const g = new Map(); rows.forEach(i => { let o = g.get(C.c[i]); if (!o) { o = { k: 0, L: new Set(), ilt: 0 }; g.set(C.c[i], o); } o.k++; o.L.add(C.e[i]); if (C.i[i]) o.ilt++; });
  const crs = [...g].map(([c, o]) => { const cr = D.courses[c]; return [cr[0], cr[1], cr[2].map(j => D.cats[j]).join(', '), cr[3] ? 'Instructor-led' : 'Online', o.k, o.L.size, o.ilt]; })
    .sort((x, y) => y[4] - x[4] || x[1].localeCompare(y[1]));
  add('Courses', head.concat([['Course ID', 'Course', 'Categories', 'Format', 'Completions', 'People', 'Complete ILT (in person)']], crs), [18, 52, 40, 14, 12, 10, 14]);
  // new hires vs tenured
  if (state.tab !== 'dlr') {
    const base = baseRows('status'), col = [1, 2].map(s => { const rs = base.filter(i => C.s[i] === s); return { a: agg(rs), r: reachOf(rs, String(s)) }; });
    const line = (l, f) => [l].concat(col.map(f));
    add('New Hires vs Tenured', head.concat([['', 'New hires', 'Tenured'],
      line('People', x => x.a.learners), line('Completions', x => x.a.k), line('Courses per person', x => x.a.learners ? +(x.a.k / x.a.learners).toFixed(1) : ''),
      line('Online', x => x.a.k - x.a.il), line('Instructor-led', x => x.a.il)].concat(col.some(x => x.r) ? [line('Reach', x => x.r ? x.r.pct / 100 : '')] : [])), [22, 12, 12]);
  }
  // upcoming classes
  const ug = new Map(); D.upcoming.filter(upFilter).forEach(u => {
    const key = u[2] + '|' + u[1]; let o = ug.get(key); if (!o) { o = [u[2], u[1], u[3], 0, 0, 0, 0]; ug.set(key, o); }
    if (u[7] === 'Cancelled') return; o[3]++; if (u[4] === 'Dealers') o[6]++; else if (u[6] === 'New hire') o[4]++; else if (u[6] === 'Tenured') o[5]++;
  });
  add('Upcoming Classes', head.concat([['Date', 'Class', 'Instructors', 'Registered', 'New hires', 'Tenured', 'Dealers']],
    [...ug.values()].filter(o => o[3] > 0).sort((x, y) => x[0].localeCompare(y[0]) || x[1].localeCompare(y[1]))), [12, 50, 30, 11, 10, 10, 10]);
  // participation breakdown, both views
  const bd = (key, names, label) => {
    const m = new Map(); rows.forEach(i => { const v = C[key][i]; let o = m.get(v); if (!o) { o = { L: new Set(), k: 0, il: 0 }; m.set(v, o); } o.L.add(C.e[i]); o.k++; if (C.v[i]) o.il++; });
    return [[label, 'People', 'Completions', 'Online', 'Instructor-led', 'Per person']].concat(
      [...m].map(([v, o]) => [names[v], o.L.size, o.k, o.k - o.il, o.il, +(o.k / o.L.size).toFixed(1)]).sort((x, y) => y[2] - x[2]));
  };
  const first = state.tab === 'org' ? bd('o', ['', 'SF (US)', 'SFC (Canada)'], 'Sales field')
    : state.tab === 'dlr' ? bd('dc', ['', 'US', 'BCD (Canada)'], 'Country') : bd('b', D.bus.map(b => BU_LABEL[b]), 'Business unit');
  add(state.tab === 'org' ? 'By Sales Field' : state.tab === 'dlr' ? 'By Country' : 'By Business Unit', head.concat(first), [44, 10, 12, 10, 14, 10]);
  add(state.tab === 'dlr' ? 'By Dealer' : 'By Branch', head.concat(bd('br', D.branches, state.tab === 'dlr' ? 'Dealer' : 'Branch')), [44, 10, 12, 10, 14, 10]);
  XLSX.writeFile(wb, 'sales-training-' + tab.id + '-' + D.asof + '.xlsx');
}
