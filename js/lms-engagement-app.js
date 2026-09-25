
const $ = id => document.getElementById(id);
const C = LDCore;
const MON = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
const MONFULL = ['April','May','June','July','August','September','October','November','December','January','February','March'];
const CATNAME = ['Online course', 'Compliance', 'LinkedIn Learning', 'Instructor-led'];
const DELNAME = ['Online', 'VILT', 'ILT'];
const SKEY = 'ld-dash-settings-v1';

const BUILTIN = JSON.parse($('builtin').textContent);
const builtin = { compliance: new Set(BUILTIN.compliance), dept: BUILTIN.dept };
let DS = JSON.parse($('dataset').textContent);
let settings = Object.assign({ defaultMin: 30, autoCompliance: true, compAdd: [], compRemove: [], dept: {} },
  JSON.parse($('settings').textContent || '{}'), loadLocalSettings());

const view = { fy: null, through: 11, unit: 'all', delivery: 'all', comp: 'include', li: 'include', metric: 'hours' };
let CLS, E, lastMi, courseInfo;

/* ---------- formatting ---------- */
const nf0 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
const nf1 = new Intl.NumberFormat('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const nf2 = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const n0 = v => nf0.format(Math.round(v)), n1 = v => nf1.format(v), n2 = v => nf2.format(v);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pct = (a, b) => (b ? (a - b) / b : null);
function pctTxt(p) { if (p == null || !isFinite(p)) return '–'; const v = p * 100; return (v >= 0.05 ? '▲ ' : v <= -0.05 ? '▼ ' : '') + nf1.format(Math.abs(v)) + '%'; }
function pctWords(p) { if (p == null || !isFinite(p)) return ''; const v = p * 100; return Math.abs(v) < 0.05 ? 'flat' : (v > 0 ? 'up ' : 'down ') + nf1.format(Math.abs(v)) + '%'; }
const monthName = m => ['January','February','March','April','May','June','July','August','September','October','November','December'][C.miMonth(m) - 1] + ' ' + C.miYear(m);
const fyLabel = fy => 'FY' + fy;
function windowLabel(fy, through) { return through === 11 ? 'April ' + fy + ' – March ' + (fy + 1) : 'April – ' + MONFULL[through] + ' ' + (through >= 9 ? fy + 1 : fy); }

/* ---------- storage ---------- */
function loadLocalSettings() { try { return JSON.parse(localStorage.getItem(SKEY) || '{}'); } catch (e) { return {}; } }
function saveSettings() { try { localStorage.setItem(SKEY, JSON.stringify(settings)); } catch (e) {} }


/* ---------- derived data ---------- */
function rebuild() {
  CLS = C.classify(DS, settings, builtin);
  E = C.effective(DS, CLS, settings.defaultMin / 60);
  lastMi = -1; for (let i = 0; i < E.n; i++) if (E.m[i] > lastMi) lastMi = E.m[i];
  const n = DS.courses.length;
  courseInfo = { count: new Int32Array(n), instr: new Uint8Array(n) };
  for (let i = 0; i < DS.lms.c.length; i++) { const c = DS.lms.c[i]; courseInfo.count[c]++; if (DS.lms.t[i] > 0) courseInfo.instr[c] = 1; }
  for (let i = 0; i < DS.li.c.length; i++) courseInfo.count[DS.li.c[i]]++;
}
function fyList() { let lo = Infinity, hi = -Infinity; for (let i = 0; i < E.n; i++) { const f = C.fyOf(E.m[i]); if (f < lo) lo = f; if (f > hi) hi = f; } const out = []; for (let f = lo; f <= hi; f++) out.push(f); return out; }
function lastFm(fy) { return C.fyOf(lastMi) === fy ? C.fmOf(lastMi) : fy < C.fyOf(lastMi) ? 11 : -1; }
function monthsFor(fy, through) { return Math.max(1, Math.min(through, lastFm(fy)) + 1); }

function predicate() {
  const { unit, delivery, comp, li } = view;
  return i => {
    if (unit !== 'all' && E.u[i] !== +unit) return false;
    const d = E.d[i];
    if (delivery === 'online' && d !== 0) return false;
    if (delivery === 'instructor' && d === 0) return false;
    if (delivery === 'vilt' && d !== 1) return false;
    if (delivery === 'ilt' && d !== 2) return false;
    const c = E.cat[i];
    if (comp === 'exclude' && c === 1) return false;
    if (comp === 'only' && c !== 1) return false;
    if (li === 'exclude' && c === 2) return false;
    if (li === 'only' && c !== 2) return false;
    return true;
  };
}
function filterPhrase() {
  const parts = [];
  if (view.delivery === 'online') parts.push('online'); else if (view.delivery === 'instructor') parts.push('instructor-led (VILT and ILT)');
  else if (view.delivery === 'vilt') parts.push('VILT'); else if (view.delivery === 'ilt') parts.push('in-person ILT');
  if (view.comp === 'only') parts.push('compliance'); if (view.li === 'only') parts.push('LinkedIn Learning');
  let s = (parts.length ? parts.join(' ') : 'all') + ' training';
  const ex = []; if (view.comp === 'exclude') ex.push('compliance'); if (view.li === 'exclude') ex.push('LinkedIn Learning');
  if (ex.length) s += ', excluding ' + ex.join(' and ');
  return s;
}

/* ---------- controls ---------- */
const SEGS = {
  unit: [['all', 'Both'], ['0', 'BUS'], ['1', 'BCA']],
  delivery: [['all', 'All'], ['online', 'Online'], ['instructor', 'VILT + ILT'], ['vilt', 'VILT'], ['ilt', 'ILT']],
  comp: [['include', 'Include'], ['exclude', 'Exclude'], ['only', 'Only']],
  li: [['include', 'Include'], ['exclude', 'Exclude'], ['only', 'Only']],
  metric: [['hours', 'Hours'], ['courses', 'Completions'], ['users', 'Learners']]
};
function buildSegs() {
  document.querySelectorAll('.seg[data-key]').forEach(el => {
    const key = el.dataset.key;
    el.setAttribute('role', 'radiogroup');
    el.innerHTML = SEGS[key].map(([v, l]) => `<label><input type="radio" name="seg-${key}" value="${v}"${view[key] === v ? ' checked' : ''}><span>${l}</span></label>`).join('');
    el.onchange = e => { view[key] = e.target.value; render(); };
  });
  // Same SEGS data, rendered as a <select> instead of a button group - used for
  // Delivery, which has enough options (5) that a dropdown is more compact.
  document.querySelectorAll('select[data-key]').forEach(el => {
    const key = el.dataset.key;
    el.innerHTML = SEGS[key].map(([v, l]) => `<option value="${v}"${view[key] === v ? ' selected' : ''}>${l}</option>`).join('');
    el.onchange = e => { view[key] = e.target.value; render(); };
  });
}
function buildPeriod() {
  const fys = fyList();
  if (view.fy == null || !fys.includes(view.fy)) view.fy = fys[fys.length - 1];
  $('fy').innerHTML = fys.slice().reverse().map(f => `<option value="${f}"${f === view.fy ? ' selected' : ''}>${fyLabel(f)} (Apr ${f} – Mar ${f + 1})</option>`).join('');
  $('through').innerHTML = `<option value="11"${view.through === 11 ? ' selected' : ''}>Full year</option>` +
    MON.slice(0, 11).map((m, i) => `<option value="${i}"${view.through === i ? ' selected' : ''}>April – ${MONFULL[i]}</option>`).join('');
}
$('fy').onchange = e => { view.fy = +e.target.value; render(); };
$('through').onchange = e => { view.through = +e.target.value; render(); };

/* ---------- render ---------- */
function render() {
  renderStatus(); renderLede(); renderBanner(); renderMonthly(); renderYoyFiltered(); renderReport(); renderHistory(); renderTop(); renderCoverage();
}

function renderStatus() {
  const last = DS.log.length ? DS.log[DS.log.length - 1] : null;
  const dateLabel = last
    ? new Date(last.at).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
    : (lastMi >= 0 ? monthName(lastMi) : null);
  $('status').textContent = dateLabel
    ? `Data through ${dateLabel}. ${n0(E.n)} employee completions loaded.`
    : 'No data loaded yet.';
}

function sumWindow(fy, through, pred) {
  const a = [C.newAgg(), C.newAgg()], all = C.newAgg();
  for (let i = 0; i < E.n; i++) { if (!C.inWindow(E.m[i], fy, through) || !pred(i)) continue; C.add(a[E.u[i]], E, i); C.add(all, E, i); }
  const m = monthsFor(fy, through);
  return { BUS: C.finish(a[0], m), BCA: C.finish(a[1], m), all: C.finish(all, m) };
}

function renderLede() {
  const fy = view.fy, pred = predicate();
  const eff = Math.min(view.through, lastFm(fy));                    // like-for-like window
  const cur = sumWindow(fy, eff, pred), prev = fyList().includes(fy - 1) ? sumWindow(fy - 1, eff, pred) : null;
  const unitTxt = view.unit === 'all' ? 'employees' : view.unit === '0' ? 'BUS employees' : 'BCA employees';
  if (!cur.all.courses) { $('lede').textContent = `No ${filterPhrase()} completions for ${unitTxt} in ${fyLabel(fy)} with these filters.`; return; }
  const cmp = (a, b) => b ? ` (${pctWords(pct(a, b))} on ${fyLabel(fy - 1)})` : '';
  const scope = filterPhrase();
  let s = `${scope[0].toUpperCase() + scope.slice(1)}, ${fyLabel(fy)} (${windowLabel(fy, eff)}): ${unitTxt} completed <strong>${n0(cur.all.courses)}</strong> courses for <strong>${n0(cur.all.hours)} hours</strong>${prev ? cmp(cur.all.hours, prev.all.hours) : ''}, with ${n0(cur.all.users)} people taking part.`;
  if (view.unit === 'all') {
    s += ` <strong class="bus">BUS</strong> logged ${n0(cur.BUS.hours)} hours${prev ? cmp(cur.BUS.hours, prev.BUS.hours) : ''} and <strong class="bca">BCA</strong> ${n0(cur.BCA.hours)}${prev ? cmp(cur.BCA.hours, prev.BCA.hours) : ''}.`;
  }
  if (prev && eff < 11) s += ` Comparisons use the same months of ${fyLabel(fy - 1)}.`;
  $('lede').innerHTML = s;
}

function renderBanner() {
  const fy = view.fy, lf = lastFm(fy), b = $('banner');
  if (lf < 11 && view.through === 11 && lf >= 0) {
    b.innerHTML = `<p class="banner">${fyLabel(fy)} only has data through ${MONFULL[lf]}, so full-year tables will look low next to earlier years. <button class="btn link" type="button" id="likeBtn">Compare April – ${MONFULL[lf]} for every year</button></p>`;
    $('likeBtn').onclick = () => { view.through = lf; buildPeriod(); render(); };
  } else if (view.through < 11) {
    b.innerHTML = `<p class="banner">Every year is counted April through ${MONFULL[view.through]} only. <button class="btn link" type="button" id="fullBtn">Show full years</button></p>`;
    $('fullBtn').onclick = () => { view.through = 11; buildPeriod(); render(); };
  } else b.innerHTML = '';
}

function monthlySeries(fy, pred) {
  const out = Array.from({ length: 12 }, () => [C.newAgg(), C.newAgg(), C.newAgg()]);
  for (let i = 0; i < E.n; i++) {
    if (C.fyOf(E.m[i]) !== fy || !pred(i)) continue;
    const r = out[C.fmOf(E.m[i])]; C.add(r[E.u[i]], E, i); C.add(r[2], E, i);
  }
  return out;
}
function metricOf(a) { return view.metric === 'hours' ? a.h : view.metric === 'courses' ? a.n : a.users.size; }

function renderMonthly() {
  const fy = view.fy, pred = predicate();
  const cur = monthlySeries(fy, pred), hasPrev = fyList().includes(fy - 1), prev = hasPrev ? monthlySeries(fy - 1, pred) : null;
  const lf = lastFm(fy);
  const label = view.metric === 'hours' ? 'Training hours' : view.metric === 'courses' ? 'Course completions' : 'People who completed something';
  $('chartTitle').textContent = `${label} by month, ${fyLabel(fy)}`;
  $('chartSub').textContent = view.metric === 'users' ? 'Learners are counted once per month per unit.' : 'Bars are stacked BUS then BCA.';
  $('priorLegend').style.display = hasPrev ? '' : 'none';
  const W = 760, H = 300, L = 54, R = 12, T = 16, B = 34, cw = (W - L - R) / 12, bw = cw * 0.56;
  const vals = cur.map((r, i) => i <= lf ? [metricOf(r[0]), metricOf(r[1])] : null);
  const pv = prev ? prev.map(r => metricOf(r[2])) : [];
  const max = Math.max(1, ...vals.filter(Boolean).map(v => v[0] + v[1]), ...pv);
  const nice = niceMax(max), y = v => T + (H - T - B) * (1 - v / nice);
  let g = '';
  for (let k = 0; k <= 4; k++) {
    const v = nice * k / 4, yy = y(v);
    g += `<line x1="${L}" x2="${W - R}" y1="${yy}" y2="${yy}" stroke="var(--rule)" stroke-width="1"/><text x="${L - 8}" y="${yy + 4}" text-anchor="end" font-size="12" fill="var(--muted)">${n0(v)}</text>`;
  }
  vals.forEach((v, i) => {
    const x = L + i * cw + (cw - bw) / 2;
    g += `<text x="${L + i * cw + cw / 2}" y="${H - 12}" text-anchor="middle" font-size="12.5" fill="var(--muted)">${MON[i]}</text>`;
    if (!v) return;
    const yb = y(v[0]), yc = y(v[0] + v[1]);
    g += `<rect x="${x}" y="${yb}" width="${bw}" height="${Math.max(0, y(0) - yb)}" fill="var(--bus)"><title>${MON[i]} BUS: ${n0(v[0])}</title></rect>`;
    g += `<rect x="${x}" y="${yc}" width="${bw}" height="${Math.max(0, yb - yc)}" fill="var(--bca)"><title>${MON[i]} BCA: ${n0(v[1])}</title></rect>`;
  });
  if (prev) {
    const pts = pv.map((v, i) => [L + i * cw + cw / 2, y(v)]);
    g += `<polyline points="${pts.map(p => p.join(',')).join(' ')}" fill="none" stroke="var(--prior)" stroke-width="1.75" stroke-dasharray="5 4"/>`;
    pts.forEach((p, i) => g += `<circle cx="${p[0]}" cy="${p[1]}" r="3.5" fill="var(--panel)" stroke="var(--prior)" stroke-width="1.75"><title>${MON[i]} ${fyLabel(fy - 1)}: ${n0(pv[i])}</title></circle>`);
  }
  $('monthly').innerHTML = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(label)} by month" font-family="Instrument Sans, system-ui, sans-serif">${g}</svg>`;
}
function niceMax(v) { const p = Math.pow(10, Math.floor(Math.log10(v))); for (const m of [1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) if (m * p >= v) return m * p; return v; }

function renderYoyFiltered() {
  const pred = predicate(), fys = fyList(), th = view.through;
  $('yoySub').textContent = `${filterPhrase()[0].toUpperCase() + filterPhrase().slice(1)}, ${th === 11 ? 'full fiscal years' : 'April through ' + MONFULL[th] + ' of each year'}. Change compares with the year before.`;
  let h = `<thead><tr><th>Fiscal year</th><th class="bus">BUS hours</th><th>Change</th><th class="bca">BCA hours</th><th>Change</th><th>Total hours</th><th>Change</th><th>Completions</th><th>Learners</th><th>Hours per learner</th></tr></thead><tbody>`;
  let prev = null;
  for (const fy of fys) {
    const s = sumWindow(fy, th, pred), part = lastFm(fy) < th;
    h += `<tr><td>${fyLabel(fy)}${part ? `<span class="tag">through ${MON[lastFm(fy)]}</span>` : ''}</td>
      <td>${n0(s.BUS.hours)}</td><td class="dim">${prev ? pctTxt(pct(s.BUS.hours, prev.BUS.hours)) : '–'}</td>
      <td>${n0(s.BCA.hours)}</td><td class="dim">${prev ? pctTxt(pct(s.BCA.hours, prev.BCA.hours)) : '–'}</td>
      <td>${n0(s.all.hours)}</td><td class="dim">${prev ? pctTxt(pct(s.all.hours, prev.all.hours)) : '–'}</td>
      <td>${n0(s.all.courses)}</td><td>${n0(s.all.users)}</td><td>${n1(s.all.hoursPer)}</td></tr>`;
    prev = s;
  }
  $('yoyFiltered').innerHTML = h + '</tbody>';
}

const RCOLS = `<thead><tr><th>Learning source</th><th># of courses</th><th>Hours</th><th>Unique users</th><th>Avg. hours per month</th><th>Avg. # of courses per person</th><th>Avg. # of hours per person</th></tr></thead>`;
function rowCells(r) { return `<td>${n0(r.courses)}</td><td>${n0(r.hours)}</td><td>${n0(r.users)}</td><td>${n1(r.perMonth)}</td><td>${n2(r.coursesPer)}</td><td>${n2(r.hoursPer)}</td>`; }

function renderReport() {
  const fy = view.fy, th = view.through, months = monthsFor(fy, th);
  $('reportTitle').textContent = `${fyLabel(fy)} learning completion summary`;
  const rows = C.reportTable(E, CLS, fy, th, months);
  let h = RCOLS + '<tbody>';
  for (const r of rows) {
    if (r.dept && r.label === 'Technical Training') h += `<tr class="l1"><td colspan="7" class="dim" style="padding-top:12px;font-size:13.5px">Instructor-led by department</td></tr>`;
    const cls = r.level === 0 ? 'l0' + (r.total ? ' total' : ' group') : 'l1';
    const lbl = r.level === 0 ? esc(r.label) : r.unit ? `<span class="${r.unit.toLowerCase()}">${r.unit}</span>` : esc(r.label.replace(' Training', '').replace('Global Client Services', 'GCS'));
    h += `<tr class="${cls}"><td>${lbl}</td>${rowCells(r)}</tr>`;
  }
  $('report').innerHTML = h + '</tbody>';
  const liM = [...E.liMonths].filter(m => C.fyOf(m) === fy && C.fmOf(m) <= th).length, used = Math.min(th, lastFm(fy)) + 1;
  const liNote = liM >= used ? 'LinkedIn hours are actual hours viewed from the LinkedIn Learning export.'
    : liM ? `LinkedIn hours use the LinkedIn export for ${liM} of ${used} months; the rest use the LMS record at ${settings.defaultMin} minutes per course.`
    : `LinkedIn hours come from the LMS file at ${settings.defaultMin} minutes per course. Add a LinkedIn Learning export for actual hours viewed.`;
  $('reportFoot').textContent = `Average hours per month divides by ${months} month${months > 1 ? 's' : ''}. ${liNote} Unique users in the total are counted once even if they appear in several sources.`;
}

function histRows(key) { return BUILTIN.history[key] || []; }
function yoyTable(el, key, pred) {
  const fys = fyList(), th = view.through, full = th === 11;
  let h = `<thead><tr><th>Fiscal year</th><th>Courses</th><th>Avg. hours per month</th><th>Hours of training</th><th>Unique users</th><th>Avg. # of courses per person</th><th>Avg. # of hours per person</th></tr></thead><tbody>`;
  const seq = [];
  if (full) for (const r of histRows(key)) if (!fys.includes(r.fy)) seq.push({ label: fyLabel(r.fy), tag: 'published', ...r });
  for (const fy of fys) {
    const t = C.yearTotals(E, fy, th, monthsFor(fy, th), pred).all;
    seq.push({ label: fyLabel(fy), tag: lastFm(fy) < th ? 'through ' + MON[lastFm(fy)] : '', ...t });
  }
  for (const r of seq) h += `<tr><td>${r.label}${r.tag ? `<span class="tag">${r.tag}</span>` : ''}</td><td>${n0(r.courses)}</td><td>${n1(r.perMonth)}</td><td>${n0(r.hours)}</td><td>${n0(r.users)}</td><td>${n2(r.coursesPer)}</td><td>${n2(r.hoursPer)}</td></tr>`;
  if (seq.length > 1) {
    const a = seq[seq.length - 1], b = seq[seq.length - 2];
    h += `<tr class="change"><td>Change, ${b.label} to ${a.label}</td>${['courses', 'perMonth', 'hours', 'users', 'coursesPer', 'hoursPer'].map(k => `<td>${pctTxt(pct(a[k], b[k]))}</td>`).join('')}</tr>`;
  }
  el.innerHTML = h + '</tbody>';
}
function renderHistory() {
  const th = view.through;
  $('histSub').textContent = th === 11
    ? 'FY2023 and FY2024 are the figures published in the FY2025 report. Later years are calculated from the loaded data.'
    : `Counting April through ${MONFULL[th]} of each year. Published FY2023 and FY2024 figures are full-year only, so they’re hidden for this comparison.`;
  yoyTable($('yoyAll'), 'all', null);
  yoyTable($('yoyNoComp'), 'noComp', i => E.cat[i] !== 1);
  const fys = fyList(), full = th === 11, pred = i => E.cat[i] !== 1;
  let h = `<thead><tr><th>Fiscal year</th><th>Courses</th><th>Avg. hours per month</th><th>Hours of training</th><th>Unique users</th><th>Avg. # of courses per person</th><th>Avg. # of hours per person</th></tr></thead><tbody>`;
  const cells = r => `<td>${n0(r.courses)}</td><td>${n1(r.perMonth)}</td><td>${n0(r.hours)}</td><td>${n0(r.users)}</td><td>${n2(r.coursesPer)}</td><td>${n2(r.hoursPer)}</td>`;
  const block = (label, tag, t) => {
    h += `<tr class="l0 group"><td>${label} BUS and BCA${tag ? `<span class="tag">${tag}</span>` : ''}</td>${cells(t.all)}</tr>`;
    h += `<tr class="l1"><td><span class="bus">BUS</span></td>${cells(t.BUS)}</tr><tr class="l1"><td><span class="bca">BCA</span></td>${cells(t.BCA)}</tr>`;
  };
  if (full) for (const r of histRows('units')) if (!fys.includes(r.fy)) block(fyLabel(r.fy), 'published', r);
  for (const fy of fys) block(fyLabel(fy), lastFm(fy) < th ? 'through ' + MON[lastFm(fy)] : '', C.yearTotals(E, fy, th, monthsFor(fy, th), pred));
  $('yoyUnits').innerHTML = h + '</tbody>';
  const p = BUILTIN.published2025;
  $('histFoot').textContent = fys.includes(2025)
    ? `For reference, the published FY2025 report shows ${n0(p.all.courses)} courses and ${n0(p.all.hours)} hours (${n0(p.noComp.hours)} hours without compliance). The Course Completions export only includes people still in the LMS, so calculated FY2025 figures run lower; use the calculated rows when comparing with FY2026 so both years are counted the same way.`
    : '';
}

function renderTop() {
  const fy = view.fy, th = view.through, pred = predicate();
  const map = new Map();
  for (let i = 0; i < E.n; i++) {
    if (!C.inWindow(E.m[i], fy, th) || !pred(i)) continue;
    const c = E.c[i]; let a = map.get(c);
    if (!a) { a = { n: 0, h: 0, users: new Set(), hb: 0, hc: 0, cat: E.cat[i], d: new Set() }; map.set(c, a); }
    a.n++; a.h += E.h[i]; a.users.add(E.p[i]); if (E.u[i]) a.hc += E.h[i]; else a.hb += E.h[i]; a.d.add(E.d[i]);
  }
  const list = [...map.entries()].sort((a, b) => b[1].h - a[1].h).slice(0, 20);
  $('topSub').textContent = `The 20 courses with the most training hours in ${fyLabel(fy)}, using all filters above.`;
  if (!list.length) { $('top').innerHTML = '<tbody><tr><td class="empty">No courses match these filters.</td></tr></tbody>'; return; }
  let h = `<thead><tr><th>Course</th><th>Type</th><th>Completions</th><th>Learners</th><th class="bus">BUS hours</th><th class="bca">BCA hours</th><th>Total hours</th></tr></thead><tbody>`;
  for (const [c, a] of list) {
    const [id, title] = DS.courses[c];
    const type = a.cat === 3 ? [...a.d].map(d => DELNAME[d]).join(' / ') : CATNAME[a.cat];
    h += `<tr><td class="title">${esc(title || id)}<br><span class="dim" style="font-size:13px">${esc(id.replace(/^LI:/, 'LinkedIn '))}</span></td><td>${type}</td><td>${n0(a.n)}</td><td>${n0(a.users.size)}</td><td>${n0(a.hb)}</td><td>${n0(a.hc)}</td><td><strong>${n0(a.h)}</strong></td></tr>`;
  }
  $('top').innerHTML = h + '</tbody>';
}

/* ---------- data panel ---------- */
function range(arr) { let lo = Infinity, hi = -Infinity; for (const v of arr) { if (v < lo) lo = v; if (v > hi) hi = v; } return lo <= hi ? [lo, hi] : null; }
function renderCoverage() {
  const a = range(DS.lms.m), b = range(DS.li.m), items = [];
  items.push(a ? `Course Completions: ${monthName(a[0])} – ${monthName(a[1])}, ${n0(DS.lms.m.length)} completions` : 'Course Completions: none yet');
  items.push(b ? `LinkedIn Learning export: ${monthName(b[0])} – ${monthName(b[1])}, ${n0(DS.li.m.length)} courses. For these months it replaces LinkedIn courses in the LMS file.` : 'LinkedIn Learning export: none. LinkedIn courses from the LMS file count at the default course length.');
  const recent = DS.log.slice(-4).reverse().map(l => `${esc(l.file)} added ${new Date(l.at).toLocaleDateString('en-US')} (${monthName(l.from)} – ${monthName(l.to)})`);
  if (recent.length) items.push('Recent updates: ' + recent.join('; '));
  $('coverage').innerHTML = items.map(x => `<li>${x}</li>`).join('');
  $('defMin').value = settings.defaultMin;
  $('autoComp').checked = settings.autoCompliance !== false;
}

function renderRules() {
  const show = $('ruleShow').value, q = $('ruleSearch').value.trim().toLowerCase();
  const out = [];
  for (let i = 0; i < DS.courses.length; i++) {
    const [id, title] = DS.courses[i];
    if (/^LI:/.test(id)) continue;
    if (show === 'comp' && !CLS.comp[i]) continue;
    if (show === 'ilt' && !courseInfo.instr[i]) continue;
    if (q && !(id.toLowerCase().includes(q) || String(title).toLowerCase().includes(q))) continue;
    out.push(i);
  }
  out.sort((a, b) => courseInfo.count[b] - courseInfo.count[a]);
  const shown = out.slice(0, 150);
  let h = `<thead><tr><th>Course</th><th>Completions loaded</th><th>Compliance</th><th>Department (instructor-led)</th></tr></thead><tbody>`;
  for (const i of shown) {
    const [id, title] = DS.courses[i];
    const deptSel = courseInfo.instr[i] ? `<select data-dept="${esc(id)}">${C.DEPTS.map(d => `<option${CLS.dept[i] === d ? ' selected' : ''}>${d}</option>`).join('')}</select>` : '<span class="dim">–</span>';
    h += `<tr><td class="title">${esc(title || id)}<br><span class="dim" style="font-size:13px">${esc(id)}</span></td><td>${n0(courseInfo.count[i])}</td>
      <td style="text-align:center"><input type="checkbox" data-comp="${esc(id)}"${CLS.comp[i] ? ' checked' : ''} aria-label="Compliance: ${esc(title || id)}"></td><td>${deptSel}</td></tr>`;
  }
  if (!shown.length) h += `<tr><td colspan="4" class="empty">No courses match. Try another search, or choose “All courses”.</td></tr>`;
  $('rules').innerHTML = h + '</tbody>';
  $('rulesFoot').textContent = out.length > shown.length ? `Showing the 150 most-completed of ${n0(out.length)} matching courses. Search to find others.` : `${n0(out.length)} course${out.length === 1 ? '' : 's'}.`;
}
$('ruleShow').onchange = renderRules;
$('ruleSearch').oninput = () => { clearTimeout(renderRules.t); renderRules.t = setTimeout(renderRules, 150); };
$('rules').onchange = e => {
  const t = e.target;
  if (t.dataset.comp) {
    const id = t.dataset.comp;
    settings.compAdd = settings.compAdd.filter(x => x !== id); settings.compRemove = settings.compRemove.filter(x => x !== id);
    const base = builtin.compliance.has(id) || (settings.autoCompliance !== false && C.autoCompliance(id, (DS.courses.find(c => c[0] === id) || [])[1] || ''));
    if (t.checked && !base) settings.compAdd.push(id);
    if (!t.checked && base) settings.compRemove.push(id);
  } else if (t.dataset.dept) {
    settings.dept[t.dataset.dept] = t.value;
  } else return;
  saveSettings(); rebuild(); render();
};
$('defMin').onchange = e => { const v = parseFloat(e.target.value); if (v > 0) { settings.defaultMin = v; saveSettings(); rebuild(); render(); renderRules(); } };
$('autoComp').onchange = e => { settings.autoCompliance = e.target.checked; saveSettings(); rebuild(); render(); renderRules(); };



/* ---------- house style: theme (shared pb-theme convention) ---------- */
(function () {
  if (localStorage.getItem('pb-theme') !== 'dark') document.body.classList.add('light-mode');
  if ($('themeBtn')) $('themeBtn').textContent = document.body.classList.contains('light-mode') ? '\ud83c\udf19 Dark' : '\u2600 Light';
})();
window.toggleTheme = function () {
  const light = document.body.classList.toggle('light-mode');
  localStorage.setItem('pb-theme', light ? 'light' : 'dark');
  if ($('themeBtn')) $('themeBtn').textContent = light ? '\ud83c\udf19 Dark' : '\u2600 Light';
};
if ($('themeBtn')) $('themeBtn').onclick = window.toggleTheme;

/* ---------- house style: home nav (triple-click the title) ---------- */
(function homeTripleClick() {
  const h1 = document.querySelector('h1'); if (!h1) return;
  let n = 0, t = null;
  h1.style.cursor = 'default';
  h1.addEventListener('click', () => {
    n++; clearTimeout(t); t = setTimeout(() => { n = 0; }, 600);
    if (n >= 3) { n = 0; window.location.href = 'index.html'; }
  });
})();

/* ---------- house style: info tooltips ---------- */
document.addEventListener('click', function (e) {
  if (!e.target.classList.contains('info-btn')) { const p = $('info-popover'); if (p) p.classList.remove('visible'); }
});
const INFO_MSGS = {
  'export': 'Download a report based on the current fiscal year, months and filters above. PDF opens a print-ready version of everything on screen (except the Data & calculation rules panel); Excel gives the same numbers as a spreadsheet, across every section.',
  'monthly-chart': 'Hours, completions, or learners for the selected fiscal year, broken out by month and stacked BUS then BCA. The dashed line (when shown) is the prior fiscal year, both units combined, for the same metric.',
  'yoy-filtered': 'Every fiscal year of data loaded, using the unit/delivery/compliance/LinkedIn filters and month window above. Change columns compare each year to the one before it.',
  'report': 'The same breakdown as the published FY L&D Learning Completion report: Online, Compliance, LinkedIn and Instructor-Led (with departments), split BUS/BCA. Always shows every source regardless of the filters above \u2014 only the fiscal year and month window apply.',
  'yoy-history': 'Calculated figures for every year of data loaded, plus FY2023/FY2024 as published in the FY2025 report (full fiscal years only). The BUS vs BCA table excludes compliance courses to match the published comparison.',
  'top-courses': 'The 20 courses with the most training hours in the selected fiscal year, using all filters above.',
};
window.showInfo = function (e, key) {
  const pop = $('info-popover'); if (!pop) return;
  pop.textContent = INFO_MSGS[key] || '';
  pop.classList.add('visible');
  const r = e.target.getBoundingClientRect();
  pop.style.top = (r.bottom + 6) + 'px';
  pop.style.left = Math.min(r.left, window.innerWidth - 300) + 'px';
  e.stopPropagation();
};

/* ---------- house style: export dropdown ---------- */
window.toggleExportDrop = function () { $('export-menu').classList.toggle('open'); };
document.addEventListener('click', function (e) {
  const d = $('export-drop');
  if (d && !d.contains(e.target)) $('export-menu').classList.remove('open');
});

function exportFilterSummary() {
  const parts = [`${fyLabel(view.fy)}, ${windowLabel(view.fy, view.through)}`, filterPhrase()];
  return parts.join(' \u00b7 ');
}

/* ---------- house style: export - PDF (print what's on screen) ---------- */
window.runExportPDF = function () {
  $('export-menu').classList.remove('open');
  const dateStr = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  $('ph-date').textContent = 'Generated: ' + dateStr;
  $('ph-filters').textContent = exportFilterSummary();
  window.print();
};

/* ---------- house style: export - Excel ---------- */
function buildExportSheets() {
  const fy = view.fy, th = view.through, months = monthsFor(fy, th), pred = predicate();
  const sheets = [];

  const cur = sumWindow(fy, Math.min(th, lastFm(fy)), pred);
  sheets.push(['Summary', [
    ['LMS Engagement Report'],
    ['Generated: ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })],
    ['Scope: ' + exportFilterSummary()],
    [],
    ['SUMMARY', fyLabel(fy)],
    ['', 'Courses', 'Hours', 'Learners'],
    ['All', cur.all.courses, Math.round(cur.all.hours), cur.all.users],
    ['BUS', cur.BUS.courses, Math.round(cur.BUS.hours), cur.BUS.users],
    ['BCA', cur.BCA.courses, Math.round(cur.BCA.hours), cur.BCA.users],
  ]]);

  const monthRows = [['Month', 'BUS', 'BCA', 'Prior year (both units)']];
  const curSeries = monthlySeries(fy, pred), lf = lastFm(fy);
  const hasPrev = fyList().includes(fy - 1), prevSeries = hasPrev ? monthlySeries(fy - 1, pred) : null;
  curSeries.forEach((r, i) => {
    if (i > lf) return;
    monthRows.push([MONFULL[i], Math.round(metricOf(r[0])), Math.round(metricOf(r[1])), prevSeries ? Math.round(metricOf(prevSeries[i][2])) : '']);
  });
  sheets.push(['Monthly (' + (view.metric === 'hours' ? 'Hours' : view.metric === 'courses' ? 'Completions' : 'Learners') + ')', monthRows]);

  const yoyRows = [['Fiscal Year', 'BUS Hours', 'BCA Hours', 'Total Hours', 'Completions', 'Learners', 'Hours / Learner']];
  let prevY = null;
  fyList().forEach(f => {
    const s = sumWindow(f, th, pred);
    yoyRows.push([fyLabel(f), Math.round(s.BUS.hours), Math.round(s.BCA.hours), Math.round(s.all.hours), s.all.courses, s.all.users, +s.all.hoursPer.toFixed(2)]);
    prevY = s;
  });
  sheets.push(['Your Filtered View YoY', yoyRows]);

  const reportRows = [['Learning Source', '# of Courses', 'Hours', 'Unique Users', 'Avg Hours / Month', 'Avg Courses / Person', 'Avg Hours / Person']];
  C.reportTable(E, CLS, fy, th, months).forEach(r => {
    const label = r.level === 0 ? r.label : r.unit || r.dept || r.label;
    reportRows.push(['  '.repeat(r.level) + label, r.courses, Math.round(r.hours), r.users, +r.perMonth.toFixed(1), +r.coursesPer.toFixed(2), +r.hoursPer.toFixed(2)]);
  });
  sheets.push([fyLabel(fy) + ' Learning Completion Summary', reportRows]);

  function historySheet(name, predH) {
    const fys = fyList(), full = th === 11;
    const rows = [['Fiscal Year', 'Courses', 'Avg Hours / Month', 'Hours', 'Unique Users', 'Avg Courses / Person', 'Avg Hours / Person']];
    if (full) histRows(name).forEach(r => { if (!fys.includes(r.fy)) rows.push([fyLabel(r.fy) + ' (published)', r.courses, +r.perMonth.toFixed(1), Math.round(r.hours), r.users, +r.coursesPer.toFixed(2), +r.hoursPer.toFixed(2)]); });
    fys.forEach(f => { const t = C.yearTotals(E, f, th, monthsFor(f, th), predH).all; rows.push([fyLabel(f), t.courses, +t.perMonth.toFixed(1), Math.round(t.hours), t.users, +t.coursesPer.toFixed(2), +t.hoursPer.toFixed(2)]); });
    return rows;
  }
  sheets.push(['YoY All Learning', historySheet('all', null)]);
  sheets.push(['YoY Without Compliance', historySheet('noComp', i => E.cat[i] !== 1)]);

  const topMap = new Map();
  for (let i = 0; i < E.n; i++) {
    if (!C.inWindow(E.m[i], fy, th) || !pred(i)) continue;
    const c = E.c[i]; let a = topMap.get(c);
    if (!a) { a = { n: 0, h: 0, users: new Set(), hb: 0, hc: 0, cat: E.cat[i] }; topMap.set(c, a); }
    a.n++; a.h += E.h[i]; a.users.add(E.p[i]); if (E.u[i]) a.hc += E.h[i]; else a.hb += E.h[i];
  }
  const topRows = [['Course', 'Type', 'Completions', 'Learners', 'BUS Hours', 'BCA Hours', 'Total Hours']];
  [...topMap.entries()].sort((a, b) => b[1].h - a[1].h).slice(0, 20).forEach(([c, a]) => {
    const [id, title] = DS.courses[c];
    topRows.push([title || id, CATNAME[a.cat], a.n, a.users.size, Math.round(a.hb), Math.round(a.hc), Math.round(a.h)]);
  });
  sheets.push(['Top Courses', topRows]);

  return sheets;
}
window.runExportXLSX = function () {
  $('export-menu').classList.remove('open');
  const wb = XLSX.utils.book_new();
  buildExportSheets().forEach(([name, rows]) => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = rows[0].map(() => ({ wch: 20 }));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, 'lms-engagement-report.xlsx');
};

(function init() {
  rebuild(); buildSegs(); buildPeriod(); render(); renderRules();
})();

