/*
 * LMS Engagement dashboard data pipeline.
 *
 * Runs Resmie's own parsing/classification/merge code (js/lms-engagement-core.js,
 * verbatim from her LD-Engagement-Dashboard.html build) against every raw
 * source file in the data folder, and prints the resulting merged dataset as
 * JSON on stdout. update_learning_engagement_dashboard.py shells out to this
 * script rather than reimplementing any of her filtering/classification rules
 * in Python.
 *
 * Usage: node lms_engagement_build.js <data-dir>
 *
 * Input files in <data-dir>, processed in this order:
 *   1. *baseline*.json   - pre-parsed records ({kind, recs}), for source data
 *                          we don't have the raw workbook for (see
 *                          linkedin-baseline.json).
 *   2. everything else (*.xlsx, *.xls, *.csv), sorted alphabetically by
 *      filename - relies on the MM.DD.YYYY filename convention used
 *      everywhere else in this repo for chronological order.
 * builtin-reference.json is reference/config data (compliance list, dept
 * overrides, published-report history), not completion records, so it is
 * never read as a data input here - update_learning_engagement_dashboard.py
 * embeds it directly.
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

require('./js/lms-engagement-core.js'); // attaches globalThis.LDCore
const C = global.LDCore;

const dataDir = process.argv[2];
if (!dataDir) {
  console.error('Usage: node lms_engagement_build.js <data-dir>');
  process.exit(1);
}

const allFiles = fs.readdirSync(dataDir).filter(f => !f.startsWith('.'));
const baselineFiles = allFiles.filter(f => /baseline.*\.json$/i.test(f)).sort();
const rawFiles = allFiles
  .filter(f => /\.(xlsx|xls|csv)$/i.test(f))
  .sort();
const skipped = allFiles.filter(f => !baselineFiles.includes(f) && !rawFiles.includes(f));

const ds = C.emptyDS();
const report = [];

for (const f of baselineFiles) {
  const full = path.join(dataDir, f);
  const parsed = JSON.parse(fs.readFileSync(full, 'utf-8'));
  if (!parsed || !parsed.kind || !Array.isArray(parsed.recs)) {
    console.error(`Skipping ${f}: not a recognised baseline file (needs {kind, recs}).`);
    continue;
  }
  const { replaced } = C.merge(ds, parsed, parsed.sourceFile || f);
  report.push({ file: f, kind: parsed.kind, rows: parsed.recs.length, from: replaced[0], to: replaced[1] });
}

for (const f of rawFiles) {
  const full = path.join(dataDir, f);
  const buf = fs.readFileSync(full);
  const uint8 = new Uint8Array(buf);
  const results = C.readBook(XLSX, uint8, () => {});
  if (!results.length) {
    console.error(`WARNING: ${f} has no recognisable Course Completions or LinkedIn Learning sheet - skipped.`);
    continue;
  }
  for (const r of results) {
    if (!r.parsed || !r.parsed.recs.length) {
      console.error(`WARNING: ${f} sheet "${r.sheet}" produced zero usable rows.`);
      continue;
    }
    const { replaced } = C.merge(ds, r.parsed, f);
    report.push({ file: f, sheet: r.sheet, kind: r.parsed.kind, rows: r.parsed.recs.length, stats: r.parsed.st, from: replaced[0], to: replaced[1] });
  }
}

if (skipped.length) console.error('Ignored (not a data input): ' + skipped.join(', '));

console.error('--- LMS Engagement data pipeline ---');
for (const r of report) {
  console.error(`${r.file}${r.sheet ? ' [' + r.sheet + ']' : ''}: ${r.kind} - ${r.rows} rows` + (r.stats ? ` (kept ${r.stats.kept} of ${r.stats.read} read)` : ''));
}
console.error(`Final dataset: ${ds.users.length} users, ${ds.courses.length} courses, ${ds.lms.m.length} LMS completions, ${ds.li.m.length} LinkedIn records.`);

process.stdout.write(JSON.stringify(ds));
