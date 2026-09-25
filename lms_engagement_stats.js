/*
 * Computes headline stats for the LMS Engagement homepage card, using
 * Resmie's own corelib functions (classify/effective/reportTable) so the
 * numbers always match what the dashboard itself shows for the latest fiscal
 * year through the last month of real data - no month/fiscal-year math is
 * reimplemented here.
 *
 * Usage: node lms_engagement_stats.js <dataset.json> <builtin-reference.json>
 * Prints {lastMonthLabel, fy, learners, hours, courses} as JSON on stdout.
 */
const fs = require('fs');
require('./js/lms-engagement-core.js');
const C = global.LDCore;

const [, , dsPath, builtinPath] = process.argv;
const ds = JSON.parse(fs.readFileSync(dsPath, 'utf-8'));
const builtinRaw = JSON.parse(fs.readFileSync(builtinPath, 'utf-8'));
const builtin = { compliance: new Set(builtinRaw.compliance), dept: builtinRaw.dept };
const settings = { defaultMin: 30, autoCompliance: true, compAdd: [], compRemove: [], dept: {} };

const cls = C.classify(ds, settings, builtin);
const eff = C.effective(ds, cls, settings.defaultMin / 60);

let lastMi = -1;
for (let i = 0; i < eff.n; i++) if (eff.m[i] > lastMi) lastMi = eff.m[i];

const MONTHFULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const monthName = m => MONTHFULL[C.miMonth(m) - 1] + ' ' + C.miYear(m);

const fy = C.fyOf(lastMi);
const through = C.fmOf(lastMi);
const months = through + 1;
const totals = C.yearTotals(eff, fy, through, months, null).all;

process.stdout.write(JSON.stringify({
  lastMonthLabel: monthName(lastMi),
  fy,
  learners: totals.users,
  hours: Math.round(totals.hours),
  courses: totals.courses,
}));
