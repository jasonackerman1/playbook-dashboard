"""
LMS Engagement Dashboard generator.

Built 2026-09-24/25 from a real dashboard Resmie delivered (originally named
LD-Engagement-Dashboard.html). This is a SEPARATE dashboard from the existing
Learning Engagement dashboard (learning-engagement.html /
update_learning_engagement_dashboard.py) - NOT a replacement of it. The two
cover genuinely different analyses: Learning Engagement is LinkedIn Learning
content-mix/reach (still placeholder data as of this build, pending Resmie's
real export in that format); this one is a fiscal-year learning-completion
summary matching the published FY2025 L&D report (Online/Compliance/LinkedIn/
Instructor-Led split, BUS vs BCA, year-over-year comparisons, top courses),
built from real Course Completions + LinkedIn Learning exports. First build
briefly replaced the Learning Engagement dashboard in place; Jason corrected
that same day - keep them separate going forward, including their homepage
cards, data folders, and file names. Originally shipped under the name
"Course Completions"; renamed to "LMS Engagement" 2026-09-25 per Jason,
including all file/script/data-folder names - see the 2026-09-25 changes
note further down.

ARCHITECTURE, why it looks different from every other dashboard here:
Resmie's file does its own parsing, classification (compliance-course
detection, instructor-led department assignment, LinkedIn precedence) and all
report/chart math in a single JS module (js/lms-engagement-core.js, copied
verbatim from her file, unmodified - see its own header comment: "Runs in
browser, worker and Node"). Rather than reimplementing any of those rules in
Python (real risk of subtly mismatching her business logic - compliance
keyword list, fiscal-year definitions, BCA detection, etc.), this script
SHELLS OUT TO NODE (lms_engagement_build.js) to run her exact code against
the raw source files in lms-engagement-data/, and only takes over once the
data itself is ready: embedding it in the shared house-style shell (KM logo,
triple-click-to-home, theme toggle wired to the shared pb-theme key, a
Print/Save-PDF button) and writing lms-engagement.html. Everything else - the
filter bar, charts, fiscal-year tables, Course/Calculation rules override
panel - is Resmie's own HTML/CSS/JS (templates/lms-engagement-shell.html +
js/lms-engagement-app.js), lightly trimmed (see below), not reimplemented.

PIPELINE: node lms_engagement_build.js <data-dir> processes every file in
lms-engagement-data/ in order (any *baseline*.json first, then *.xlsx/*.xls/
*.csv alphabetically - the MM.DD.YYYY filename convention used everywhere
else in this repo gives chronological order) and merges them with Resmie's
own merge() function (later files replace only the months they cover, so
monthly drops never double-count). Verified byte-for-byte against her own
processing log: her real Course-Completions-09.24.2026.xlsx reproduces her
exact "235,358 read -> 133,426 kept" breakdown and the exact FY2025
calculated totals in her table (Online 51,548/25,404hrs/4,119 users;
Compliance 41,103/17,236/4,939; LinkedIn 8,122/8,138/1,224; Instructor Led
2,972/22,753/1,602; Total 103,745/73,531/5,207).

linkedin-baseline.json: we don't have the raw LinkedIn Learning source
workbook (FY2025_L_D_Completion_Report.xlsx), only Resmie's already-merged
result for it. Extracted from her file's baked-in dataset once and stored
here as pre-parsed {kind:"linkedin", recs:[...]} records so
lms_engagement_build.js can merge it like any other source. Once a real
monthly LinkedIn export exists, just drop the raw .xlsx in this folder like
the LMS file - it'll parse and merge over the matching months normally.

builtin-reference.json: static reference/config data, not completion records
- the compliance course-ID list, instructor-led department overrides, and the
published FY2023/FY2024/FY2025 report figures Resmie hardcoded for
year-over-year comparison. Never treated as a data input by
lms_engagement_build.js.

CHANGES FROM RESMIE'S FILE (confirmed with Jason before building):
- The in-browser "upload a new file" feature (drag-drop, Web Worker Excel
  parsing, IndexedDB storage, "Save dashboard with this data" / "Remove
  updates saved in this browser" buttons) is REMOVED. One update path only:
  drop a file in lms-engagement-data/, push, GitHub Actions rebuilds -
  matching every other dashboard in this repo. The "Course rules" and
  "Calculation rules" panels are KEPT (they only tweak local display/
  classification via a localStorage setting, never touch the underlying data)
  since Jason was only asked about the upload path specifically.
- Her embedded xlsxlib (SheetJS, ~250KB) is dropped from the generated page -
  no longer needed once nothing parses Excel in the browser.
- Added: KM Academy logo, triple-click-the-title nav to index.html, a theme
  toggle wired to the shared pb-theme localStorage key (her CSS already had a
  data-theme escape hatch, just never wired to a control), and a plain
  Print/Save-PDF button (window.print(), with print CSS hiding the filter bar
  and data panel).
- Title/h1 changed from her original "Learning engagement, BUS and BCA" to
  "LMS engagement dashboard, BUS and BCA" (renamed 2026-09-25 - see below;
  briefly "Course completions, BUS and BCA" in between) so it's visually
  distinct from the separate Learning Engagement dashboard on the
  homepage/nav.

CHANGES 2026-09-25 (visual pass + rename, per Jason):
- Restyled to match the shared house look exactly: CSS variables/fonts/
  spacing pulled from cert-healthcare.html and update_onboarding_dashboard.py
  (not approximated), full-width layout (no centered/max-width column - main
  and every section span the true viewport width, matching every other
  dashboard), filter bar matches the reference pixel-for-pixel (flat
  var(--surface) bar, inline uppercase labels, same select/button-group
  styling). Added real "?" tooltips and a real Export dropdown (PDF prints
  the on-screen report as-is; Excel writes a 7-sheet workbook built from the
  same computed numbers).
- All 6 filter-bar controls (Fiscal year, Months, Business unit, Delivery,
  Compliance, LinkedIn Learning) are plain <select> dropdowns - Resmie's
  original segmented button-group widget is kept only for the in-panel
  Hours/Completions/Learners metric toggle, which isn't part of the filter
  bar. buildSegs() in js/lms-engagement-app.js handles both `.seg[data-key]`
  (button group) and `select[data-key]` (dropdown) generically off the same
  SEGS data, so no data/interaction logic changed - only which widget renders
  each filter.
- Fixed a real full-bleed bug: .filters is a child of <main>, which has its
  own 28px side padding, so the filter bar's background was inset an extra
  28px on both sides versus .header (which sits outside <main>). Fixed with
  a canceling negative margin on .filters so its background reaches the true
  viewport edge on every row, matching .header exactly.
- Header status line combined: was "Data through Sept 2026. N employee
  completions loaded, last updated Sept 25, 2026" (two different dates in one
  line); now "Data through September 25, 2026. N employee completions
  loaded." - uses the pipeline's last-updated date as the single "Data
  through" date, and drops the separate "last updated" clause entirely.
- Renamed from "Course Completions" to "LMS Engagement" throughout (title,
  h1, homepage card). File/script/data-folder names below were renamed to
  match (course-completions* -> lms-engagement*); only the internal LMS
  export-format name "Course Completions" (the actual report Resmie's
  pipeline detects by sheet headers) was left alone, since that refers to a
  real, fixed data-source name, not the dashboard's identity.

KNOWN GAPS, flagged to Jason rather than silently built around:
- No Hide/Show TLG toggle. Resmie's data model deliberately never stores
  names or emails, only anonymized user IDs (her own doc: "Names, emails, job
  titles, supervisors and locations are not stored") - there is no field left
  to match the TLG list against.
- No custom PDF/Excel export matching the other dashboards' per-report-type
  export (Full Report / Not Certified / Manager Summary style). Given the
  scope of this build, only a plain browser Print/Save-PDF was added - worth
  a real export build later if wanted.
- The homepage card shows unique learners / courses / hours / data-through
  date for the latest fiscal year - there's no company headcount figure in
  this data model to compute a "% reached" stat from.
"""
import json
import os
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'lms-engagement-data')
CORE_JS = os.path.join(SCRIPT_DIR, 'js', 'lms-engagement-core.js')
APP_JS = os.path.join(SCRIPT_DIR, 'js', 'lms-engagement-app.js')
SHELL_HTML = os.path.join(SCRIPT_DIR, 'templates', 'lms-engagement-shell.html')
BUILD_JS = os.path.join(SCRIPT_DIR, 'lms_engagement_build.js')
STATS_JS = os.path.join(SCRIPT_DIR, 'lms_engagement_stats.js')
BUILTIN_JSON = os.path.join(DATA_DIR, 'builtin-reference.json')
OUTPUT_HTML = os.path.join(SCRIPT_DIR, 'lms-engagement.html')
TMP_DS_CACHE = os.path.join(SCRIPT_DIR, '.lms-engagement-dataset-cache.json')


def run_pipeline():
    """Runs Resmie's own parsing/classification/merge code (via Node) against
    every file in lms-engagement-data/. Returns the merged dataset."""
    result = subprocess.run(
        ['node', BUILD_JS, DATA_DIR],
        cwd=SCRIPT_DIR, capture_output=True, text=True,
    )
    for line in result.stderr.splitlines():
        print(f"    {line}")
    if result.returncode != 0:
        raise RuntimeError(f"lms_engagement_build.js failed (exit {result.returncode}). See output above.")
    return json.loads(result.stdout)


def compute_stats(ds):
    """Headline numbers for the console + homepage card, computed via
    Resmie's own classify/effective/yearTotals functions (Node) so they can
    never drift from what the dashboard itself shows."""
    with open(TMP_DS_CACHE, 'w', encoding='utf-8') as f:
        json.dump(ds, f)
    try:
        result = subprocess.run(
            ['node', STATS_JS, TMP_DS_CACHE, BUILTIN_JSON],
            cwd=SCRIPT_DIR, capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(f"    lms_engagement_stats.js failed: {result.stderr}")
            return None
        return json.loads(result.stdout)
    finally:
        if os.path.exists(TMP_DS_CACHE):
            os.remove(TMP_DS_CACHE)


def generate_html(ds, builtin, settings):
    shell = open(SHELL_HTML, encoding='utf-8').read()
    core_js = open(CORE_JS, encoding='utf-8').read()
    app_js = open(APP_JS, encoding='utf-8').read()

    def script_block(id_, content, is_json=False):
        type_attr = ' type="application/json"' if is_json else ''
        return f'<script id="{id_}"{type_attr}>{content}</script>'

    scripts = '\n'.join([
        script_block('corelib', core_js),
        script_block('builtin', json.dumps(builtin, separators=(',', ':')), is_json=True),
        script_block('settings', json.dumps(settings, separators=(',', ':')), is_json=True),
        script_block('dataset', json.dumps(ds, separators=(',', ':')), is_json=True),
        f'<script>{app_js}</script>',
    ])
    return shell + '\n' + scripts + '\n</body>\n</html>\n'


def build():
    """Runs the full pipeline and writes lms-engagement.html. Returns the
    stats dict (see compute_stats) for callers that also need headline
    numbers, e.g. generate_homepage.py."""
    print("  Running LMS Engagement data pipeline (Node, Resmie's own parsing code)...")
    ds = run_pipeline()
    with open(BUILTIN_JSON, encoding='utf-8') as f:
        builtin = json.load(f)
    settings = {}  # defaults + any per-course overrides are applied client-side from localStorage

    html = generate_html(ds, builtin, settings)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)

    stats = compute_stats(ds)
    print(f"    {len(ds['users'])} users, {len(ds['courses'])} courses, "
          f"{len(ds['lms']['m'])} LMS completions, {len(ds['li']['m'])} LinkedIn records")
    if stats:
        print(f"    Data through {stats['lastMonthLabel']} | FY{stats['fy']}: "
              f"{stats['learners']} learners, {stats['hours']} hours, {stats['courses']} courses")
    print(f"  Written -> {OUTPUT_HTML}")
    return stats


def load_lms_engagement_stats():
    """Used by generate_homepage.py for the LMS Engagement card. Re-runs
    the pipeline rather than caching, so the homepage card can never show
    numbers that disagree with the dashboard itself."""
    ds = run_pipeline()
    return compute_stats(ds)


def main():
    build()


if __name__ == '__main__':
    main()
